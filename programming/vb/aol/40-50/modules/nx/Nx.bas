Attribute VB_Name = "Nx0"
'--
'
'--
'
'--Nx.bas
'--made in vb6
'--by Nx
'
'--module stats
'total procedures: 282
'AIM Related: 45
'AOL Related: 52
'IE Related: 3
'mIRC Related: 7
'WinAMP Related: 62
'Windows Related/Misc: 113
'
'--from the maker
'hope you enjoy, if you have any
'complaints/questions/comments e-mail me with them
'and i'll get to them as soon as i can
'
'--props/greets
'dos, hider, ecco, sonic, hix, knk, pat or jk,
'cheetah, chip, kid, magus, Steve J. Gray,
'Compulsion Software
'perolta, scooby, cracklyn, flex, unity,
'Amanda, heckyl, clone, guillo, shock programming
'
'--pages
'dos:
'   http://www.dosfx.com
'KnK:
'   http://www.knk2000.com/knk
'PAT or JK:
'   http://www.patorjk.com
'cheetah:
'   http://www.seratsuki.com
'EasyVB
'   http://www.easyvb.com
'Plastik Designs
'   http://www.dosfx.com/~plastik
'Nx:
'   http://scnx.cjb.net
'
'--questions/comments
'   mail me:
'       faded_concept@hotmail.com
'
'   AIM:
'       nxer
'
'   ©ZeroProductions 2000
'   ™ Nx 1999-2000
'--
'
'--

Option Explicit


'--declares
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FileTime, lpLastAccessTime As FileTime, lpLastWriteTime As FileTime) As Long
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FileTime, lpLocalFileTime As FileTime) As Long
Public Declare Function GetFileTitle Lib "comdlg32.dll" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Public Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Public Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Public Declare Function SHPathIsExe Lib "shell32" Alias "#43" (ByVal szPath As String) As Long
Public Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function SHRunDialog Lib "shell32" Alias "#61" (ByVal hOwner As Long, ByVal Unknown1 As Long, ByVal Unknown2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Public Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Public Declare Function destroywindow Lib "user32" Alias "DestroyWindow" (ByVal hwnd As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndparent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Byte) As Long
Public Declare Function FindText Lib "comdlg32.dll" Alias "FindTextA " (pFindreplace As String) As Long
Public Declare Function AbortSystemShutdown Lib "advapi32.dll" Alias "AbortSystemShutdownA" (ByVal lpMachineName As String) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessagebyString& Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, wMsg As Long, ByVal wParam As Long, ByVal lParam As String)
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageCDS Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As COPYDATASTRUCT) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function FlashWindow& Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long)
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerLanguageName Lib "VERSION.DLL" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
Public Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (pBlock As Byte, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function GetVersionOS Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Sub VBGetTarget Lib "kernel32" Alias "RtlMoveMemory" (Target As Any, ByVal lPointer As Long, ByVal cbCopy As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Sub FatalAppExit Lib "kernel32" Alias "FatalAppExitA" (ByVal uAction As Long, ByVal lpMessageText As String)
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


'--constants
Public Const EXIT_LOGOFF = 0
Public Const EXIT_SHUTDOWN = 1
Public Const EXIT_REBOOT = 2

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const EWX_POWEROFF = 8

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Const GWL_WNDPROC = (-4)

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const CD_ADDITEM = &H143
Public Const CB_INSERTSTRING = &H14A
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETCOUNT = &H146

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_FINDSTRING = &H18F
Public Const LB_SETHORIZONTALEXTENT = &H194

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_RESTORE = 9
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_USER = &H400
Public Const WM_WA_IPC = WM_USER
Public Const WM_COPYDATA = &H4A
Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_Gettext = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONCLICK = &H202
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT& = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_NCLBUTTONDOWN = &HA1

Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000
Public Const OFN_EXPLORER = &H80000
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0

Public Const HTCAPTION = 2

Public Const MF_String = &H0&
Public Const MF_Separator = &H800&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const shrdNoMRUString = &H2
Public Const WS_CHILD = &H40000000

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

Public Const BIF_RETURNONLYFSDIRS = &H1

Private Const MAX_PATH = 260
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000
Private Const SHGFI_LARGEICON = &H0
Private Const SHGFI_SMALLICON = &H1
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE
Private Const BOLD_FONTTYPE = &H100
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_APPLY = &H200&
Private Const CF_SCREENFONTS = &H1
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_EFFECTS = &H100&
Private Const CF_PALETTE = 9
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FO_DELETE = &H3

Public Const EM_LIMITTEXT = &HC5

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const KEY_QUERY_VALUE = &H1

Public Const vbKeyShift = &H10
Public Const vbKeyCtrl = &H11
Public Const vbKeyAlt = &H12

Public Const IPC_DELETE = 101
Public Const IPC_ISPLAYING = 104
Public Const IPC_GETOUTPUTTIME = 105
Public Const IPC_JUMPTOTIME = 106
Public Const IPC_WRITEPLAYLIST = 120
Public Const IPC_SETPLAYLISTPOS = 121
Public Const IPC_SETVOLUME = 122
Public Const IPC_SETPANNING = 123
Public Const IPC_GETLISTLENGTH = 124
Public Const IPC_SETSKIN = 200
Public Const IPC_GETSKIN = 201
Public Const IPC_GETLISTPOS = 125
Public Const IPC_GETINFO = 126
Public Const IPC_GETEQDATA = 127
Public Const IPC_PLAYFILE = 100
Public Const IPC_CHDIR = 103

Public Const WINAMP_REG_KEY = "WinAmp.File\shell\play\command"
Public Const WINAMP_OPTIONS_EQ = 40036
Public Const WINAMP_OPTIONS_PLEDIT = 40040
Public Const WINAMP_VOLUMEUPS = 40058
Public Const WINAMP_VOLUMEDOWNS = 40059
Public Const WINAMP_FFWD5S = 40060
Public Const WINAMP_REW5S = 40061
Public Const WINAMP_BUTTON1 = 40044
Public Const WINAMP_BUTTON2 = 40045
Public Const WINAMP_BUTTON3 = 40046
Public Const WINAMP_BUTTON4 = 40047
Public Const WINAMP_BUTTON5 = 40048
Public Const WINAMP_BUTTON1_SHIFT = 40144
Public Const WINAMP_BUTTON4_SHIFT = 40147
Public Const WINAMP_BUTTON5_SHIFT = 40148
Public Const WINAMP_BUTTON1_CTRL = 40154
Public Const WINAMP_BUTTON2_CTRL = 40155
Public Const WINAMP_BUTTON5_CTRL = 40158
Public Const WINAMP_FILE_PLAY = 40029
Public Const WINAMP_OPTIONS_PREFS = 40012
Public Const WINAMP_OPTIONS_AOT = 40019
Public Const WINAMP_HELP_ABOUT = 40041

Public Const KEYEVENTF_KEYUP = &H2

Public Const LANG_ENGLISH = &H409
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2


'--types
Public Type Mp3InfoType
        Title As String
        Artist As String
        Album As String
        Year As Integer
        Genre As String
        Comment As String
End Type
Public Type COPYDATASTRUCT
        dwData As Long
        cbData As Long
        lpData As String
End Type
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Type VS_FIXEDFILEINFO
   dwSignature As Long
   dwStrucVersion As Long
   dwFileVersionMS As Long
   dwFileVersionLS As Long
   dwProductVersionMS As Long
   dwProductVersionLS As Long
   dwFileFlagsMask As Long
   dwFileFlags As Long
   dwFileOS As Long
   dwFileType As Long
   dwFileSubtype As Long
   dwFileDateMS As Long
   dwFileDateLS As Long
End Type
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Public Type Info
    sTitle As String * 30
    sArtist As String * 30
    sAlbum As String * 30
    sComment As String * 30
    sYear As String * 4
    sGenre As String * 21
End Type
Public Type NxUPLOADSTAT
    UL_PERDONE As String
    UL_MINLEFT As String
    UL_FILENAME As String
End Type
Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Type SHITEMID
  cb      As Long
  abID    As Byte
End Type
Public Type ITEMIDLIST
  mkid    As SHITEMID
End Type
Public Type BROWSEINFO
  hOwner          As Long
  pidlRoot        As Long
  pszDisplayName  As String
  lpszTitle       As String
  ulFlags         As Long
  lpfn            As Long
  lParam          As Long
  iImage          As Long
End Type
Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * 260
    szTypeName As String * 80
End Type
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Type FileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Public Type ChooseFont
        lStructSize As Long
        hwndOwner As Long
        hdc As Long
        lpLogFont As Long
        iPointSize As Long
        Flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
        lpszStyle As String
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
        nSizeMax As Long
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
        lfFaceName(LF_FACESIZE) As Byte
End Type
Public Type ChooseColor
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type IDTag
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 28
    Genre As String * 1
End Type


'--definitions
Private vbTray As NOTIFYICONDATA
Private shinfo As SHFILEINFO
Private Mp3Info As Info
Private sFilename As String
Public hwndOwner As Long
Private lOldProc&, lpMT$
Public IsAbout As Boolean
Public Filter As String
Public OpenDialogTitle As String
Public SaveDialogTitle As String
Public FolderDialogTitle As String
Public ShowDirsOnly As Boolean
Public hwnd_winamp As Long
Public WMp3Info As Mp3InfoType
Global glrInt As Integer
Global StopIt As Boolean

Public Sub lastboot(Dest As Label)
 Dim lngHours As Long, lngMinutes As Long, lngcount As Long
 lngcount = GetTickCount
 lngHours = ((lngcount / 1000) / 60) / 60
 lngMinutes = ((lngcount / 1000) / 60) Mod 60
 Dest.Caption = lngHours & "hrs. " & lngMinutes & "mins."
End Sub

Public Property Get MpFilename() As String
'--retrieves the Filename from the file set
'--by the MpFilename property let
'--ex:
'Nx0.MpFilename = "C:\nx.mp3"
'Text1.Text = Nx0.MpFilename
'
    MpFilename = sFilename
End Property
Public Property Let MpFilename(ByVal sPassFilename As String)
'--sets the file that all id3 info will be retrieved from
'--ex:
'Nx0.MpFilename = "C:\nx.mp3"
'
    Dim iFreefile As Integer
    Dim lFilePos As Long
    Dim sData As String * 128
    Dim sGenreMatrix As String
    Dim sGenre() As String
    
    ' Genre
    
    sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
        "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
        "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
        "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
        "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
        "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
        "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
        "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
        "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
        "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
        "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
        "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
        "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
        "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
        "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
        "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
        "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
        
    ' Build the Genre array (VB6+ only)
    
    sGenre = Split(sGenreMatrix, "|")
    
    ' Store the filename (for "Get Filename" property)

    sFilename = sPassFilename
    
    ' Clear the info variables
    
    Mp3Info.sTitle = ""
    Mp3Info.sArtist = ""
    Mp3Info.sAlbum = ""
    Mp3Info.sYear = ""
    Mp3Info.sComment = ""
    
    ' Ensure the MP3 file exists
    
    If Dir(sFilename) = "" Then Exit Property
    ' Retrieve the info data from the MP3
    
    iFreefile = FreeFile
    lFilePos = FileLen(sFilename) - 127
    Open sFilename For Binary As #iFreefile
        Get #iFreefile, lFilePos, sData
    Close #iFreefile
    
    ' Populate the info variables
    
    If Left(sData, 3) = "TAG" Then
        Mp3Info.sTitle = Mid(sData, 4, 30)
        Mp3Info.sArtist = Mid(sData, 34, 30)
        Mp3Info.sAlbum = Mid(sData, 64, 30)
        Mp3Info.sYear = Mid(sData, 94, 4)
        Mp3Info.sComment = Mid(sData, 98, 28)
        '----------------------------------------------------------
        '--returns "subscript out of range"
        'MP3Info.sGenre = sGenre(Asc(Mid(sData, 128, 1)))
        '---------------------------------------------------------
    End If
End Property
Public Property Get MpTitle() As String
'--retrieves the Title from the file set
'--by the MpFilename property let
'--ex:
'Nx0.MpFilename = "C:\nx.mp3"
'Text1.Text = Nx0.MpTitle

    MpTitle = RTrim(Mp3Info.sTitle)
End Property
Public Property Get MpArtist() As String
'--retrieves the Artist from the file set
'--by the MpFilename property let
'--ex:
'Nx0.MpFilename = "C:\nx.mp3"
'Text1.Text = Nx0.MpArtist

    MpArtist = RTrim(Mp3Info.sArtist)
End Property
Public Property Get MpGenre() As String
'--will work in the next release
    MpGenre = RTrim(Mp3Info.sGenre)
End Property
Public Property Get MpAlbum() As String
'--retrieves the Album from the file set
'--by the MpFilename property let
'--ex:
'Nx0.MpFilename = "C:\nx.mp3"
'Text1.Text = Nx0.MpAlbum

    MpAlbum = RTrim(Mp3Info.sAlbum)
End Property

Public Property Get MpYear() As String
'--retrieves the Year from the file set
'--by the MpFilename property let
'--ex:
'Nx0.MpFilename = "C:\nx.mp3"
'Text1.Text = Nx0.MpYear

    MpYear = Mp3Info.sYear
End Property
Public Property Get MpComment() As String
'--retrieves the Comment from the file set
'--by the MpFilename property let
'--ex:
'Nx0.MpFilename = "C:\nx.mp3"
'Text1.Text = Nx0.MpComment

    MpComment = RTrim(Mp3Info.sComment)
End Property

Sub SaveIDTag(mFilename As String, FTitle As String, FArtist As String, FAlbum As String, FYear As String, FComment As String)  'FTrack As String,

Dim F As String, FIO As Integer, n As Long, tagpos As Long
Dim p As Long, M As String, TPos As Double
Dim NewTag As IDTag, Inbuf As String * 256
    
    F = mFilename
    FIO = FreeFile
    Open F For Binary As FIO
    n = LOF(FIO): If n < 256 Then Close: Exit Sub
    Get #FIO, (n - 255), Inbuf
    p = InStr(1, Inbuf, "tag", 1)
    If p = 0 Then
        tagpos = n + 1: M = "Added!"
    Else
        M = "Updated!"
        tagpos = n - 256 + p
    End If
    Close FIO
    
    With NewTag
        .Title = FTitle
        .Artist = FArtist
        .Album = FAlbum
        .Year = FYear
        .Comment = FComment
'------------------------------------------------
'originally designed to work w/ a combobox but
'since it got module-ified it doesn't work, sorry

        'n = cboGenre.ListIndex - 1: If n < 0 Then n = 254
        '.Genre = Chr(12)
    
'------------------------------------------------
    End With
    FIO = FreeFile
    Open F For Binary As FIO
    Put #FIO, tagpos, "TAG"
    Put #FIO, tagpos + 3, NewTag
    Close FIO
    MsgBox "Tag " & M, vbExclamation, "Nx id3 editor"
    Exit Sub
End Sub

Function decryptstring(txt As String, key As String) As String
Dim l0062 As Variant
Dim l0066 As Variant
Dim l006A As Variant
Dim l006E As Variant
Dim l0072 As Variant
Dim l0076 As Variant
Dim l007A As Variant
Dim l007E As Variant
Dim l0082 As Variant
Dim l0086 As Variant
Dim l008A As Variant
Dim l008E As Variant
On Error Resume Next
If txt = "" Then
MsgBox "please select a string to decrypt", vbCritical & vbOKOnly, "string decryption"
Exit Function
End If
If key = "" Then
MsgBox "please select an decryption key", vbCritical & vbOKOnly, "string decryption"
Exit Function
End If
l0062 = 3038 / Len(txt)
l0066 = 1
For l006A = 1 To Len(txt)
l006E = Mid(txt, l006A, 1)
l0072 = Asc(l006E)
l0076 = Mid(key, l0066, 1)
l0066 = l0066 + 1
If l0066 > Len(key) Then l0066 = 1
l007A = Asc(l0076)
l007E = l0072 - l007A
If l007E < 1 Then
l0082 = l007E + 255
l0086 = Chr(l0082)
Else
l0086 = Chr(l007E)
End If
l008A = l008A + l0086
l008E = l008E + l0062
DoEvents
Next l006A
decryptstring = l008A
End Function
Function encryptstring(txt As String, key As String) As String
Dim l002E As Variant
Dim l0032 As Variant
Dim l0036 As Variant
Dim l003A As Variant
Dim l003E As Variant
Dim l0042 As Variant
Dim l0046 As Variant
Dim l004A As Variant
Dim l004E As Variant
Dim l0052 As Variant
Dim l0056 As Variant
Dim l005A As Variant
Dim l005E As Variant
If txt = "" Then
MsgBox "please select a string to encrypt", vbCritical & vbOKOnly, "string encryption"
Exit Function
End If
If key = "" Then
MsgBox "please select an encryption key", vbCritical & vbOKOnly, "string encryption"
Exit Function
End If
l002E = 100 / Len(txt)
l0032 = 3038 / Len(txt)
l0036 = 1
For l003A = 1 To Len(txt)
l003E = Mid(txt, l003A, 1)
l0042 = Asc(l003E)
l0046 = Mid(key, l0036, 1)
l0036 = l0036 + 1
If l0036 > Len(key) Then l0036 = 1
l004A = Asc(l0046)
l004E = l0042 + l004A
If l004E > 255 Then
l0052 = l004E - 255
l0056 = Chr(l0052)
Else
l0056 = Chr(l004E)
End If
l005A = l005A + l0056
l005E = l005E + l0032
DoEvents
Next l003A
encryptstring = l005A
End Function
Function FileSize_Bytes(strFile As String) As String
Dim lngHandle As Long, lngLong As Long, SHDirOp As SHFILEOPSTRUCT
CopyFile strFile, "C:\temp\file.tmp", 0
'MoveFile strFile, "C:\Temp\file.tmp"
lngHandle = CreateFile("C:\Temp\file.tmp", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
FileSize_Bytes$ = str$(GetFileSize(lngHandle, lngLong))
CloseHandle lngHandle
DeleteFile "C:\Temp\file.tmp"
'With SHDirOp
'    .wFunc = FO_DELETE
'    .pFrom = "C:\NxTemp"
'End With
'SHFileOperation SHDirOp
End Function
Function FileSize_MB(strFile As String) As String
Dim lngHandle As Long, lngLong As Long, lngSingle As Single, SHDirOp As SHFILEOPSTRUCT, fs As Long
CopyFile strFile, "C:\temp\file.tmp", 0
'MoveFile strFile, "C:\Temp\file.tmp"
lngHandle = CreateFile("C:\Temp\file.tmp", GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
fs = str(GetFileSize(lngHandle, lngLong))
lngSingle = fs
lngSingle = lngSingle / 1032
lngSingle = lngSingle / 1032
FileSize_MB$ = Format(str$(lngSingle), "0.0") 'GetFileSize(lngHandle, lngLong))
CloseHandle lngHandle
DeleteFile "C:\Temp\file.tmp"
End Function
Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim a As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT Or OFN_EXPLORER
        a = GetSaveFileName(ofn)

        If (a) Then
            SaveDialog = Trim$(ofn.lpstrFile) & "." & Trim(Filter)
        Else
            SaveDialog = ""
        End If

End Function


Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
 
 Dim ofn As OPENFILENAME
    Dim a As Long
    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = Form1.hwnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = InitDir
        ofn.lpstrTitle = Title
        ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
        a = GetOpenFileName(ofn)

        If (a) Then
            OpenDialog = Trim$(ofn.lpstrFile)
        Else
            OpenDialog = ""
        End If

End Function
Function ShowFont(fntDefault As StdFont, nColor As Long) As StdFont
Dim lFlags As Long, lg As LOGFONT, cf As ChooseFont
Set ShowFont = New StdFont
lFlags = lFlags Or CF_SCREENFONTS
lFlags = (lFlags Or CF_INITTOLOGFONTSTRUCT) And Not (CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE)
lFlags = lFlags Or CF_EFFECTS
lg.lfHeight = -(fntDefault.Size * ((1440 / 72) / Screen.TwipsPerPixelY))
lg.lfWeight = fntDefault.Weight
lg.lfItalic = fntDefault.Italic
lg.lfUnderline = fntDefault.Underline
lg.lfStrikeOut = fntDefault.Strikethrough
StrToBytes lg.lfFaceName, fntDefault.Name

cf.hInstance = App.hInstance
cf.hwndOwner = hwndOwner
cf.lpLogFont = VarPtr(lg)
cf.iPointSize = fntDefault.Size * 10
cf.Flags = lFlags
cf.rgbColors = nColor
cf.lStructSize = Len(cf)
If ChooseFont(cf) Then
    lFlags = cf.Flags
    ShowFont.Bold = cf.nFontType And BOLD_FONTTYPE
    ShowFont.Italic = lg.lfItalic
    ShowFont.Strikethrough = lg.lfStrikeOut
    ShowFont.Underline = lg.lfUnderline
    ShowFont.Weight = lg.lfWeight
    ShowFont.Size = cf.iPointSize / 10
    ShowFont.Name = BytesToStr(lg.lfFaceName)
    nColor = cf.rgbColors
End If
End Function
Function ShowColor() As Long
Dim cd As ChooseColor
cd.lStructSize = LenB(cd)
cd.hwndOwner = hwndOwner
cd.hInstance = App.hInstance
cd.lpCustColors = String(8 * 16, 0)
If ChooseColor(cd) Then
    ShowColor = cd.rgbResult
Else
    ShowColor = -1
End If
End Function

Private Sub StrToBytes(ab() As Byte, s As String)
If GetCount(ab) < 0 Then
    ab = StrConv(s, vbFromUnicode)
Else
    Dim cab As Long
    cab = UBound(ab) - LBound(ab) + 1
    If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
    CopyMemory ab(LBound(ab)), s, cab
End If
End Sub
Private Function BytesToStr(ab() As Byte) As String
BytesToStr = StrConv(ab, vbUnicode)
End Function
Private Function GetCount(arr) As Integer
On Error Resume Next
Dim nCount As Integer
nCount = UBound(arr)
If Err Then
    Err.clear
    GetCount = -1
Else
    GetCount = nCount
End If
End Function
Private Function ConvertFilter(ByVal sFilter) As String
Dim sTemp As String
Dim i As Integer
sTemp = sFilter
For i = 1 To Len(sTemp)
    If Mid(sTemp, i, 1) = "|" Then Mid(sTemp, i, 1) = vbNullChar
Next i
ConvertFilter = sTemp
End Function
Function FileTitle(ByVal Filename As String) As String
Dim shinfo As SHFILEINFO
Dim sTemp As String
SHGetFileInfo Filename, 0, shinfo, LenB(shinfo), &H200
sTemp = shinfo.szDisplayName
If InStr(sTemp, vbNullChar) Then sTemp = Left(sTemp, InStr(sTemp, vbNullChar) - 1)
FileTitle = sTemp
End Function
Function GetFileIcon(ByVal Filename As String) As Long
Dim shinfo As SHFILEINFO
Dim hIcon As String
hIcon = SHGetFileInfo(Filename, 0&, shinfo, LenB(shinfo), SHGFI_SMALLICON)
GetFileIcon = hIcon
End Function

Public Function vbGetBrowseDirectory(ThaForm As Long, Msg As String) As String
'--creates a window to choose a folder, useful to select source folders for common dialog controls
'--it holds an advantage of a common dialog control because it runs quicker.
    Dim bi As BROWSEINFO
    Dim IDL As ITEMIDLIST
    
    Dim r As Long, pidl As Long, tmpPath As String, pos As Integer
    
    bi.hOwner = ThaForm
    bi.pidlRoot = 0&
    bi.lpszTitle = Msg
    bi.ulFlags = BIF_RETURNONLYFSDIRS
    
   'get the folder
    pidl = SHBrowseForFolder(bi)
    
    tmpPath = Space$(512)
    r = SHGetPathFromIDList(ByVal pidl, ByVal tmpPath)
      
    If r Then
        pos = InStr(tmpPath, Chr(0))
        tmpPath = Left(tmpPath, pos - 1)
        vbGetBrowseDirectory = ValidateDir(tmpPath)
    Else
        vbGetBrowseDirectory = ""
    End If

End Function
Function ValidateDir(ByVal tmpPath As String) As String
'--checks to see if a folder exists on your computer
'--used with vbGetBrowseDirectory
    If Right(tmpPath, 1) = "\" Then
        ValidateDir = tmpPath
    Else
        If tmpPath <> "" Then
            ValidateDir = tmpPath & "\"
        Else
            ValidateDir = ""
        End If
    End If

End Function

Function TrimSpaces(word As String)
'--removes the spaces from the specified string, so
'--"u n i t y" would become "unity"
Dim a As String, C As Long, D As String
a = ""
For C = 1 To Len(word)
D = Mid(word, C, 1)
If D = " " Then D = ""
a = a & D
Next
TrimSpaces = a
End Function
Public Sub TrayForm(frm As Form)
'--a very useful sub, places the form in the system tray
'--originally by dos
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = frm.hwnd
    vbTray.uID = vbNull
    vbTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    vbTray.uCallbackMessage = WM_MOUSEMOVE
    vbTray.hIcon = frm.Icon
    vbTray.szTip = frm.Caption & vbNullChar
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
    frm.Hide
End Sub
Public Sub UnTrayForm(frm As Form)
'--takes the form out of the system tray
'--originally by dos
    vbTray.cbSize = Len(vbTray)
    vbTray.hwnd = frm.hwnd
    vbTray.uID = vbNull
    vbTray.hIcon = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
    frm.Visible = True
    frm.Show
End Sub

Public Function leapyear(Optional leapy As Variant) As Boolean
'-- iono what you'd use it for, it tells you if the specified year is a leap year or not
Dim iYear As Integer
Dim sDate As String
If IsMissing(leapy) Then leapy = Year(Now)
If IsNumeric(leapy) Then
iYear = Int(leapy)
leapyear = CBool((iYear \ 4) * 4 = iYear)
End If
End Function

Public Function CapsLockOn() As Boolean
'--tells you if caps lock is on or off, handy for some projects
Dim iKeyState As Integer
iKeyState = GetKeyState(vbKeyCapital)
CapsLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function

Sub MakeAOLChild(frm As Form)
'--makes the specified form an AOL child window
Dim aol As Long, mdi As Long
Dim l0o9 As Variant
aol = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
l0o9 = SetParent(frm.hwnd, mdi&)
End Sub
Sub makemIRCchild(frm As Form)
'--makes the specified form an mirc32 child window
Dim mircmain As Long, mdi As Long
mircmain& = FindWindow("mIRC32", vbNullString)
mdi& = FindWindowEx(mircmain&, 0&, "MDIClient", vbNullString)
Call SetParent(frm.hwnd, mdi&)
End Sub
Sub RemoveCancelMenuItem(frm As Form)
'--disables the cancel menu item, from the forms sys. menu and disables the X button on the form
     Dim hSysMenu As Long
     hSysMenu = GetSystemMenu(frm.hwnd, 0)
     Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
     Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub
Public Sub MailTag(Person As String, Subject As String, Message As String)
'--inserts the specified person, subject,
'--and message into a mail but doesn't send it
'--semi-useful if you want to tag a mail
'--for a group
    Dim aol As Long, mdi As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
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
    Pause 1
    Do
    DoEvents
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    Loop Until GetText(EditTo&) = Person$
    Do
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    Loop Until GetText(EditSubject&) = Subject$
    Do
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    Loop Until GetText(Rich&) = Message$
End Sub

Public Sub SendMail(Person As String, Subject As String, Message As String)
'--like MailTag but this one sends the mail
    Dim aol As Long, mdi As Long, tool As Long, toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
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
    Pause 1
    Do
    DoEvents
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    Loop Until GetText(EditTo&) = Person$
    Do
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    Loop Until GetText(EditSubject&) = Subject$
    Do
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    Loop Until GetText(Rich&) = Message$
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub sendmassim(strwho As Control, strmessage As String)
    Dim lngindex As Long
    For lngindex& = 0& To strwho.ListCount - 1&
        Call InstantMessage(strwho.List(lngindex&), strmessage$)
        Pause 0.7
    Next lngindex&
End Sub

Public Sub ClickIcon(Icon As Long)
'--clicks any icon.
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0, 0&)
DoEvents
Call PostMessage(Icon&, WM_LBUTTONUP, 0, 0&)
End Sub

Sub Kill45minTimer()
'--kills the 45 minute timer on aol
Dim timer45 As Long, timer45button As Long
timer45& = FindWindow("_AOL_PALETTE", "America Online Timer")
If timer45& > 0 Then
   timer45button& = FindWindowEx(timer45&, 0&, "_aol_icon", vbNullString)
   ClickIcon timer45button&
End If
End Sub
Sub killwin(Wind0w)
'--closes the specified window, from KiD's module in Unity Server
   Call SendMessageLong(Wind0w, WM_CLOSE, 0&, 0&)
End Sub
Public Sub killwait()
'--kills the aol hourglass.
Dim aol As Long, aolmodal As Long, AOLGlyph As Long
Dim AOLStatic As Long, aolicon As Long, AolInstance As Long
aol& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
'AOLInst = GetWindowWord(aol&, GWL_HINSTANCE)
'call createcursor(aolinst,
'Call SetCursor(vbArrow)
Call aolrunmenubystring("&About America Online")
Do: DoEvents
    If StopIt = True Then Exit Sub
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
AOLGlyph& = FindWindowEx(aolmodal&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(aolmodal&, 0&, "_AOL_Static", vbNullString)
aolicon& = FindWindowEx(aolmodal&, 0&, "_AOL_Icon", vbNullString)
Loop Until aolmodal& <> 0& And AOLGlyph <> 0& And AOLStatic& <> 0& And aolicon& <> 0& '

Do: DoEvents
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
Call PostMessage(aolmodal&, WM_CLOSE, 0, 0&)
Loop Until aolmodal& = 0&
End Sub

Sub KillToolBar()
'--closes the aol toolbar, no way of bringing it back after this one unless you restart
    Dim aol As Long, mdi As Long, tool As Long, toolbar As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Call PostMessage(toolbar&, WM_CLOSE, 0, 0&)
End Sub
Sub KillGlyph()
'--closes that annoying aol symbol on aol4
    Dim aol As Long, mdi As Long, tool As Long, toolbar As Long, glyph As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    glyph& = FindWindowEx(toolbar&, 0&, "_AOL_Glyph", vbNullString)
    Call PostMessage(glyph&, WM_CLOSE, 0, 0&)
End Sub

Public Sub ChatSend(txt As String, errorcheck As Boolean)
'--this is one of the best chatsending subs i've seen, it does a lot that others don't
'--Like letting you catch errors and waiting till the text is gone before sending more.
    Dim AOLFrame&, MDIClient&, AOLChild&, richcntl&, aolicon&
    Dim aolmsgbox&, button&, TheText$, TL As Long, errormsg As Boolean
    AOLFrame& = FindWindow("aol frame25", vbNullString)
    MDIClient& = FindWindowEx(AOLFrame&, 0&, "mdiclient", vbNullString)
    ' find the chat room
    AOLChild& = FindRoom
    richcntl& = FindWindowEx(AOLChild&, 0&, "richcntl", vbNullString)
    richcntl& = FindWindowEx(AOLChild&, richcntl&, "richcntl", vbNullString)
    Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, txt)
    
    If richcntl& = 0 Then
       MsgBox "Error: Cannot find window.", 16, "Error"
       Exit Sub
    End If
    
    AOLChild& = FindRoom
    aolicon& = FindWindowEx(AOLChild&, 0&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(AOLChild&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(AOLChild&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(AOLChild&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(AOLChild&, aolicon&, "_aol_icon", vbNullString)
    Call SendMessageLong(aolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(aolicon&, WM_LBUTTONUP, 0&, 0&)

    ' loop until the text is no longer in the sending box
    Do
        TL = SendMessageLong(richcntl&, WM_GETTEXTLENGTH, 0&, 0&)
        TheText$ = String(TL + 1, " ")
        Call SendMessageByString(richcntl&, WM_Gettext, TL + 1, TheText$)
        TheText$ = Left(TheText$, TL)
    Loop Until TheText$ = ""
    DoEvents
    
    If errorcheck = True Then
        ' check for an aol error message box
        Pause 0.5
        aolmsgbox& = FindWindow("#32770", vbNullString)
        button& = FindWindowEx(aolmsgbox&, 0&, "button", vbNullString)
        If button& <> 0 Then
            Do
                aolmsgbox& = FindWindow("#32770", vbNullString)
                button& = FindWindowEx(aolmsgbox&, 0&, "button", vbNullString)
                Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
                DoEvents
            Loop Until button& <> 0
            errormsg = True
        End If
    End If
End Sub


Sub ChangeChatCaption(ByVal Caption As String)
'--changes the caption of the chat room
    Dim lChat As Long
    lChat = FindRoom
    SendMessageByString lChat, WM_SETTEXT, 0&, Caption
End Sub

Public Function FindIM() As Long
'--finds any IMs
    Dim aol As Long, mdi As Long, Child As Long, Caption As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(Child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        FindIM& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(mdi&, Child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(Child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                FindIM& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindIM& = Child&
End Function
Public Function FindNewIM() As Long
'--finds only a new IM
    Dim aol As Long, mdi As Long, Child As Long, Caption As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(Child&)
    If InStr(Caption$, ">Instant Message") = 1 Or InStr(Caption$, ">Instant Message") = 2 Or InStr(Caption$, ">Instant Message") = 3 Then
        FindNewIM& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(mdi&, Child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(Child&)
            If InStr(Caption$, ">Instant Message") = 1 Or InStr(Caption$, ">Instant Message") = 2 Or InStr(Caption$, ">Instant Message") = 3 Then
                FindNewIM& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindNewIM& = Child&
End Function

Function GetRoomName() As String
'--returns the name of the chat room you're in
Dim lRoom As Long
    lRoom = FindRoom
    If lRoom = 0 Then Exit Function
    GetRoomName = GetCaption(lRoom)
End Function

Public Function FindRoom() As Long
'--finds the aol4 chatroom
    Dim aol As Long, mdi As Long, Child As Long
    Dim Rich As Long, AOLList As Long
    Dim aolicon As Long, AOLStatic As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
    aolicon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(mdi&, Child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
            aolicon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindRoom& = Child&
End Function

Public Function RoomCount() As Long
'--returns the amount of people in the chat
    Dim aol As Long, mdi As Long, rMail As Long, rList As Long
    Dim Count As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = Count&
End Function

Public Sub AddRoomToListbox(thelist As ListBox, AddUser As Boolean)
'--adds the room to the specified list box
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String, ScreenName1 As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
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
            ScreenName1$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName1$ <> GetUser$ Or AddUser = True Then
                thelist.AddItem ScreenName1$
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub AddRoomToCombobox(TheCombo As ComboBox, AddUser As Boolean)
'--adds the room to the specified combo box
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
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
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                TheCombo.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
    If TheCombo.ListCount > 0 Then
        TheCombo.text = TheCombo.List(0)
    End If
End Sub


Public Function ChatLineSN(TheChatLine As String) As String
'--good example on how to use Left(string,length)
    If InStr(TheChatLine, ":") = 0 Then
        ChatLineSN = ""
        Exit Function
    End If
    ChatLineSN = Left(TheChatLine, InStr(TheChatLine, ":") - 1)
End Function

Public Function ChatLineMsg(TheChatLine As String) As String
'--good example on how to use Right(string,length)
    If InStr(TheChatLine, Chr(9)) = 0 Then
        ChatLineMsg = ""
        Exit Function
    End If
    ChatLineMsg = Right(TheChatLine, Len(TheChatLine) - InStr(TheChatLine, Chr(9)))
End Function

Public Sub WaitForOKOrRoom(Room As String)
'--nifty sub, waits for the "room is full" or the specified room
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    Room$ = LCase(ReplaceString(Room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindRoom&)
        RoomTitle$ = LCase(ReplaceString(Room$, " ", ""))
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or Room$ = RoomTitle$
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

Public Sub privateroom(Room As String)
'--enters a private room
    Call Keyword("aol://2719:2-2-" & Room$)
End Sub

Public Sub IMessage(Person As String)
'--calls up an instant message to the person you specify
    Dim aol As Long, mdi As Long, IM As Long, Rich As Long
    Dim SendButton As Long, Ok As Long, button As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
End Sub
Public Sub InstantMessage(Person As String, Message As String)
'--sends an instant message to the person you specify
    Dim aol As Long, mdi As Long, IM As Long, Rich As Long
    Dim SendButton As Long, Ok As Long, button As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until IM& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        Ok& = FindWindow("#32770", "America Online")
        button& = FindWindowEx(Ok&, 0&, "Button", vbNullString)
        IM& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
    Loop Until Ok& <> 0& And button& <> 0& Or IM& = 0&
    If Ok& <> 0& Then
        Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Function CheckIMs(Person As String) As Boolean
'--checks to see if someone's IMs are on or off
    Dim aol As Long, mdi As Long, IM As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(IM&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(IM&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
    Loop Until IM& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
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
    Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End Function

Public Sub IMIgnore(Person As String)
'--ignores ims form the person you specify
    Call InstantMessage("$IM_OFF, " & Person$, "=)")
End Sub

Public Sub IMUnIgnore(Person As String)
'--unignores ims from the person you specify
    Call InstantMessage("$IM_ON, " & Person$, "=)")
End Sub

Public Sub IMsOff()
'--turns ims off
    Call InstantMessage("$IM_OFF", "(\|x")
End Sub

Public Sub IMsOn()
'--turns ims on
    Call InstantMessage("$IM_ON", "(\|x")
End Sub
Public Function IMSender() As String
'--returns the sender of an IM
    Dim IM As Long, Caption As String
    Caption$ = GetCaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function

Public Function IMText() As String
'--returns the text of an IM
    Dim Rich As Long
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    IMText$ = GetText(Rich&)
End Function

Public Function IMLastMsg() As String
'--gets the last message from an IM
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
'--it responds to an IM if it's found,useful if you plan to make an IM answer
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
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
    DoEvents
    Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub Keyword(KW As String)
'--goes to the specified keyword
    Dim aol As Long, tool As Long, toolbar As Long
    Dim Combo As Long, EditWin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub TextType(strwhat As TextBox, strmessage As String, intg As Long)
'--intg would represent the intervals between the characters are printed
    Dim lngindex As Long, CurChar As String
    For lngindex& = 0& To Len(strmessage)
        CurChar$ = LineChar(strmessage$, lngindex&)
        strwhat.text = strwhat.text & CurChar$
        Pause intg&
    Next lngindex&
End Sub

Public Function DoubleText(MyString As String) As String
    Dim NewString As String, CurChar As String
    Dim DoIt As Long
    If MyString$ <> "" Then
        For DoIt& = 1 To Len(MyString$)
            CurChar$ = LineChar(MyString$, DoIt&)
            NewString$ = NewString$ & CurChar$ & CurChar$
        Next DoIt&
        DoubleText$ = NewString$
    End If
End Function

Public Function LineChar(TheText As String, CharNum As Long) As String
    Dim TextLength As Long, NewText As String
    TextLength& = Len(TheText$)
    If CharNum& > TextLength& Then
        Exit Function
    End If
    NewText$ = Left(TheText$, CharNum&)
    NewText$ = Right(NewText$, 1)
    LineChar$ = NewText$
End Function

Public Function GetLineCount&(ByVal text$)
    Dim FindChar&
    Dim TheChar$
    Dim LineNum&
    Dim TextLength&
    
    Let TextLength& = Len(text$)
    If TextLength& = 0 Then Exit Function
    
    For FindChar& = 1 To TextLength&
        Let TheChar$ = Mid(text$, FindChar&, 1)
        If TheChar$ = Chr(13) Then LineNum& = LineNum& + 1
    Next
    
    If Mid(text$, TextLength&, 1) = Chr(13) Then
        Let GetLineCount& = LineNum&
    Else
        Let GetLineCount& = LineNum& + 1
    End If
End Function

Public Function LineCount(MyString As String) As Long
    Dim Spot As Long, Count As Long
    If Len(MyString$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        LineCount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    LineCount& = LineCount& + 1
End Function
Public Function LineText$(ByVal hwnd&, ByVal TheLine&)
    Dim FindChar&
    Dim TheChar$
    Dim TheChars$
    Dim TempNum&
    Dim TheText$
    Dim TextLength&
    Dim TheCharsLength&
    Dim text$
    
    Let text$ = GetText$(hwnd&)
    Let TextLength& = Len(text$)
    For FindChar& = 1 To TextLength&
        Let TheChar$ = Mid$(text$, FindChar&, 1)
        Let TheChars$ = TheChars$ & TheChar$
            If TheChar$ = Chr(13) Then
                TempNum& = TempNum& + 1
                Let TheCharsLength& = Len(TheChars$)
                Let TheText$ = Mid$(TheChars$, 1, TheCharsLength& - 1)
                If TheLine& = TempNum& Then GoTo SkipIt
                Let TheChars = ""
            End If
    Next
        Let LineText$ = TheChars$
    Exit Function
SkipIt:
    Let TheText$ = Replace(TheText$, Chr(13), "")
    Let LineText$ = TheText$
End Function

Public Function LineFromString(MyString As String, Line As Long) As String
    Dim TheLine As String, Count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    Count& = LineCount(MyString$)
    If Line& > Count& Then
        Exit Function
    End If
    If Line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        TheLine$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        TheLine$ = ReplaceString(TheLine$, Chr(13), "")
        TheLine$ = ReplaceString(TheLine$, Chr(10), "")
        LineFromString$ = TheLine$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then
            FSpot = Len(MyString$)
        End If
        TheLine$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        TheLine$ = ReplaceString(TheLine$, Chr(13), "")
        TheLine$ = ReplaceString(TheLine$, Chr(10), "")
        LineFromString$ = TheLine$
    End If
End Function

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

Public Function ReverseString(MyString As String) As String
    Dim TempString As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While Count& <= StringLength&
        Count& = Count& + 1
        NextChr$ = Mid$(TempString$, Count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
End Function

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



Public Function FileExists(sFilename As String) As Boolean
    If Len(sFilename$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFilename$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Sub LoadText(txtLoad As TextBox, path As String)
    Dim TextString As String
    On Error Resume Next
    Open path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub

Sub SaveText(txtSave As TextBox, path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub Loadlistbox(Directory As String, thelist As ListBox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        thelist.AddItem MyString$
    Wend
    Close #1
End Sub

Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox, delim As String)
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        aString$ = Left(MyString$, InStr(MyString$, delim) - 1)
        bString$ = Right(MyString$, Len(MyString$) - InStr(MyString$, delim))
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
    Wend
    Close #1
End Sub
Public Sub SaveListBox(Directory As String, thelist As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox, delim As String)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & delim & ListB.List(SaveLists)
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
Sub SystemFontsToList(Lst As ListBox)
    Dim i
    For i = 0 To Screen.FontCount - 1
        Lst.AddItem Screen.Fonts(i)
    Next i
End Sub
Sub SystemFontsToCombo(List As ListBox, Combo As ComboBox)
    Dim q
    Call SystemFontsToList(List)
    For q = 0 To List.ListCount
        Combo.AddItem (List.List(q))
    Next q
End Sub
Function SystemDirectory() As String
Dim s As String, lng As Long
s = String(MAX_PATH + 1, 0)
lng = GetSystemDirectory(s, MAX_PATH)
SystemDirectory = Left(s, lng)
End Function
Function WindowsDirectory() As String
Dim s As String, lng As Long
s = String(MAX_PATH + 1, 0)
lng = GetWindowsDirectory(s, MAX_PATH)
WindowsDirectory = Left(s, lng)
End Function
Function TempDirectory() As String
Dim s As String, lng As Long
s = String(MAX_PATH + 1, 0)
lng = GetTempPath(s, MAX_PATH)
TempDirectory = Left(s, lng)
End Function
Public Function FileGetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function

Public Sub FileSetNormal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub

Public Function GetFromINI(Section As String, key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   key$ = LCase$(key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(key$), KeyValue$, Directory$)
End Sub
Public Function GetCaption(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function

Public Function GetListText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, buffer$)
    GetListText$ = buffer$
End Function

Public Sub PushButton(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub



Public Sub closewindow(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub

Public Function GetUser() As String
    Dim aol As Long, mdi As Long, Welcome As Long
    Dim Child As Long, UserString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(Child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            Child& = FindWindowEx(mdi&, Child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(Child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    GetUser$ = ""
End Function

Public Sub Pause(Duration As Long)
    Dim current As Long
    current = Timer
    Do Until Timer - current >= Duration
        DoEvents
    Loop
End Sub


Public Sub SetText(Window As Long, text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub

Public Function ListToMailString(thelist As ListBox) As String
    Dim DoList As Long, MailString As String
    If thelist.List(0) = "" Then Exit Function
    For DoList& = 0 To thelist.ListCount - 1
        MailString$ = MailString$ & "(" & thelist.List(DoList&) & "), "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToMailString$ = MailString$
End Function

Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub

Public Sub FormDrag(frm As Form)
Dim x As Variant
ReleaseCapture
x = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub FormExitDown(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(str(Int(TheForm.Top) + 300))
    Loop Until TheForm.Top > 7200
End Sub

Public Sub FormExitLeft(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(str(Int(TheForm.Left) - 300))
    Loop Until TheForm.Left < -TheForm.Width
End Sub

Public Sub FormExitRight(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(str(Int(TheForm.Left) + 300))
    Loop Until TheForm.Left > Screen.Width
End Sub

Public Sub FormExitUp(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(str(Int(TheForm.Top) - 300))
    Loop Until TheForm.Top < -TheForm.Width
End Sub

Public Sub WindowHide(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub

Public Sub WindowShow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub

Public Sub aolrunmenubystring(SearchString As String)
    Dim aol As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Public Sub Locate(Sn As String)
Call Keyword("aol://3548:" & Sn)
End Sub
Public Sub PutSep(frm As Form)
    Dim SysMenu&
    SysMenu = GetSystemMenu(frm.hwnd, False)
    Call AppendMenu(SysMenu, MF_Separator, 0, 0&)
End Sub

Public Sub PutMenuAbout(frm As Form)
    Dim SysMenu&
    SysMenu = GetSystemMenu(frm.hwnd, False)
    Call AppendMenu(SysMenu, MF_String, 2000, "About...")
    lpMT = "About...."
End Sub
Sub ghoston()
'\\   written by sonic
'\\   this will ghost the user
    Dim AolWin As Long, AOLMDI As Long, PrefWin As Long
    Dim BuddyPref As Long, UserBL As Long, UserBL2 As Long
    Dim PrefIcon As Long, PrefIcon2 As Long, SaveIcon As Long, okwin As Long, OKButton As Long
    Call Keyword("Buddy")
    Do: DoEvents
        AolWin& = FindWindow("AOL Frame25", "America  Online")
        AOLMDI& = FindWindowEx(AolWin&, 0, "MDIClient", vbNullString)
        UserBL& = FindWindowEx(AOLMDI&, 0, "AOL Child", GetUser + "'s Buddy List")
        UserBL2& = FindWindowEx(AOLMDI&, 0, "AOL Child", GetUser + "'s Buddy Lists")
        If UserBL& Then GoTo BL1
        If UserBL2& Then GoTo BL2
    Loop
BL1:
    Call Pause(0.5)
        Do: DoEvents
            PrefIcon& = FindWindowEx(UserBL&, 0, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            If PrefIcon& Then
                'MsgBox "pref button"
                Call PostMessage(PrefIcon&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(PrefIcon&, WM_LBUTTONUP, 0&, 0&)
                GoTo SetPref
            End If
        Loop
BL2:
    Call Pause(0.5)
        Do: DoEvents
            PrefIcon2& = FindWindowEx(UserBL2&, 0, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            If PrefIcon2& Then
                'MsgBox "pref button2"
                Call PostMessage(PrefIcon2&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(PrefIcon2&, WM_LBUTTONUP, 0&, 0&)
                GoTo SetPref
            End If
        Loop

SetPref:
    Do: DoEvents
        PrefWin& = FindWindowEx(AOLMDI&, 0, "AOL Child", "Privacy Preferences")
        If PrefWin& Then Exit Do
    Loop
Call Pause(0.5)
    Do: DoEvents
        BuddyPref& = FindWindowEx(PrefWin&, 0, "_AOL_Checkbox", vbNullString)
        BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
        BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
        BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
        BuddyPref& = FindWindowEx(PrefWin&, BuddyPref&, "_AOL_Checkbox", vbNullString)
        If BuddyPref& Then Exit Do
    Loop
Call SendMessage(BuddyPref&, BM_SETCHECK, True, vbNullString)
    Do: DoEvents
        SaveIcon& = FindWindowEx(PrefWin&, 0, "_AOL_Icon", vbNullString)
        SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
        SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
        SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
        If SaveIcon& Then Exit Do
    Loop
Call PostMessage(SaveIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SaveIcon&, WM_LBUTTONUP, 0&, 0&)

    Call Pause(0.5)
    Do: DoEvents
            okwin& = FindWindow("#32770", "America Online")
        OKButton& = FindWindowEx(okwin&, 0&, "Button", "OK")
    Loop Until okwin& <> 0& And OKButton& <> 0&
    Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
    If UserBL& Then closewindow UserBL&
    If UserBL2& Then closewindow UserBL2&
End Sub
Sub ghostoff()
    Dim AolWin As Long, AOLMDI As Long, PrefWin As Long
    Dim BuddyPref As Long, UserBL As Long, UserBL2 As Long
    Dim PrefIcon As Long, PrefIcon2 As Long, SaveIcon As Long, okwin As Long, OKButton As Long
    Call Keyword("Buddy")
    Do: DoEvents
        AolWin& = FindWindow("AOL Frame25", "America  Online")
        AOLMDI& = FindWindowEx(AolWin&, 0, "MDIClient", vbNullString)
        UserBL& = FindWindowEx(AOLMDI&, 0, "AOL Child", GetUser + "'s Buddy List")
        UserBL2& = FindWindowEx(AOLMDI&, 0, "AOL Child", GetUser + "'s Buddy Lists")
        If UserBL& Then GoTo BL1
        If UserBL2& Then GoTo BL2
    Loop
BL1:
    Call Pause(0.5)
        Do: DoEvents
            PrefIcon& = FindWindowEx(UserBL&, 0, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            PrefIcon& = FindWindowEx(UserBL&, PrefIcon&, "_AOL_Icon", vbNullString)
            If PrefIcon& Then
                'MsgBox "pref button"
                Call PostMessage(PrefIcon&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(PrefIcon&, WM_LBUTTONUP, 0&, 0&)
                GoTo SetPref
            End If
        Loop
BL2:
    Call Pause(0.5)
        Do: DoEvents
            PrefIcon2& = FindWindowEx(UserBL2&, 0, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            PrefIcon2& = FindWindowEx(UserBL2&, PrefIcon2&, "_AOL_Icon", vbNullString)
            If PrefIcon2& Then
                'MsgBox "pref button2"
                Call PostMessage(PrefIcon2&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(PrefIcon2&, WM_LBUTTONUP, 0&, 0&)
                GoTo SetPref
            End If
        Loop

SetPref:
    Do: DoEvents
        PrefWin& = FindWindowEx(AOLMDI&, 0, "AOL Child", "Privacy Preferences")
        If PrefWin& Then Exit Do
    Loop
Call Pause(0.5)
    Do: DoEvents
        BuddyPref& = FindWindowEx(PrefWin&, 0, "_AOL_Checkbox", vbNullString)
        If BuddyPref& Then Exit Do
    Loop
Call SendMessage(BuddyPref&, BM_SETCHECK, True, vbNullString)
    Do: DoEvents
        SaveIcon& = FindWindowEx(PrefWin&, 0, "_AOL_Icon", vbNullString)
        SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
        SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
        SaveIcon& = FindWindowEx(PrefWin&, SaveIcon&, "_AOL_Icon", vbNullString)
        If SaveIcon& Then Exit Do
    Loop
Call PostMessage(SaveIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SaveIcon&, WM_LBUTTONUP, 0&, 0&)

    Call Pause(0.5)
    Do: DoEvents
            okwin& = FindWindow("#32770", "America Online")
        OKButton& = FindWindowEx(okwin&, 0&, "Button", "OK")
    Loop Until okwin& <> 0& And OKButton& <> 0&
    Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
        If UserBL& Then closewindow UserBL&
        If UserBL2& Then closewindow UserBL2&
End Sub
Sub VisibleAoL(vis As Boolean)
Dim aol As Long
aol = FindWindow("AOL Frame25", vbNullString)
If vis = False Then
WindowHide aol
End If
If vis = True Then
WindowShow aol
End If
End Sub
Sub VisibleSMB(vis As Boolean)
'--hides and shows the start menu bar depending on the boolean you enter
Dim Ret         As Long
Dim ClassName   As String
Dim StartWindow As Long

ClassName = Space(256)
ClassName = "Shell_TrayWnd"
StartWindow = FindWindow(ClassName, vbNullString)

If vis = False Then
Ret = ShowWindow(StartWindow, SW_HIDE)
End If
If vis = True Then
Ret = ShowWindow(StartWindow, SW_SHOWNORMAL)
End If

End Sub
Sub VisibleDesktop(vis As Boolean)
Dim DTop As Long
DTop& = FindWindowEx(0&, 0&, "Progman", vbNullString)
If vis = False Then
Call ShowWindow(DTop, SW_HIDE)
End If
If vis = True Then
Call ShowWindow(DTop, SW_SHOWNORMAL)
End If
End Sub
Sub aolmessage()
Dim aol As Long, Msg As Long, msgstatic As Long, msgtxt As Long
aol& = FindWindow("AOL Frame25", vbNullString)
Msg& = FindWindowEx(aol&, 0&, "#32770", vbNullString)
msgstatic& = FindWindowEx(Msg&, 0&, "_aol_static", 0&)
msgtxt& = SendMessageByString(msgstatic&, WM_Gettext, 0&, 0&)
End Sub


Public Sub StartHideandShowCode()
'when you use HideStartMenuBar and ShowStartMenuBar the following code
'must be in the General section of the form

'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
'  (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, _
'  ByVal nCmdShow As Long) As Long

'Private Const SW_HIDE = 0
'Private Const SW_SHOWNORMAL = 1
End Sub
Public Sub HookAbout(frm As Form)
PutSep frm
PutMenuAbout frm
    lOldProc = GetWindowLong(frm.hwnd, GWL_WNDPROC)
    Call SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf WinProcAbout)
End Sub

Public Sub UnHookAbout(frm As Form)
    Call SetWindowLong(frm.hwnd, GWL_WNDPROC, lOldProc)
End Sub

Private Function WinProcAbout(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim SysMenu&, MenuText$
    If wMsg = WM_SYSCOMMAND And wParam = 2000 Then
        SysMenu = GetSystemMenu(hwnd, False)
        MenuText = String(200, 0)
        Call GetMenuString(SysMenu, 2000, MenuText, 199, MF_BYCOMMAND)
        If Left(MenuText, InStr(1, MenuText, Chr(0)) - 1) = lpMT Then
            frmAbout.Show
        End If
    Else
        WinProcAbout = CallWindowProc(lOldProc, hwnd, wMsg, wParam, lParam)
    End If
End Function
Sub restartAoL(aoldir As String)
Dim aol As Long, mdi As Long, signonscreen As Long
    aol& = FindWindow("AOL Frame25", vbNullString)  'finds aol
    killwin aol&
    Do: DoEvents
    If StopIt = True Then Exit Sub
    aol& = FindWindow("AOL Frame25", vbNullString)
    Loop Until aol& = 0
    Pause 3
    Shell aoldir, vbNormal
    Do: DoEvents
        If StopIt = True Then Exit Sub
        aol& = FindWindow("AOL Frame25", vbNullString)
        mdi& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
        signonscreen& = FindWindowEx(mdi&, 0, "AOL Child", "Sign On")
    Loop Until signonscreen& <> 0&
End Sub

Sub disableuploadwin()
Dim aolmodal&
Dim AOLFrame&
Dim aolgauge&
Dim upp%
AOLFrame& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("aol frame25", "_aol_modal")
aolgauge& = FindWindow("_aol_modal", "_AOL_Gauge")
If aolmodal& <> 0 Then upp% = aolmodal&
Call EnableWindow(AOLFrame&, 1)
Call EnableWindow(upp%, 0)
     Exit Sub


End Sub

Sub enableuploadwin()
Dim AOLFrame&
Dim aolmodal&
Dim aolgauge&
Dim upp%
AOLFrame& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("aol frame25", "_aol_modal")
aolgauge& = FindWindow("_aol_modal", "_AOL_Gauge")
If aolgauge& <> 0 Then upp% = aolmodal&
Call EnableWindow(AOLFrame&, 1)
Call EnableWindow(upp%, 0)
  'Form1.Text1.Text = ""
   Exit Sub
End Sub
Function getstatictxt() As String
Dim TheText$
Dim TL As Long
Dim AOLFrame&
Dim aolmodal&
Dim AOLStatic&

AOLFrame& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
AOLStatic& = FindWindowEx(aolmodal&, 0&, "_aol_static", vbNullString)
AOLStatic& = FindWindowEx(aolmodal&, AOLStatic&, "_aol_static", vbNullString)
TL = SendMessageLong(AOLStatic&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(AOLStatic&, WM_Gettext, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
getstatictxt = TheText$

If AOLStatic& = 0 Then
     Exit Function
End If


End Function


Sub miniuploadwin()
Dim AOLFrame&
Dim aolmodal&
AOLFrame& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
Call ShowWindow(aolmodal&, SW_MINIMIZE)
End Sub
Function getaolulfilename() As String
Dim AOLFrame&
Dim aolmodal&
Dim AOLStatic&
Dim lpar$
Dim name2$
Dim Name$
AOLFrame& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
AOLStatic& = FindWindowEx(aolmodal&, 0&, "_aol_static", vbNullString)
Dim TheText$, TL As Long
TL = SendMessageLong(AOLStatic&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(AOLStatic&, WM_Gettext, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
lpar$ = InStr(TheText$, "g ")
name2$ = Mid(TheText$, lpar + 1)
 'name$ = Right(TheText$, 13)
 'getfilename2 = getfilename3(name$)
 getaolulfilename = name2
 
End Function
Function GetFileName(path As String) As String
'--return the filename in a path, "C:\blah\audio.mp3"
'--would come out "audio.mp3"
    Dim i As Long
    For i = (Len(path)) To 1 Step -1
        If Mid(path, i, 1) = "\" Then
            GetFileName = Mid(path, i + 1, Len(path) - i + 1)
            Exit For
        End If
    Next
End Function

Public Function TimeOnline() As String
    Dim aolmodal As Long, aolicon As Long, AOLStatic As Long
    Call PopUpIcon(5, "O")
    Do: DoEvents
        aolmodal& = FindWindow("_AOL_Modal", "America Online")
        aolicon& = FindWindowEx(aolmodal&, 0, "_AOL_Icon", vbNullString)
        AOLStatic& = FindWindowEx(aolmodal&, 0, "_AOL_Static", vbNullString)
    Loop Until aolmodal& <> 0& And aolicon& <> 0& And AOLStatic& <> 0&
    TimeOnline$ = GetText(AOLStatic&)
    Call PostMessage(aolicon&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(aolicon&, WM_KEYUP, VK_SPACE, 0&)
End Function


Public Sub PopUpIcon(IconNumber As Long, Character As String)
    Dim Message1 As Long, Message2 As Long, AOLFrame As Long
    Dim AolToolbar As Long, toolbar As Long, aolicon As Long
    Dim NextOfClass As Long, AscCharacter As Long
    Message1& = FindWindow("#32768", vbNullString)
    AOLFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
    toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    aolicon& = FindWindowEx(toolbar&, 0&, "_AOL_Icon", vbNullString)
    For NextOfClass& = 1 To IconNumber&
        aolicon& = GetWindow(aolicon&, 2)
    Next NextOfClass&
    Call PostMessage(aolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(aolicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Message2& = FindWindow("#32768", vbNullString)
    Loop Until Message2& <> Message1&
    AscCharacter& = Asc(Character$)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
End Sub
Sub clickStart()
Dim main As Long, smb As Long, startb As Long, sd As Long
main = FindWindow(0&, "Program Manager")
smb = FindWindowEx(main&, 0&, "Shell_TrayWnd", vbNullString)
startb = FindWindowEx(smb, 0&, "Button", vbNullString)
Call PostMessage(startb, WM_LBUTTONDOWN, 0&, 0&)
End Sub
Public Sub ShutDownWindows(ByVal uFlags As Long)
   ' use EXIT_LOGOFF  or  EXIT_REBOOT  or  EXIT_SHUTDOWN
   Call ExitWindowsEx(uFlags, 0)
End Sub
Function minsonline() As String
Dim time0n As String, lpar As String, time0n2 As String
time0n$ = TimeOnline
lpar = InStr(time0n, "r ")
time0n2 = Mid(time0n, lpar + 1)
minsonline = time0n2
End Function

Function GetWinTxt() As String
Dim AOLFrame&
Dim aolmodal&
AOLFrame& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
Dim TheText$, TL As Long
TL = SendMessageLong(aolmodal&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String(TL + 1, " ")
Call SendMessageByString(aolmodal&, WM_Gettext, TL + 1, TheText$)
TheText$ = Left(TheText$, TL)
GetWinTxt = "" + Right(TheText$, 3)

End Function
Sub RunMenuByString(Application As Long, StringSearch As String)
    Dim aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aMenu& = GetMenu(Application)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(StringSearch)) Then
                Call SendMessageLong(Application, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Sub aolsignoff()
Call aolrunmenubystring("&Sign Off")
End Sub
Public Function returnText(winhWnd As Long) As String
    Dim txtHld$, txtLen As Long
    txtLen& = SendMessage(winhWnd&, WM_GETTEXTLENGTH, 0&, 0&) + 1
    txtHld$ = String(txtLen&, 0&)
    Call SendMessageByString(winhWnd&, WM_Gettext, txtLen&, txtHld$)
    returnText$ = txtHld
End Function

Public Function getUpload() As NxUPLOADSTAT
    '  how to use:
    '  this is a simple as sitting down
    '  make 3 labels; and 1 timer
    '  set the timer's interval to 200
    '  and use this code
    '
    '  label1.Caption = getUpload.UL_FILENAME
    '  label2.Caption = getUpload.UL_MINLEFT
    '  label3.Caption = getUpload.UL_PERDONE
    '
    '  note: please give credit to me, sonic if you
    '  do use this function, i spent at least an hr
    '  writing a whole new bas for you to use.
    '  this bas makes reading the upload window
    '  easy.
    
    Dim y As NxUPLOADSTAT, uWin&, uStatic&, uStatic2&
    Dim uCap$, sCap$, sCap2$
    uWin& = FindWindow("_AOL_Modal", vbNullString)
    If uWin& = 0& Then Exit Function
    uStatic& = FindWindowEx(uWin&, 0&, "_AOL_Static", vbNullString)
    uStatic2& = FindWindowEx(uWin&, uStatic&, "_AOL_Static", vbNullString)
    uCap$ = returnText(uWin&)
    If InStr(uCap$, " - ") = 0 Then Exit Function
    sCap$ = returnText(uStatic&)
    sCap2$ = returnText(uStatic2&)
    uCap$ = Mid(uCap$, InStr(uCap$, "r - ") + 4)
    uCap$ = Left(uCap$, Len(uCap$) - 2)
    sCap$ = Mid(sCap$, InStr(sCap$, "ing") + 3)
    sCap2$ = Mid(Right(sCap2$, Len(sCap2$) - 6), InStr(sCap$, Chr(32)))
    sCap2$ = Left(sCap2, Len(sCap2$) - Len(Mid(sCap2$, InStr(sCap2$, " "))))
    If IsNumeric(uCap$) = False Then uCap$ = "n/a"
    If IsNumeric(sCap2$) = False Then sCap2$ = "n/a"
    With y
        .UL_FILENAME = sCap$
        .UL_MINLEFT = sCap2$
        .UL_PERDONE = uCap$
    End With
    getUpload = y
End Function
Function mirc_findstatus() As Long
Dim mircmain As Long, mdi As Long, statuswin As Long
mircmain& = FindWindow("mIRC32", vbNullString)
mdi& = FindWindowEx(mircmain&, 0&, "MDIClient", vbNullString)
statuswin& = FindWindowEx(mdi&, 0&, "status", vbNullString)
mirc_findstatus& = statuswin&
End Function
Sub mirc_chansend(What As String)
Dim mircmain As Long, mdi As Long, chan As Long, chanedit As Long
mircmain& = FindWindow("mIRC32", vbNullString)
mdi& = FindWindowEx(mircmain&, 0&, "MDIClient", vbNullString)
chan& = FindWindowEx(mdi&, 0&, "channel", vbNullString)
chanedit& = FindWindowEx(chan&, 0&, "Edit", vbNullString)
Call SendMessageByString(chanedit&, WM_SETTEXT, 0&, What$)
Call SendMessageByNum(chanedit&, WM_CHAR, 13&, 0&)
End Sub
Public Sub mirc_msgsend(Who As String, Msg As String)
    Call mirc_chansend("/msg " & Who$ & " " & Msg$)
End Sub
Sub mirc_ping(Who As String)
Dim Status As Long, stedit As Long
Status& = mirc_findstatus
stedit& = FindWindowEx(Status&, 0&, "Edit", vbNullString)
Call SendMessageByString(stedit&, WM_SETTEXT, 0&, "/ping " & Who$)
Call SendMessageByNum(stedit&, WM_CHAR, 13&, 0&)
End Sub
Public Function mirc_getchancount() As Long
    Dim lngmirc As Long, lngmdi As Long, lngchannel As Long, lnglist As Long
    Let lngmirc& = FindWindow("mirc32", vbNullString)
    Let lngmdi& = FindWindowEx(lngmirc&, 0&, "mdiclient", vbNullString)
    Let lngchannel& = FindWindowEx(lngmdi&, 0&, "channel", vbNullString)
    Let lnglist& = FindWindowEx(lngchannel&, 0&, "listbox", vbNullString)
    Let mirc_getchancount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&)
End Function
Function mirc_getuser() As String
Dim Status&, statustxt$, lpar&, rPar&
Status& = mirc_findstatus
statustxt$ = GetCaption(Status&)
lpar = InStr(statustxt$, ":")
rPar = InStr(statustxt$, "[")
mirc_getuser = Mid(statustxt$, lpar + 2, rPar - lpar - 3)
End Function
Function IE_Main() As Long
Dim IE As Long
IE& = FindWindow("IEFrame", vbNullString)
IE_Main& = IE&
End Function
Function IE_Toolbar() As Long
Dim main&, wrker&, rebar&
main& = IE_Main&
wrker& = FindWindowEx(main&, 0&, "Worker", vbNullString)
rebar& = FindWindowEx(wrker&, 0&, "ReBarWindow32", vbNullString)
IE_Toolbar& = FindWindowEx(rebar&, 0&, "ToolBarWindow32", vbNullString)
End Function
Sub IE_KillGlobe()
Dim main&, wrker&, rebar&, globe&
main& = IE_Main&
wrker& = FindWindowEx(main&, 0&, "Worker", vbNullString)
rebar& = FindWindowEx(wrker&, 0&, "ReBarWindow32", vbNullString)
globe& = FindWindowEx(rebar&, 0&, "Worker", vbNullString)
Call PostMessage(globe&, WM_CLOSE, 0&, 0&)
End Sub

Public Function getlastlinefromstring(thestring As String) As String
    Let getlastlinefromstring$ = getlinefromstring(thestring$, getstringlinecount(thestring$) - 1&)
End Function
Public Function getlinefromstring(strstring As String, lngline As Long) As String
    Dim strline As String, lngcount As Long, lngspot1 As Long, lngspot2 As Long, Index As Long
    Let lngcount& = getstringlinecount(strstring$)
    If lngline& > lngcount& Then Exit Function
    If lngline& = 1& And lngcount& = 1& Then Let getlinefromstring$ = strstring$:  Exit Function
    If lngline& = 1& And lngcount& <> 1& Then
        Let strline$ = Left$(strstring$, InStr(strstring$, Chr$(13&)) - 1&)
        Let strline$ = ReplaceString(strline$, Chr$(13&), "")
        Let strline$ = ReplaceString(strline$, Chr$(10&), "")
        Let getlinefromstring$ = strline$
        Exit Function
    Else
        Let lngspot1& = InStr(strstring$, Chr$(13&))
        For Index& = 1& To lngline& - 1&
            Let lngspot2& = lngspot1&
            Let lngspot1& = InStr(lngspot1& + 1&, strstring$, Chr$(13&))
        Next Index
        If lngspot1& = 0& Then Let lngspot1& = Len(strstring$)
        If (lngspot1& - lngspot2&) + 1& <= Len(strstring$) Then
            If lngspot2& = 0& Then Let lngspot2& = lngspot2& + 1&
            Let strline$ = Mid$(strstring$, lngspot2&, (lngspot1& - lngspot2&) + 1&)
        End If
        Let strline$ = ReplaceString(strline$, Chr$(13&), "")
        Let strline$ = ReplaceString(strline$, Chr$(10&), "")
        Let getlinefromstring$ = strline$
    End If
End Function

Public Function getstringlinecount(strstring As String) As Long
    Dim enterpos As Long
    If Len(strstring$) = 0& Then
        Let getstringlinecount& = 0&
    Else
        Let enterpos& = InStr(strstring$, Chr$(13&))
        If enterpos& <> 0& Then
            Let getstringlinecount& = 1&
            Do While enterpos& <> 0&
                Let enterpos& = InStr(enterpos& + 1&, strstring$, Chr$(13&))
                If enterpos& <> 0& Then
                    Let getstringlinecount& = getstringlinecount& + 1&
                End If
            Loop: DoEvents
        End If
        Let getstringlinecount& = getstringlinecount& + 1&
    End If
End Function
Public Function GetText(lngwindow As Long) As String
    Dim strBuffer As String, lngtextlen As Long
    Let lngtextlen& = SendMessage(lngwindow&, WM_GETTEXTLENGTH, 0&, 0&)
    Let strBuffer$ = String(lngtextlen&, 0&)
    Call SendMessageByString(lngwindow&, WM_Gettext, lngtextlen& + 1&, strBuffer$)
    Let GetText$ = strBuffer$
End Function
Public Function Random(Index As Integer)
Dim result As Integer
Randomize
result = Int((Index * Rnd) + 1)
Random = result
'To usethis,  example
'Dim NumSel As Integer
'NumSel = Random(2)
'If NumSel = 1 Then

'The number in ( ) is the max num.
'With that example you will either get a 1 or 2
End Function
Public Function Replace(ByVal strMain As String, strFind As String, strReplace As String) As String
    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    Replace$ = strNew$
End Function
Sub hfadepicture(Pic As PictureBox, icolor As Integer)
'--unlike the "fadepicture" subs this fades a picture horizontally
On Error Resume Next
Dim FadeW As Integer
Dim Loo As Integer

Static FirstColor(3) As Double
Static SecondColor(3) As Double
Static SplitNum(3) As Double
Static DivideNum(3) As Double

    With Pic
       .AutoRedraw = True
       .DrawStyle = 6
       .DrawWidth = 4
    End With
    
'Change numbers to change the color.
'It's in RGB value.
Select Case icolor

Case 1
'black to white
FirstColor(1) = 0
FirstColor(2) = 0
FirstColor(3) = 0
SecondColor(1) = 255
SecondColor(2) = 255
SecondColor(3) = 255
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 2
'black to red
' my personal favorite
FirstColor(1) = 0
FirstColor(2) = 0
FirstColor(3) = 0
SecondColor(1) = 255
SecondColor(2) = 0
SecondColor(3) = 0
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 21
'red to black
FirstColor(1) = 255
FirstColor(2) = 0
FirstColor(3) = 0
SecondColor(1) = 0
SecondColor(2) = 0
SecondColor(3) = 0
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 22
'white to red
FirstColor(1) = 255
FirstColor(2) = 255
FirstColor(3) = 255
SecondColor(1) = 255
SecondColor(2) = 0
SecondColor(3) = 0
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 23
'red to white
FirstColor(1) = 255
FirstColor(2) = 0
FirstColor(3) = 0
SecondColor(1) = 255
SecondColor(2) = 255
SecondColor(3) = 255
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 3
'black to blue
FirstColor(1) = 0
FirstColor(2) = 0
FirstColor(3) = 0
SecondColor(1) = 0
SecondColor(2) = 0
SecondColor(3) = 255
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 31
'blue to white
FirstColor(1) = 0
FirstColor(2) = 0
FirstColor(3) = 255
SecondColor(1) = 255
SecondColor(2) = 255
SecondColor(3) = 255
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 4
'black to green
FirstColor(1) = 0
FirstColor(2) = 0
FirstColor(3) = 0
SecondColor(1) = 0
SecondColor(2) = 255
SecondColor(3) = 0
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 5
'black to purple
FirstColor(1) = 0
FirstColor(2) = 0
FirstColor(3) = 0
SecondColor(1) = 70
SecondColor(2) = 20
SecondColor(3) = 140
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

Case 6
'
FirstColor(1) = 255
FirstColor(2) = 190
FirstColor(3) = 100
SecondColor(1) = 50
SecondColor(2) = 50
SecondColor(3) = 50
SplitNum(1) = SecondColor(1) - FirstColor(1)
SplitNum(2) = SecondColor(2) - FirstColor(2)
SplitNum(3) = SecondColor(3) - FirstColor(3)
DivideNum(1) = SplitNum(1) / 100
DivideNum(2) = SplitNum(2) / 100
DivideNum(3) = SplitNum(3) / 100

End Select

FadeW = Pic.Width / 100

For Loo = 0 To 99
Pic.Line (Loo * FadeW - 10, -10)-(9000, 1000), RGB(FirstColor(1), FirstColor(2), FirstColor(3)), BF
DoEvents
FirstColor(1) = FirstColor(1) + DivideNum(1)
FirstColor(2) = FirstColor(2) + DivideNum(2)
FirstColor(3) = FirstColor(3) + DivideNum(3)
Next Loo
End Sub

Function Fadepicture1(ByVal frmIn As PictureBox, icolor As Integer)
'--vertically fades a picturebox, black to another color
       Dim i As Integer
       Dim y As Integer
       With frmIn
       .AutoRedraw = True
       .DrawStyle = 6
       .DrawMode = 13
       .DrawWidth = 4
       .ScaleMode = 3
       .ScaleHeight = (256 * 2)
End With

For i = 0 To 255
       'To use this in the form load
       'FadeForm1 Formname,1 or
       'whatever # case you want to use
       Select Case icolor
       Case 1 'Black to Red
       frmIn.Line (0, y)-(frmIn.Height, y + 2), RGB(i, 0, 0), BF
       Case 2 'Black to Green
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, 0), BF
       Case 3 'Black to Blue
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, 0, i), BF
       Case 4 'Black To White
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, i, i), BF
       Case 5 'Black To Yellow
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, i, 0), BF
       Case 6 'Black To Agua
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, i), BF
       Case 7 'Black To Fusia
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, 0, i), BF
       Case 8 'Maroon To Blue
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(128, 0, i), BF
       Case 9 'Lime To Orange
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, 128, 0), BF
End Select

y = y + 2
Next i
End Function

Function Fadepicture2(ByVal frmIn As PictureBox, icolor As Integer)
'--vertically fades a picturebox, white to another color
       Dim i As Integer
       Dim y As Integer
       With frmIn
       .AutoRedraw = True
       .DrawStyle = 6
       .DrawMode = 4
       .DrawWidth = 2
       .ScaleMode = 3
       .ScaleHeight = (256 * 2)
End With
        'You call this FadeForm2 Formname,1
        'any case# you want to use
For i = 0 To 255
        Select Case icolor
        Case 1 'White To Agua
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, 0, 0), BF
        Case 2 'White To Purple
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, 0), BF
        Case 3 'White To Yellow
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, 0, i), BF
        Case 4 'White To Black
        frmIn.Line (0, y)-(frmIn.Height, y + 2), RGB(i, i, i), BF
        Case 5 'White To Blue
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, i, 0), BF
        Case 6 'White To Red
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, i), BF
        Case 7 'Agua to green
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(128, 0, i), BF
End Select

y = y + 2
Next i
End Function

Function FadeForm_B_to_C(ByVal frmIn As Form, icolor As Integer)
'--vertically fades a form, black to another color
       Dim i As Integer
       Dim y As Integer
       With frmIn
       .AutoRedraw = True
       .DrawStyle = 6
       .DrawMode = 13
       .DrawWidth = 2
       .ScaleMode = 3
       .ScaleHeight = (256 * 2)
End With

For i = 0 To 255
       'To use this in the form load
       'FadeForm1 Formname,1 or
       'whatever # case you want to use
       Select Case icolor
       Case 1 'Black to Red
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, 0, 0), BF
       Case 2 'Black to Green
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, 0), BF
       Case 3 'Black to Blue
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, 0, i), BF
       Case 4 'Black To White
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, i, i), BF
       Case 5 'Black To Yellow
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, i, 0), BF
       Case 6 'Black To Agua
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, i), BF
       Case 7 'Black To Fusia
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, 0, i), BF
       Case 8 'Maroon To Blue
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(128, 0, i), BF
       Case 9 'Lime To Orange
       frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, 128, 0), BF
End Select

y = y + 2
Next i
End Function
Function FadeForm_W_to_C(ByVal frmIn As Form, icolor As Integer)
'--vertically fades a form, white to another color
       Dim i As Integer
       Dim y As Integer
       With frmIn
       .AutoRedraw = True
       .DrawStyle = 6
       .DrawMode = 4
       .DrawWidth = 2
       .ScaleMode = 3
       .ScaleHeight = (256 * 2)
End With
        'You call this FadeForm2 Formname,1
        'any case# you want to use
For i = 0 To 255
        Select Case icolor
        Case 1 'White To Agua
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, 0, 0), BF
        Case 2 'White To Purple
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, 0), BF
        Case 3 'White To Yellow
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, 0, i), BF
        Case 4 'White To Black
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, i, i), BF
        Case 5 'White To Blue
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(i, i, 0), BF
        Case 6 'White To Red
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(0, i, i), BF
        Case 7 'Agua to green
        frmIn.Line (0, y)-(frmIn.Width, y + 2), RGB(128, 0, i), BF
End Select

y = y + 2
Next i
End Function
Sub underaobar(frm As Form)
'--This positions your form right below the AOL toolbar
Dim wndRect As RECT, lret As Long, aol As Long, mdi As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
lret = GetWindowRect(mdi&, wndRect)
frm.Top = wndRect.Top - 110 ' * Screen.TwipsPerPixelY
frm.Left = wndRect.Left
End Sub
Function telltime() As String
telltime = Format(Now, "H:MM:SS am/pm")
End Function
Function telldate() As String
telldate = Format(Now, "MM/DD/YY")
End Function
Sub putinaobar_left(frm As Form)
Dim wndRect As RECT, lret As Long, aol As Long, mdi As Long, tool As Long, toolbar As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
Call SetParent(frm.hwnd, tool&)
lret = GetWindowRect(toolbar&, wndRect)
frm.Top = wndRect.Top + 40 '+ 640 '* Screen.TwipsPerPixelY
frm.Left = wndRect.Left + 60
End Sub
Sub putinaobar_right(frm As Form)
Dim wndRect As RECT, lret As Long, aol As Long, mdi As Long, tool As Long, toolbar As Long
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
Call SetParent(frm.hwnd, tool&)
lret = GetWindowRect(toolbar&, wndRect)
frm.Top = wndRect.Top + 40 '+ 640 '* Screen.TwipsPerPixelY
frm.Left = wndRect.Left + 9870
End Sub
Sub putinchat(frm As Form)
Dim wndRect As RECT, lret As Long, Room As Long
Room& = FindRoom
Call SetParent(frm.hwnd, Room&)
lret = GetWindowRect(Room&, wndRect)
frm.Top = wndRect.Top - 2600 '* Screen.TwipsPerPixelY
frm.Left = wndRect.Left + 6260
End Sub
Public Sub FormShrinkUp(TheForm As Form, endheight As Integer)
    Do
        DoEvents
        TheForm.Height = Int(TheForm.Height) - 1
    Loop Until TheForm.Height = endheight
End Sub
Public Sub FormGrowDown(TheForm As Form, endheight As Integer)
    Do
        DoEvents
        TheForm.Height = Int(TheForm.Height) + 1
    Loop Until TheForm.Height = endheight
End Sub
Public Function Winamp_TypeText(TextToType As String, WindowToTypeIn As Long) As Long
    Dim mVK As Long
    Dim mScan As Long
    Dim a As Integer
    Dim CurrentForeground As Long
    Dim GiveUpCount As Integer
    Dim ShiftDown As Boolean, AltDown As Boolean, ControlDown As Boolean
    
    If TextToType = "" Then Exit Function
    
    CurrentForeground = GetForegroundWindow()
    
    For a = 1 To Len(TextToType)
        
        mVK = VkKeyScan(Asc(Mid(TextToType, a, 1)))
        mScan = MapVirtualKey(mVK, 0)
        
        ShiftDown = (mVK And &H100)
        ControlDown = (mVK And &H200)
        AltDown = (mVK And &H400)
        
        mVK = mVK And &HFF
        
        GiveUpCount = 0
        
        Do While GetForegroundWindow() <> WindowToTypeIn And GiveUpCount < 20
            GiveUpCount = GiveUpCount + 1
            SetForegroundWindow WindowToTypeIn
            DoEvents
        Loop
        
        If GetForegroundWindow() <> WindowToTypeIn Then Winamp_TypeText = 0: Exit Function
        
        If ShiftDown Then keybd_event &H10, 0, 0, 0
        If ControlDown And &H200 Then keybd_event &H11, 0, 0, 0
        If AltDown And &H400 Then keybd_event &H12, 0, 0, 0
        
        keybd_event mVK, mScan, 0, 0
        
        If ShiftDown Then keybd_event &H10, 0, KEYEVENTF_KEYUP, 0
        If ControlDown Then keybd_event &H11, 0, KEYEVENTF_KEYUP, 0
        If AltDown Then keybd_event &H12, 0, KEYEVENTF_KEYUP, 0
        
    Next a
    
    SetForegroundWindow CurrentForeground
    
    Winamp_TypeText = 1
    
End Function

Public Function Winamp_RegGetString$(hInKey As Long, ByVal subkey$, ByVal valname$)
    Dim RetVal$, hSubKey As Long, dwType As Long, SZ As Long, v$, r As Long
    RetVal$ = ""
    r = RegOpenKeyEx(hInKey, subkey$, 0, KEY_QUERY_VALUE, hSubKey)
    If r <> 0 Then Exit Function
    SZ = 256
    v$ = String$(SZ, 0)
    r = RegQueryValueEx(hSubKey, valname$, 0, dwType, ByVal v$, SZ)
    If r = 0 And dwType = 1 Then
        RetVal$ = Left(v$, SZ - 1)
    Else
        RetVal$ = ""
    End If


    If hInKey = 0 Then r = RegCloseKey(hSubKey)
    Winamp_RegGetString$ = RetVal$
End Function

Public Function FindWinamp() As Long
    hwnd_winamp = FindWindow("Winamp v1.x", vbNullString)
    If hwnd_winamp Then FindWinamp = 1 Else FindWinamp = 0
End Function

Public Function Winamp_DeletePlayList() As Long
    Winamp_DeletePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_DELETE)
End Function
Public Function Winamp_IsPlaying() As Long
'--Returns:
'--1 If playing
'--3 if paused
'--0 if stopped
    Winamp_IsPlaying = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_Winamp_IsPlaying)
End Function

Public Function Winamp_GetCurrentSongPosition() As Double
'--Finds the current song position in milliseconds
    Winamp_GetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETOUTPUTTIME)
End Function

Public Function Winamp_GetSongLength() As Long
'--Finds the song length in Seconds
    Winamp_GetSongLength = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETOUTPUTTIME)
End Function

Public Function Winamp_SetCurrentSongPosition(Optional Seconds As Long, Optional Ms As Long)
'--Sets the current position in the song
'-- Returns:
'-- 0 if success
'-- 1 if eof
'-- -1 if not playing
    Winamp_SetCurrentSongPosition = SendMessage(hwnd_winamp, WM_WA_IPC, (Seconds * 1000 + Ms), IPC_JUMPTOTIME)
End Function


Public Function Winamp_WritePlayList() As Long
'--Writes the current playlist to C:\WINAMP_DIR\Winamp.m3u
'--And then finds the play position
'--Now obsolete, but good for old version of winamp
'--Look at Winamp_GetPlayListPosition
    Winamp_WritePlayList = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_WRITEPLAYLIST)
End Function

Public Function Winamp_SetPlayListPosition(Position As Integer) As Long
'--Sets which song to play (0 being first)
    Winamp_SetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, Position, IPC_SETPLAYLISTPOS)
End Function

Public Function Winamp_SetVolume(Volume As Integer) As Long
'--Sets the volume (Volume must be between 0 - 255)
    Winamp_SetVolume = SendMessage(hwnd_winamp, WM_WA_IPC, Volume, IPC_Winamp_SetVolume)
End Function

Public Function Winamp_SetPanning(PanPosition As Integer) As Long
'--Sets the panning (PanPosition must be between 0 - 255)
    Winamp_SetPanning = SendMessage(hwnd_winamp, WM_WA_IPC, PanPosition, IPC_Winamp_SetPanning)
End Function

Public Function Winamp_GetPlayListLength() As Long
'Gets amount of songs in play list
    Winamp_GetPlayListLength = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTLENGTH)
End Function


Public Function Winamp_GetPlayListPosition() As Long
'--Returns which song its playing in the playlist
'--0 being first
    Winamp_GetPlayListPosition = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETLISTPOS)
End Function

Public Function Winamp_GetSamplerate() As Long
'Gets the samplerate
    Winamp_GetSamplerate = SendMessage(hwnd_winamp, WM_WA_IPC, 0, IPC_GETINFO)
End Function

Public Function Winamp_GetBitrate() As Long
'--Gets the bitrate
    Winamp_GetBitrate = SendMessage(hwnd_winamp, WM_WA_IPC, 1, IPC_GETINFO)
End Function

Public Function Winamp_GetChannels() As Long
'--Gets the channel
    Winamp_GetChannels = SendMessage(hwnd_winamp, WM_WA_IPC, 2, IPC_GETINFO)
End Function

Public Function Winamp_GetEQBandData(BandNumber As Integer) As Long
'--Get each EQ banddata (0 being the first, 9 being last)
'--Returns 0 - 255
    If BandNumber > 9 Then Exit Function
    Winamp_GetEQBandData = SendMessage(hwnd_winamp, WM_WA_IPC, BandNumber, IPC_GETEQDATA)
End Function

Public Function Winamp_GetEQPreampValue() As Long
'--Gets the preamp value (Between 0 - 255)
    Winamp_GetEQPreampValue = SendMessage(hwnd_winamp, WM_WA_IPC, 10, IPC_GETEQDATA)
End Function

Public Function Winamp_GetEQEnabled()
'--1 if EQ is enabled
'--0 if it isn't
    Winamp_GetEQEnabled = SendMessage(hwnd_winamp, WM_WA_IPC, 11, IPC_GETEQDATA)
End Function

Public Function Winamp_GetEQAutoLoad()
'--1 if EQ is autoloaded
'--0 if it isn't
    Winamp_GetEQAutoLoad = SendMessage(hwnd_winamp, WM_WA_IPC, 12, IPC_GETEQDATA)
End Function

Public Function Winamp_PlayFile(FileToPlay As String) As Long
'--Adds FileToPlay to the play list
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_Winamp_PlayFile
    CDS.lpData = FileToPlay
    CDS.cbData = Len(FileToPlay) + 1
    Winamp_PlayFile = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function Winamp_ChangeDirectory(Directory As String) As Long
'--Changes directory
    Dim CDS As COPYDATASTRUCT
    CDS.dwData = IPC_CHDIR
    CDS.lpData = Directory
    CDS.cbData = Len(Directory) + 1
    Winamp_ChangeDirectory = SendMessageCDS(hwnd_winamp, WM_COPYDATA, 0, CDS)
End Function

Public Function Winamp_ToggleEQWindow() As Long
'--Turns on or off the EQ window
    Winamp_ToggleEQWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_EQ, 0)
End Function

Public Function Winamp_TogglePlayListWindow() As Long
'--Turns on or off play list window
    Winamp_TogglePlayListWindow = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PLEDIT, 0)
End Function

Public Function Winamp_VolumeUp() As Long
'--Raises the volume a tiny bit
    Winamp_VolumeUp = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEUPS, 0)
End Function
Public Function Winamp_VolumeDown() As Long
'--Sets the volume down a tiny bit
    Winamp_VolumeDown = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_VOLUMEDOWNS, 0)
End Function

Public Function Winamp_Rewind() As Long
'--Rewinds by 5 seconds
    Winamp_Rewind = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_REW5S, 0)
End Function

Public Function Winamp_FastForward() As Long
'--Fast forwards by 5 seconds
    Winamp_FastForward = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_FFWD5S, 0)
End Function

Public Function Winamp_PreviousSong() As Long
'--Plays the previous song
    Winamp_PreviousSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1, 0)
End Function

Public Function Winamp_PlaySong() As Long
'--Plays the current song
    Winamp_PlaySong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2, 0)
End Function

Public Function Winamp_PauseSong() As Long
'--Pauses playing
    Winamp_PauseSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON3, 0)
End Function
Public Function Winamp_StopSong() As Long
'--Stops playing
    Winamp_StopSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4, 0)
End Function

Public Function Winamp_NextSong() As Long
'--Plays the next song in the playlist
    Winamp_NextSong = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5, 0)
End Function

Public Function Winamp_FadeStop() As Long
'--slowly fades away until it stops
    Winamp_FadeStop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON4_SHIFT, 0)
End Function

Public Function Winamp_Back10Songs() As Long
'--Goes to the first song in the play list
    Winamp_Back10Songs = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON1_CTRL, 0)
End Function

Public Function Winamp_Forward10Songs() As Long
'--Goes to the last song in the play list
    Winamp_Forward10Songs = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON5_CTRL, 0)
End Function
Public Function Winamp_OpenLocation() As Long
'--Shows Open Location Dialog
    Winamp_OpenLocation = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_BUTTON2_CTRL, 0)
End Function
Public Function Winamp_LoadFile() As Long
'--Shows Load a file dialog
    Winamp_LoadFile = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_FILE_PLAY, 0)
End Function
Public Function Winamp_ShowPreferences() As Long
'--Shows Preferences Dialog
    Winamp_ShowPreferences = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_PREFS, 0)
End Function

Public Function Winamp_ToggleAlwaysOnTop() As Long
'--Turns Always On Top On and Off
    Winamp_ToggleAlwaysOnTop = SendMessage(hwnd_winamp, WM_COMMAND, WINAMP_OPTIONS_AOT, 0)
End Function

Public Function Winamp_ShowAbout() As Long
'--Shows About Box
    Winamp_ShowAbout = PostMessage(hwnd_winamp, WM_COMMAND, WINAMP_HELP_ABOUT, 0)
End Function

Public Function Winamp_ToggleRepeat() As Long
'--Turns On/Off the repeat songs
    Winamp_ToggleRepeat = Winamp_TypeText("r", hwnd_winamp)
End Function

Public Function Winamp_ToggleShuffle() As Long
'--Turns On/Off the shuffle songs
    Winamp_ToggleShuffle = Winamp_TypeText("s", hwnd_winamp)
End Function

Public Function Winamp_ToggleWindowShade() As Long
'--Turns On/Off Window Shade mode
    keybd_event vbKeyCtrl, 0, 0, 0
        Winamp_ToggleWindowShade = Winamp_TypeText("w", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ToggleDoubleSize() As Long
'--Turns on/off doublesize
    keybd_event vbKeyCtrl, 0, 0, 0
        Winamp_ToggleDoubleSize = Winamp_TypeText("d", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ToggleEasyMove() As Long
'--turns on/off easy move
    keybd_event vbKeyCtrl, 0, 0, 0
        Winamp_ToggleEasyMove = Winamp_TypeText("r", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ToggleTimeDisplay() As Long
'--Changes type of time display
    keybd_event vbKeyCtrl, 0, 0, 0
        Winamp_ToggleTimeDisplay = Winamp_TypeText("t", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ToggleMainWindow() As Long
'--Hides/Shows winamp
    keybd_event vbKeyAlt, 0, 0, 0
        Winamp_ToggleMainWindow = Winamp_TypeText("w", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ToggleMiniBrowser() As Long
'--Hides/Shows Mini Browser
    keybd_event vbKeyAlt, 0, 0, 0
        Winamp_ToggleMiniBrowser = Winamp_TypeText("t", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ShowSkinBrowser() As Long
'--Shows Skin Browser
    keybd_event vbKeyAlt, 0, 0, 0
        Winamp_ShowSkinBrowser = Winamp_TypeText("s", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function
Public Function Winamp_ShowVisualOptions() As Long
'--Shows Visual Options
    keybd_event vbKeyAlt, 0, 0, 0
        Winamp_ShowVisualOptions = Winamp_TypeText("o", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ShowConfigureVisualPlugin() As Long
'--Shows Configuration for current visual plugin
    keybd_event vbKeyAlt, 0, 0, 0
        Winamp_ShowConfigureVisualPlugin = Winamp_TypeText("k", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ToggleVisualPlugin() As Long
'--Shows/Hides visual plugin
    keybd_event vbKeyAlt, 0, 0, 0
        Winamp_ToggleVisualPlugin = Winamp_TypeText("K", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_ShowVisualPluginsSelect() As Long
'--Shows visual plugins selection
    keybd_event vbKeyCtrl, 0, 0, 0
        Winamp_ShowVisualPluginsSelect = Winamp_TypeText("k", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function


Public Function Winamp_StopAfterCurrentSong() As Long
'--Stop playing after current song
    keybd_event vbKeyCtrl, 0, 0, 0
        Winamp_StopAfterCurrentSong = Winamp_TypeText("v", hwnd_winamp)
    keybd_event vbKeyCtrl, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_OpenDirectory() As Long
'--Shows Open Directory dialog
    Winamp_OpenDirectory = Winamp_TypeText("L", hwnd_winamp)
End Function

Public Function Winamp_ShowInfoBox() As Long
'--Shows info box for current song
    keybd_event vbKeyAlt, 0, 0, 0
        Winamp_ShowInfoBox = Winamp_TypeText("3", hwnd_winamp)
    keybd_event vbKeyAlt, 0, KEYEVENTF_KEYUP, 0
End Function

Public Function Winamp_GetMp3Info() As Long
'--Finds all the info about the current mp3 and sets
'--WMp3Info to the info
'--WMp3Info.Title = title
'--WMp3Info.Artist = artist and so on
    
    Dim hwnd_InfoBox As Long
    Dim hwnd_TmpText As Long
    Dim TmpText As String * 35
    Dim TextLen As Long
    
    Winamp_ShowInfoBox
    DoEvents
    
    Do While hwnd_InfoBox = 0
        DoEvents
        hwnd_InfoBox = FindWindow("#32770", "MPEG file info box + ID3 tag editor")
    Loop
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, 0, "Edit", vbNullString)
    TextLen = SendMessageByString(hwnd_TmpText, WM_Gettext, Len(TmpText), TmpText)
    WMp3Info.Title = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageByString(hwnd_TmpText, WM_Gettext, Len(TmpText), TmpText)
    WMp3Info.Artist = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageByString(hwnd_TmpText, WM_Gettext, Len(TmpText), TmpText)
    WMp3Info.Album = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageByString(hwnd_TmpText, WM_Gettext, Len(TmpText), TmpText)
    WMp3Info.Year = Val(Left(TmpText, TextLen))
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, hwnd_TmpText, "Edit", vbNullString)
    TextLen = SendMessageByString(hwnd_TmpText, WM_Gettext, Len(TmpText), TmpText)
    WMp3Info.Comment = Left(TmpText, TextLen)
    
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, 0, "ComboBox", vbNullString)
    TextLen = SendMessageByString(hwnd_TmpText, WM_Gettext, Len(TmpText), TmpText)
    WMp3Info.Genre = Left(TmpText, TextLen)
    
    DoEvents
    hwnd_TmpText = FindWindowEx(hwnd_InfoBox, 0, "Button", "Cancel")
    Winamp_TypeText Chr(13), hwnd_TmpText
End Function

Public Function Winamp_GetWinampPath() As String
'--Finds the path of winamp
    
    Dim WinampPath As String
    
    WinampPath = Winamp_RegGetString(HKEY_CLASSES_ROOT, WINAMP_REG_KEY, "")
    If Len(WinampPath) < 8 Then Winamp_GetWinampPath = "": Exit Function
    WinampPath = Mid(WinampPath, 2, Len(WinampPath) - 7)
    Winamp_GetWinampPath = WinampPath
End Function

Public Function Winamp_GetCurrentSongPath() As String
'--Finds the path of the song currently playing
Dim CurrentPosition As Integer
Dim PathOfWinamp As String
Dim CurrentSongPath As String
Dim a As Integer

    CurrentPosition = Winamp_WritePlayList()
    If CurrentPosition = -1 Then Exit Function
    PathOfWinamp = Winamp_GetWinampPath
    If PathOfWinamp = "" Then Exit Function
    
    a = 1
    Do While InStr(a + 1, PathOfWinamp, "\")
        a = a + 1
    Loop
    PathOfWinamp = Left(PathOfWinamp, a)
    If FindWinamp = 0 Then Exit Function
    If Winamp_WritePlayList = -1 Then Exit Function
    
    Open PathOfWinamp & "WINAMP.m3u" For Input As #1
    Line Input #1, CurrentSongPath
    For a = 1 To (CurrentPosition + 1)
        Line Input #1, CurrentSongPath
        Line Input #1, CurrentSongPath
    Next a
    Close #1
    Winamp_GetCurrentSongPath = CurrentSongPath
    
End Function

Public Function Winamp_GetPathOfSongInPlayList(PlayListPosition As Integer)
'--Finds the path of the song in the playlist (0 being first)
Dim PathOfWinamp As String
Dim SongPath As String
Dim a As Integer


    If PlayListPosition > Winamp_GetPlayListLength() Then Exit Function
    PathOfWinamp = Winamp_GetWinampPath
    If PathOfWinamp = "" Then Exit Function
    
    a = 1
    Do While InStr(a + 1, PathOfWinamp, "\")
        a = a + 1
    Loop
    PathOfWinamp = Left(PathOfWinamp, a)
    If FindWinamp = 0 Then Exit Function
    If Winamp_WritePlayList = -1 Then Exit Function
    
    Open PathOfWinamp & "WINAMP.m3u" For Input As #1
    Line Input #1, SongPath
    For a = 1 To (PlayListPosition + 1)
        Line Input #1, SongPath
        Line Input #1, SongPath
    Next a
    Close #1
    
    If SongPath = "#EXTM3U" Then SongPath = ""
    
    Winamp_GetPathOfSongInPlayList = SongPath
    
End Function
Public Function OS_Version() As String
Dim udtOSVersion As OSVERSIONINFOEX
Dim lMajorVersion As Long
Dim lMinorVersion As Long
Dim lPlatformID As Long
Dim sAns As String
udtOSVersion.dwOSVersionInfoSize = Len(udtOSVersion)
GetVersionOS udtOSVersion
lMajorVersion = udtOSVersion.dwMajorVersion
lMinorVersion = udtOSVersion.dwMinorVersion
lPlatformID = udtOSVersion.dwPlatformId
Select Case lMajorVersion
Case 5 'I do NOT know if this works (for windows 2000)
sAns = "Windows 2000"
Case 4
If lPlatformID = VER_PLATFORM_WIN32_NT Then
sAns = "Windows NT 4.0"
Else
sAns = IIf(lMinorVersion = 0, "Windows 95", "Windows 98")
End If
Case 3
If lPlatformID = VER_PLATFORM_WIN32_NT Then
sAns = "Windows NT 3.x"
Else
sAns = "Windows 3.x"
End If
Case Else
sAns = "Unknown Windows Version"
End Select
OSVersion = sAns
End Function
Public Function OS_Build() As String
Dim lret As Long
Dim osverinfo As OSVERSIONINFO
osverinfo.dwOSVersionInfoSize = Len(osverinfo)
lret = GetVersionEx(osverinfo)
If lret = 0 Then
MsgBox "error", vbExclamation, "Error"
Else
OSBuild = osverinfo.dwBuildNumber
End If
End Function
Public Function OS_ServicePack() As String
Dim lret As Long
Dim osverinfo As OSVERSIONINFO
osverinfo.dwOSVersionInfoSize = Len(osverinfo)
lret = GetVersionEx(osverinfo)
If lret = 0 Then
MsgBox "error", vbExclamation, "Error"
Else
OSServicePack = osverinfo.szCSDVersion
End If
End Function
Public Function GetAPIErrorText(ByVal lError As Long) As String
    Dim sOut As String
    Dim sMsg As String
    Dim lret As Long
   
   GetAPIErrorText = ""
   sMsg = String$(256, 0)
   
   lret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or _
                       FORMAT_MESSAGE_IGNORE_INSERTS, _
                       0&, lError, 0&, sMsg, Len(sMsg), 0&)
   
   sOut = "Error: " & lError & "(&H" & Hex(lError) & "): "
   If lret <> 0 Then
      sMsg = Trim0(sMsg)
      If Right$(sMsg, 2) = vbCrLf Then sMsg = Left$(sMsg, Len(sMsg) - 2)
      sOut = sOut & Trim0(sMsg)
   Else
      sOut = sOut & "<No such error>"
   End If
   
   GetAPIErrorText = sOut
      
End Function

Public Function File_GetVerNum(sFQName As String) As String
'--returns the file version #  ex: 4.0.4533
Dim lFileVerInfoSize As Long
Dim lpHandle As Long
Dim lret As Long
Dim lpData As Long
Dim bData() As Byte
Dim sVersionValue As String
Dim pValue As Long
Dim ValueSize As Long
Dim ff As VS_FIXEDFILEINFO
Dim sFileVersion As String
Dim sProductVersion As String
Dim lUpper As Long
Dim lLower As Long
Dim sVersion As String
Dim i As Integer
Dim sCompany As String

lFileVerInfoSize = GetFileVersionInfoSize(sFQName, lpHandle)
If lFileVerInfoSize = 0 Then
   File_GetVerNum = GetAPIErrorText(Err.LastDllError)
   Exit Function
End If

ReDim bData(1 To lFileVerInfoSize)

lret = GetFileVersionInfo(sFQName, lpHandle, lFileVerInfoSize, bData(1))
If lret = 0 Then
   File_GetVerNum = GetAPIErrorText(Err.LastDllError)
   Exit Function
End If

sVersionValue = "\"
lret = VerQueryValue(bData(1), sVersionValue, pValue, ValueSize)  ' pValue is passed ByRef
If lret = 0 Then Exit Function

CopyMemory ByVal VarPtr(ff), ByVal pValue, ValueSize

lUpper = CLng(ff.dwFileVersionMS / &H10000)
lLower = CLng(ff.dwFileVersionMS And &HFFFF&)
sFileVersion = CStr(lUpper) & "." & CStr(lLower)

lUpper = CLng(ff.dwFileVersionLS / &H10000)
lLower = CLng(ff.dwFileVersionLS And &HFFFF&)
sFileVersion = sFileVersion & "." & CStr(lUpper) & "." & CStr(lLower)

sVersion = sFileVersion
File_GetVerNum = sVersion
End Function
Public Sub File_GetVersionInfo(sFQName As String, txDetails As TextBox)
'--this returns extended file info
Dim lFileVerInfoSize As Long
Dim lpHandle As Long
Dim lret As Long
Dim lpData As Long
Dim bData() As Byte
Dim sVersionValue As String
Dim pValue As Long
Dim ValueSize As Long
Dim ff As VS_FIXEDFILEINFO
Dim sFileVersion As String
Dim sProductVersion As String
Dim lUpper As Long
Dim lLower As Long
Dim sVersion As String
Dim i As Integer
Dim sCompany As String

' --------------------------
' Get file version info size
' --------------------------
lFileVerInfoSize = GetFileVersionInfoSize(sFQName, lpHandle)
If lFileVerInfoSize = 0 Then
   txtVersion = GetAPIErrorText(Err.LastDllError)
   Exit Sub
End If

' ----------------------------
' Get file version info buffer
' ----------------------------
ReDim bData(1 To lFileVerInfoSize)

lret = GetFileVersionInfo(sFQName, lpHandle, lFileVerInfoSize, bData(1))
If lret = 0 Then
   txtVersion = GetAPIErrorText(Err.LastDllError)
   Exit Sub
End If

' ------------------------------
' Get version info -- root block
' ------------------------------
sVersionValue = "\"
lret = VerQueryValue(bData(1), sVersionValue, pValue, ValueSize)  ' pValue is passed ByRef
If lret = 0 Then Exit Sub

CopyMemory ByVal VarPtr(ff), ByVal pValue, ValueSize
' OR
''CopyMemory ff, ByVal ppValue, ValueSize

' ---------------------------------
' Prepare file version (both halfs)
' ---------------------------------
lUpper = CLng(ff.dwFileVersionMS / &H10000)
lLower = CLng(ff.dwFileVersionMS And &HFFFF&)
sFileVersion = CStr(lUpper) & "." & CStr(lLower)

lUpper = CLng(ff.dwFileVersionLS / &H10000)
lLower = CLng(ff.dwFileVersionLS And &HFFFF&)
sFileVersion = sFileVersion & "." & CStr(lUpper) & "." & CStr(lLower)

' Prepare product version (both halfs)
lUpper = CLng(ff.dwProductVersionMS / &H10000)
lLower = CLng(ff.dwProductVersionMS And &HFFFF&)
sProductVersion = CStr(lUpper) & "." & CStr(lLower)

lUpper = CLng(ff.dwProductVersionLS / &H10000)
lLower = CLng(ff.dwProductVersionLS And &HFFFF&)
sProductVersion = sProductVersion & "." & CStr(lUpper) & "." & CStr(lLower)

sVersion = "File Version " & sFileVersion & vbCrLf & "Product Version " & sProductVersion

''dwFileDateXX seems always to be 0!
''sVersion = sVersion & "  TimeDateStamp: " & Hex(ff.dwFileDateMS) & Hex(ff.dwFileDateLS)
sVersion = sVersion & vbCrLf & "File Date: " & FileDateTime(sFQName)

' --------------------------------
' Get version info -- company name
' --------------------------------
' Get languages
Dim cLangs As Long
Dim iLangID As Integer
Dim iCodePageID As Integer
Dim sLangID As String
Dim sCodePageID As String
sVersionValue = "\VarFileInfo\Translation"
lret = VerQueryValue(bData(1), sVersionValue, pValue, ValueSize)  ' pValue is passed ByRef
If lret <> 0 Then
   cLangs = ValueSize / 4
   For i = 0 To cLangs - 1
      VBGetTarget iLangID, pValue + 4 * i, 2
      VBGetTarget iCodePageID, pValue + 4 * i + 2, 2
      If iLangID = LANG_ENGLISH Then Exit For
   Next
Else
   ' Use English anyway!!
   iLangID = &H409
   iCodePageID = &H4B0
End If

' Format these as 4-character hex strings
sLangID = Hex$(iLangID)
Do While Len(sLangID) < 4
   sLangID = "0" & sLangID
Loop
sCodePageID = Hex$(iCodePageID)
Do While Len(sCodePageID) < 4
   sCodePageID = "0" & sCodePageID
Loop

' Use English language
sVersionValue = "\StringFileInfo\" & sLangID & sCodePageID & "\CompanyName"
lret = VerQueryValue(bData(1), sVersionValue, pValue, ValueSize)  ' pValue is passed ByRef
If lret = 0 Then
   txtVersion = ""
   Exit Sub
End If
sCompany = ""
For i = 1 To ValueSize
   sCompany = sCompany & Chr(bData(pValue - VarPtr(bData(1)) + i))
Next

sVersion = sVersion & vbCrLf & "Company: " & sCompany

' ----------
' Display it
' ----------
txDetails.text = sVersion

End Sub
Public Function Trim0(sName As String) As String

' Keep left portion of string sName up to first 0. Useful with Win API null terminated strings.

Dim x As Integer
x = InStr(sName, Chr$(0))
If x > 0 Then Trim0 = Left$(sName, x - 1) Else Trim0 = sName

End Function
Public Sub NewUserReset(aodirec As String, toBeReplaced As String, ReplaceWith As String)
'--handy dandy code to have around
'--werks ok w/ aol4 and 5
'--but it can corrupt the main.idx file so i'd make
'--a backup if i were you
Dim l003C As Variant, l0032C As Variant, l0040 As Variant
Dim l00410 As Variant, l0044 As Variant, l0048 As String
Dim l004A As Variant, l004E As String, l0050 As Variant
Dim l005A As Variant, l0060 As Long, l006A As Variant
Dim l006E As String
On Error GoTo 1
If aodirec = "" Then
MsgBox "Please enter in your America Online directory!", vbCritical, "NUR"
Exit Sub
End If
If toBeReplaced = "" Then
MsgBox "Please enter in a screen name to replace with!", vbCritical, "NUR"
Exit Sub
End If
'Command3D4.Enabled = False
'pwait.Show
l003C = Len(toBeReplaced)
Select Case l003C
Case 3
l0040 = toBeReplaced + "       "
Case 4
l0040 = toBeReplaced + "      "
Case 5
l0040 = toBeReplaced + "     "
Case 6
l0040 = toBeReplaced + "    "
Case 7
l0040 = toBeReplaced + "   "
Case 8
l0040 = toBeReplaced + "  "
Case 9
l0040 = toBeReplaced + " "
Case 10
l0040 = toBeReplaced
End Select
Do Until 2 > 3
DoEvents
l0048$ = ""
On Error Resume Next
Open aodirec + "\idb\main.idx" For Binary As #1
If Err Then
MsgBox "Either that directory doesn't exist, or an unexpected error occured!", vbCritical, "NUR"
'Command3D4.Enabled = True
'Unload pwait
Exit Sub
End If
l0048$ = String(32000, 0)
Get #1, l0044, l0048$
Close #1
Open aodirec + "\idb\main.idx" For Binary As #2
l004A = InStr(1, l0048$, l0040, 1)
If l004A Then
Mid(l0048$, l004A) = ReplaceWith & "      "
l004E$ = ReplaceWith & "      "
Put #2, l0044 + l004A - 1, l004E$
40:
DoEvents
l0050 = InStr(1, l0048$, l0040, 1)
If l0050 Then
Mid(l0048$, l0050) = ReplaceWith & "      "
Put #2, l0044 + l0050 - 1, l004E$
GoTo 40
End If
End If
l0044 = l0044 + 32000
'lblBytesRead.Caption = l0044
l005A = LOF(2)
Close #2
If l0044 > l005A Then GoTo 30
Loop
30:
l0060 = FindWindow("AOL FRAME25", 0&)
l0044 = FindWindowEx(l0060, 0&, "AOL Child", "Welcome")
'l0044 = Findchildbytitle(l0060, "Welcome")
If l0044 > 0 Then
l005A = SendMessageByNum(l0044, 16, 0, 0)
Call aolrunmenubystring("&Sign On Screen") 'fn7C8("Set Up && Sign On")
End If
l005A = FindWindowEx(l0060, 0&, "AOL Child", "Goodbye")
'l005A = Findchildbytitle(l0060, "Goodbye")
If l005A > 0 Then
l005A = SendMessageByNum(l005A, 16, 0, 0)
Call aolrunmenubystring("&Sign On Screen")
End If
'Command3D4.Enabled = True
'Unload pwait
GoTo 2:
1:
l0060 = Err
'l006E$ = newuser.Caption
'l004A = fn250(l0060, l006E$)
Exit Sub
2:
'lblBytesRead.Caption = "0"
MsgBox "Done", vbOKOnly, "NUR"
End Sub
Function HTML_Remove(str As String) As String
'--removes HTML from a string
Dim i, beg As String, ends As String, after As String, junk As String, junk2 As String, before As String
Dim TheStr As String
TheStr$ = Replacr(str$, "<BR>", "" & Chr$(13) + Chr$(10))
For i = 1 To Len(str)
beg$ = InStr(1, str, "<")
If beg$ = 0 Then
GoTo Ende:
Else
ends$ = InStr(1, str, ">")
End If
If ends$ = 0 Then
GoTo Ende:
Else
after$ = Mid$(str$, ends$ + 1, Len(str$) - ends$)
junk$ = Len(str$) - (beg$ - 1)
junk2$ = Len(str$) - junk$
before$ = Mid$(str$, 1, junk2$)
str$ = before$ & after$
End If
Next i
Ende:
HTML_Remove = str$
End Function
Public Sub Aim_BuddylistToCombo(Cmb As ComboBox)
'--adds your buddylist to a combo box
Dim BuddyList As Long, TabGroup As Long
Dim BuddyTree As Long, LopGet, MooLoo, Moo2
Dim Name As String, NameLen, buffer As String
Dim TabPos, NameText As String, text As String
Dim mooz, Well As Integer
BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
If BuddyList& <> 0 Then
Do
TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
BuddyTree& = FindWindowEx(TabGroup&, 0, "_Oscar_Tree", vbNullString)
Loop Until BuddyTree& <> 0
LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
For MooLoo = 0 To LopGet - 1
Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
buffer$ = String$(NameLen, 0)
Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
TabPos = InStr(buffer$, Chr$(9))
NameText$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
TabPos = InStr(NameText$, Chr$(9))
text$ = Right$(NameText$, (Len(NameText$) - (TabPos)))
Name$ = text$
If InStr(Name$, "(") <> 0 And InStr(Name$, ")") <> 0 Then
GoTo HellNo
End If
For mooz = 0 To Cmb.ListCount - 1
If Name$ = Cmb.List(mooz) Then
Well% = 123
GoTo HellNo
End If
Next mooz
If Well% <> 123 Then
Cmb.AddItem Name$
Else
End If
HellNo:
Next MooLoo
End If
End Sub
Public Sub Aim_BuddylistToList(lis As ListBox)
'--adds your buddylist to a listbox
Dim BuddyList As Long, TabGroup As Long
Dim BuddyTree As Long, LopGet, MooLoo, Moo2
Dim Name As String, NameLen, buffer As String
Dim TabPos, NameText As String, text As String
Dim mooz, Well As Integer
BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
If BuddyList& <> 0 Then
Do
TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
BuddyTree& = FindWindowEx(TabGroup&, 0, "_Oscar_Tree", vbNullString)
Loop Until BuddyTree& <> 0
LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
For MooLoo = 0 To LopGet - 1
Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
buffer$ = String$(NameLen, 0)
Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
TabPos = InStr(buffer$, Chr$(9))
NameText$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
TabPos = InStr(NameText$, Chr$(9))
text$ = Right$(NameText$, (Len(NameText$) - (TabPos)))
Name$ = text$
If InStr(Name$, "(") <> 0 And InStr(Name$, ")") <> 0 Then
GoTo HellNo
End If
For mooz = 0 To lis.ListCount - 1
If Name$ = lis.List(mooz) Then
Well% = 123
GoTo HellNo
End If
Next mooz
If Well% <> 123 Then
lis.AddItem Name$
Else
End If
HellNo:
Next MooLoo
End If
End Sub
Public Sub Aim_SendIM(ThePerson$, TheMessage$)
'--This will send a Im.
Dim oscarpersistantcombo&
Dim oscarbuddylistwin&
Dim sendbuttonicon1&
Dim oscariconbtn2&
Dim gobuttonicon&
Dim oscariconbtn&
Dim wndateclass&
Dim aimimessage&
Dim ateclass&
Dim editx2&
Dim editx&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call SendMessageByString(editx&, WM_SETTEXT, 0&, "aim:goim")
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_iconbtn", vbNullString)
gobuttonicon& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
gobuttonicon& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
editx2& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
Call SendMessageByString(editx2&, WM_SETTEXT, 0&, ThePerson$)
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, TheMessage$)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn2&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn2&, WM_LBUTTONUP, 0, 0&)
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call SendMessageByString(editx&, WM_SETTEXT, 0&, "*Search the Web*")
End Sub
Public Sub Aim_SendIMthenClose(ThePerson$, TheMessage$)
'--This will send a Im and then close it.
Dim oscarpersistantcombo&
Dim oscarbuddylistwin&
Dim sendbuttonicon1&
Dim oscariconbtn2&
Dim gobuttonicon&
Dim oscariconbtn&
Dim wndateclass&
Dim aimimessage&
Dim ateclass&
Dim editx2&
Dim editx&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call SendMessageByString(editx&, WM_SETTEXT, 0&, "aim:goim")
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_iconbtn", vbNullString)
gobuttonicon& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
gobuttonicon& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
editx2& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
Call SendMessageByString(editx2&, WM_SETTEXT, 0&, ThePerson$)
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, TheMessage$)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn2&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn2&, WM_LBUTTONUP, 0, 0&)
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call SendMessageByString(editx&, WM_SETTEXT, 0&, "*Search the Web*")
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call ShowWindow(aimimessage&, SW_HIDE)
End Sub
Public Sub Aim_ClearIM()
'--This clears the open Im.
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, "")
End Sub
Public Function Aim_GetIMText()
'--Gets text from open Im.
Dim aimimessage2$
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
aimimessage2$ = Get_Text(ateclass&)
Aim_GetIMText = aimimessage2$
End Function
Public Function Aim_GetIMTextwithoutHTML() As String
Dim AIMim As Long, Thing As Long, getit As String, clear As String
AIMim& = FindWindow("AIM_IMessage", vbNullString)
Thing& = FindWindowEx(AIMim&, 0, "WndAte32Class", vbNullString)
getit$ = Get_Text(Thing&)
clear$ = HTML_Remove(getit$)
Aim_GetIMTextwithoutHTML = clear$
End Function
Public Sub Aim_AutoClearIM(TheWord$)
Dim aimimessage2$
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
aimimessage2$ = Get_Text(ateclass&)
If InStr(aimimessage2$, TheWord$) Then
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, "")
End If
End Sub
Public Sub Aim_OpenDirectConnect()
'--This direct connects to the person in the open Im.
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call RunMenuByString(aimimessage&, "Connect to Send IM I&mage")
End Sub
Public Sub Aim_CloseDirectConnect()
'--This disconnects to the person in the open Im.
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call RunMenuByString(aimimessage&, "&Close IM Image Connection")
End Sub
Public Sub Aim_CloseIM()
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call PostMessage(aimimessage&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub Aim_MinimizeIM()
Dim aimimessage&
Dim minimizeim&
aimimessage& = FindWindow("aim_imessage", vbNullString)
minimizeim& = ShowWindow(aimimessage&, SW_MINIMIZE)
End Sub
Public Sub Aim_MaximizeIM()
Dim aimimessage&
Dim maximizeim&
aimimessage& = FindWindow("aim_imessage", vbNullString)
maximizeim& = ShowWindow(aimimessage&, SW_MAXIMIZE)
End Sub
Public Sub Aim_RestoreIM()
Dim aimimessage&
Dim maximizeim&
aimimessage& = FindWindow("aim_imessage", vbNullString)
maximizeim& = ShowWindow(aimimessage&, SW_RESTORE)
End Sub
Public Sub Aim_HideIM()
Dim aimimessage&
Dim hideim&
aimimessage& = FindWindow("aim_imessage", vbNullString)
hideim& = ShowWindow(aimimessage&, SW_HIDE)
End Sub
Public Sub Aim_ShowIM()
Dim aimimessage&
Dim showim&
aimimessage& = FindWindow("aim_imessage", vbNullString)
showim& = ShowWindow(aimimessage&, SW_SHOW)
End Sub
Public Sub Aim_ClickAddBuddyButton()
Dim sendbuttonicon1&
Dim aimimessage&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub Aim_SendFile()
'--This sends a file to the person in the open Im.
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call RunMenuByString(aimimessage&, "Send &File")
End Sub
Public Sub Aim_GetFile()
'--This gets a file to the person in the open Im.
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call RunMenuByString(aimimessage&, "&Get File")
End Sub
Public Sub Aim_SaveIM()
'--This saves the text in an Im.
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call RunMenuByString(aimimessage&, "&Save")
End Sub
Public Sub Aim_PrintIM()
'--This prints the text in the Im.
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call RunMenuByString(aimimessage&, "&Print")
End Sub
Public Function Aim_GetIMSn() As String
'--This get the name of the person your talking to.
Dim aimimessage&
Dim aimgetcaption$
Dim aimchangetext$
aimimessage& = FindWindow("aim_imessage", vbNullString)
aimgetcaption$ = Get_Caption(aimimessage&)
aimchangetext$ = ReplaceString(aimgetcaption$, " - Instant Message", "")
Aim_GetIMSn = aimchangetext$
End Function
Public Function Aim_GetDCIMSn() As String
'--This get the name of the person when you are
'--directly connected.
Dim aimimessage&
Dim aimgetcaption$
Dim aimchangetext$
aimimessage& = FindWindow("aim_imessage", vbNullString)
aimgetcaption$ = Get_Caption(aimimessage&)
aimchangetext$ = ReplaceString(aimgetcaption$, " - Direct Instant Message", "")
Aim_GetDCIMSn = aimchangetext$
End Function
Public Sub Aim_SendIMLink(ThePerson$, TheUrl$, TheLinkText$)
Call Aim_Im_Send_Normal(ThePerson$, "<a href=""" + TheUrl$ + """>" + TheLinkText$ + "")
End Sub
Public Sub Aim_SendIMLinkthenClose(ThePerson$, TheUrl$, TheLinkText$)
Call Aim_Im_Send_Normal2(ThePerson$, "<a href=""" + TheUrl$ + """>" + TheLinkText$ + "")
End Sub
Public Sub Aim_EditProfile()
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call RunMenuByString(oscarbuddylistwin&, "Edit &Profile...")
End Sub
Public Sub Aim_SaveBuddylist()
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call RunMenuByString(oscarbuddylistwin&, "&Save Buddy List...")
End Sub
Public Sub Aim_LoadBuddylist()
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call RunMenuByString(oscarbuddylistwin&, "&Load Buddy List...")
End Sub
Public Sub AIM_ExitMain()
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call RunMenuByString(oscarbuddylistwin&, "&Close")
End Sub
Public Sub Aim_MaximizeMain()
Dim maximizeim&
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
maximizeim& = ShowWindow(oscarbuddylistwin&, SW_MAXIMIZE)
End Sub
Public Sub Aim_MinimizeMain()
Dim minimizeim&
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
minimizeim& = ShowWindow(oscarbuddylistwin&, SW_MINIMIZE)
End Sub
Public Sub Aim_RestoreMain()
Dim restoreim&
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
restoreim& = ShowWindow(oscarbuddylistwin&, SW_RESTORE)
End Sub
Public Sub Aim_HideMain()
Dim hideim&
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
hideim& = ShowWindow(oscarbuddylistwin&, SW_HIDE)
End Sub
Public Sub Aim_ShowMain()
Dim hideim&
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
hideim& = ShowWindow(oscarbuddylistwin&, SW_SHOW)
End Sub
Public Sub Aim_HideBuddyList()
Dim hideim&
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscartree&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscartree& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tree", vbNullString)
hideim& = ShowWindow(oscartree&, SW_HIDE)
End Sub
Public Sub Aim_ShowBuddyList()
Dim hideim&
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscartree&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscartree& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tree", vbNullString)
hideim& = ShowWindow(oscartree&, SW_SHOW)
End Sub
Public Sub Aim_SendToOpenIM(TheMessage$)
Dim sendbuttonicon1&
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
Dim aimimessage2&
Dim oscariconbtn&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, TheMessage$)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub Aim_SendToOpenIMthenClose(TheMessage$)
Dim aimimessage2&
Dim oscariconbtn&
Dim aimimessage3&
Dim sendbuttonicon1&
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, TheMessage$)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
aimimessage3& = FindWindow("aim_imessage", vbNullString)
Call ShowWindow(aimimessage&, SW_HIDE)
End Sub
Public Sub Aim_ClickChatSend()
Dim button&
Dim sendbuttonicon1&
Dim aimchatwnd&
Dim oscariconbtn&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Function Aim_FindChat() As Long
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Aim_FindChat = aimchatwnd
End Function
Public Sub Aim_SendChatNormal(TheMessage$)
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
Dim sendbuttonicon1&
Dim oscariconbtn&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, TheMessage$)
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub Aim_SendChatLink(TheUrl$, TheLinkText$)
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
Dim sendbuttonicon1&
Dim oscariconbtn&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, "<a href=""" + TheUrl$ + """>" + TheLinkText$ + "")
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub Aim_ClearChat()
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, "")
End Sub
Public Function Aim_GetChatText() As String
Dim aimimessage2$
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
aimimessage2$ = Get_Text(ateclass&)
Aim_GetChatText = aimimessage2$
End Function
Public Function Aim_GetChatTextwithoutHTML() As String
Dim ChatWindow As Long, BorderThing As Long, getit As String
Dim clear As String
ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
BorderThing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
getit$ = Get_Text(BorderThing&)
clear$ = HTML_Remove(getit$)
Aim_GetTextwithoutHTML = clear$
End Function

