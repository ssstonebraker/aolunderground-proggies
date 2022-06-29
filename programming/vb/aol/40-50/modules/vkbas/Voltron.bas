Attribute VB_Name = "Voltron"
'this version  9-8-98

'coded by KRhyME and SkaFia
'|¯|    |¯|\¯\ |¯|    |_|/¯//¯/|¯|\¯\ /¯/|¯|
'| |/¯/ | |/ / | |\¯\   / / | ||_|| ||  |/_/_
'|_|\_\ |_|\_\ |_| |_| /_/  |_|   |_| \______|
'HEAD of the [Voltron Kru]
'Voltron Kru '98
'www.voltronkru.com
'voltronkru@juno.com

'This Bas is the Core of all the VoltronKru (VK)
'bas files. It containes all the need declaires and
'statements to run our other bas files. So if you want
'to use VK_Aol.bas, you must add this bas to the
'project also. You can get all our bas files at
'www.voltronkru.com

'many ideas for this bas came from other Voltron Kru
'members. I would Like to thank SkaFia for all the
'things he did for the series of Bas files.
'Please do not steal our codes without giving us
'credit. I would like to say thank you to KnK for
'making so many files avaible to the public, The makers
'of DiVe32.bas (the first bas i used), Toast, Magus,
'and all the other great programmers out there who
'have infuinced us

'Please join our VB mailing list
'www.voltronkru.com

Global Const META_RECTANGLE = &H41B
Global Const META_SELECTOBJECT = &H12D
Global Const SRCCOPY = &HCC0020
Global Const Pi = 3.14159265359

Global r%       'Result Code from WritePrivateProfileString
Global entry$   'Passed to WritePrivateProfileString
Global iniPath$ 'Path to .ini file

Public majornum
Public minornum
Public timesload As String
Public introyn As String
Public scinyn As String
Public scexyn As String
Public soundyn As String
Public ik As String
Public ima As String
Public removedpeeps
Public soundsyn
Public Chat
Public INTR
Public ICO
Public LB
Public EB
Public UL
'Public AOL
'Public MDi
Public nobut As Integer
Public access As String
Public fState As FormState
Public gFindString As String
Public gFindCase As Integer
Public gFindDirection As Integer
Public gCurPos As Integer
Public gFirstTime As Integer
Public tabclicked As String

Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPI_SETDESKWALLPAPER = 20
Public Const SPIF_SENDWININICHANGE = &H2

Public Const SPI_SCREENSAVERRUNNING = 97
Public Const ThisApp = "MDINote"
Public Const ThisKey = "Recent Files"
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
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_CREATE = &H1
Public Const WM_MDICREATE = &H220

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

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
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const hWnd_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_hWndFIRST = 0
Public Const GW_hWndLAST = 1
Public Const GW_hWndNEXT = 2
Public Const GW_hWndPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
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

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Const EWX_LOGOFF = 0
Public Const EWX_SHUTDOWN = 1
Public Const EWX_REBOOT = 2
Public Const EWX_FORCE = 4
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Public Const SM_CLEANBOOT = 67

'*****************************************************************************************
Type DEVMODE
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
'*****************************************************************************************
Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wId As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
'*****************************************************************************************
Type rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'*****************************************************************************************
Type POINTAPI
   X As Long
   Y As Long
End Type
'*****************************************************************************************
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type
'*****************************************************************************************
Public Type BrowseInfo
     hwndOwner As Long
     pIDLRoot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type
'*****************************************************************************************
Public Type HOSTENT
       hName As Long
       hAliases As Long
       hAddrType As Integer
       hLength As Integer
       hAddrList As Long
End Type
'*****************************************************************************************
Public Type WSADATA
       wversion As Integer
       wHighVersion As Integer
       szDescription(0 To WSADescription_Len) As Byte
       szSystemStatus(0 To WSASYS_Status_Len) As Byte
       iMaxSockets As Integer
       iMaxUdpDg As Integer
       lpszVendorInfo As Long
End Type
'*****************************************************************************************
Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, HostLen&) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
'*****************************************************************************************
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
'*****************************************************************************************
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'*****************************************************************************************
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
'*****************************************************************************************
Declare Function sndplatsound Lib "mmsystem.dll" (ByVal wavfile As Any, ByVal wFlags As Integer) As Integer '<--All one line
'*****************************************************************************************
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'*****************************************************************************************
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'*****************************************************************************************
'Public Declare Function GetWindowThreadProcessId& Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long)as long

Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long
Declare Function EnumDisplaySettings Lib "User32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Declare Function ChangeDisplaySettings Lib "User32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
Declare Function ExitWindowsEx Lib "User32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Private Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "User32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "User32" (lppbKeyState As Byte) As Long
Declare Function GetWindowWord Lib "User32" (ByVal hWnd As Long, ByVal nIndex As Long) As Integer
Declare Function EnableWindow Lib "User32" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, lpRect As rect) As Long
Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As rect) As Long
Declare Function SetWindowRgn Lib "User32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Declare Sub ReleaseCapture Lib "User32" ()
Declare Function ScreenToClient Lib "User32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long
Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Declare Function SelectObject Lib "User32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function FillRect Lib "User32" (ByVal hdc As Integer, lpRect As rect, ByVal hBrush As Integer) As Integer
Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function RedrawWindow Lib "User32" (ByVal hWnd As Long, lprcUpdate As rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "User32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetRect Lib "User32" (lpRect As rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function RegisterWindowMessage& Lib "User32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "User32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "User32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "User32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "User32" (ByVal hWnd As Long) As Long
Declare Function GetParent Lib "User32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function sendmessagebystring Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function CreatePopupMenu Lib "User32" () As Long
Declare Function ShowWindow Lib "User32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "User32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "User32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "User32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "User32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function InsertMenuItem Lib "User32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function AppendMenu Lib "User32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "User32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "User32" (ByVal hMenu%) As Integer
Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function extfn0BD2 Lib "User32" Alias "SendMessageA" (ByVal p1%, ByVal p2%, ByVal p3%, ByVal p4&) As Long
Declare Function extfn0144 Lib "User32" Alias "SendMessageA" (ByVal p1%, ByVal p2%, ByVal p3%, p4&) As Long
'*****************************************************************************************
Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Declare Function SetSystemPowerState Lib "kernel32" (ByVal fSuspend As Long, ByVal fForce As Long) As Long
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
'Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'*****************************************************************************************
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Integer, ByVal X As Integer, ByVal Y As Integer) As Long
Declare Function StretchBlt% Lib "gdi32" (ByVal hdc%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal nSrcWidth%, ByVal nSrcHeight%, ByVal dwRop&)
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Sub HowTo_ScanFileFor()

'scans a file for a string
'you need: 1 button, 1 drive list box,
'          1 dir list box, 1 text box,
'          1 file list box
'To make it work:
'
'in the drive list box :
'    On Error GoTo DriveHandler
'    Dir1.Path = Drive1.Drive: Exit Sub
'    DriveHandler:
'    Drive1.Drive = Dir1.Path
'    Resume Next
'    Exit Sub
'
'in the dir list box:
'    File1.Path = Dir1.Path
'
'in the button:
'    If File1.FileName = "" Then Exit Sub
'    FPath$ = Dir1.Path
'    If right(FPath$, 1) <> "\" Then FPath$ = FPath$ + "\"
'    SelectedFile$ = FPath$ + File1.FileName
'    A_SearchText1% = ScanFor(SelectedFile$, Text1): DoEvents
'    If A_SearchText1% = 0 Then
'    OutCome$ = " "
'    Else
'    OutCome$ = " NOT "
'    End If
'    MsgBox "The String was" + OutCome$ + "Found!"

End Sub



Sub HowTo_PopupMenu()

'The following is how to make a popup menu
'
'using the menu editor, create a menu,
'lets call this example "bob"
'uncheck the visible property....
'in the program, call a popup in a button or on
'the form itself.....PopupMenu (bob)
'
'this makes only the right mouse click call it
'If Button And vbRightButton Then PopupMenu (bob)
End Sub

Sub HowTo_ListSave()
'this is a sub to save a listbox..
'you must have a common dialog on the form

'Sub List_Save(CMD As CommonDialog, Lst As ListBox)

'Dim l0072 As Variant
'Dim l0076 As Variant
'For l0072 = 0 To Lst.ListCount - 1
'l0076 = l0076 & Lst.List(l0072) & Chr(13) & Chr(10)
'Next l0072
'CMD.DialogTitle = "Save SN List"
'CMD.Filter = "Any List (*.lst)|*.lst|"
'CMD.FilterIndex = 2
'CMD.Action = 2
'If CMD.FileTitle = "*.lst" Then Exit Sub
'On Error Resume Next
'Open CMD.FileTitle For Output As #1
'For X = 0 To Lst.ListCount - 1
'Print #1, Lst.List(X)
'Next X
'Close #1

End Sub
Sub HowTo_ListLoad()
'this will load a listbox with info from a file
'you must have a common dialog on the form

'Sub List_Load(CMD As CommonDialog, Lst As ListBox)

'Dim l006C As Integer
'Dim l006E As Variant
'On Error Resume Next
'CMD.Filter = " List Files (*.lst)|*.lst|"
'CMD.FilterIndex = 2
'CMD.Action = 1
'CMD.DialogTitle = "Load SN list"
'If CMD.FileTitle = "" Then Exit Sub
'On Error Resume Next
'Open CMD.FileTitle For Input As #1
'While Not EOF(1)
'Input #1, text$
'DoEvents
'Lst.AddItem text$
'Wend
'Close #1
End Sub


Public Sub Disable_ALT_CTRL_DEL()
'Disables the Crtl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub Enable_ALT_CTRL_DEL()
'Enables the Crtl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Sub Directory_Create(dir)
'This will make a directory
MkDir dir
End Sub

Sub Directory_Delete(dir)
'this deletes a directory
RmDir (dir)
End Sub

Public Sub CenterForm(frmForm As Form)
'this function will center the form given
'on the screen
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub

Public Function FindChildByNum&(hWnd&, num&)
'This Finds The Child Window By Its Order In The Window
'Like Using GetWindow and GW_HWNDNEXT but faster
Static child, indx, nextwnd
child = GetWindow(hWnd&, GW_CHILD)
indx = 1
nextwnd = GetWindow(child, GW_hWndFIRST)
Do While indx < num&
    nextwnd = GetWindow(nextwnd, GW_hWndNEXT)
Loop
FindChildByNum& = nextwnd
Let child = vbNull
Let indx = vbNull
Let nextwnd = vbNull
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
hChild = GetWindow(hChild, GW_hWndNEXT)
i = i + 1
Wend
GetChildCount = i
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Function GetText(child)
'This will get the Text from any window
gettrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = sendmessagebystring(child, 13, gettrim + 1, trimspace$)
GetText = trimspace$
End Function

Sub SetText(win, txt)
'This will send text to a window
thetext% = sendmessagebystring(win, WM_SETTEXT, 0, txt)
End Sub

Function Talk_Backwards(strin As String)
'Returns the string backwards
Let inptxt$ = strin$
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
Talk_Backwards = newsent$
End Function


Sub clock(lbl As label)
'Place this code in  a timer for a digital clock
'On you form..Looks really cool
lbl.Caption = Time
End Sub

Sub CloseWindow(winew)
'This will close a window
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub

Function Talk_Elite(strin As String)
'Returns the string in elite
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

If nextchr$ = "A" Then Let nextchr$ = "Å"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "h"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "K" Then Let nextchr$ = "(«"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "/\/\"
If nextchr$ = "m" Then Let nextchr$ = "‹v›"
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
If nextchr$ = "X" Then Let nextchr$ = "><"
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
Talk_Elite = newsent$

End Function

Sub File_Delete(File)
'deletes a file
Kill (File)
End Sub

Sub File_ReName(sFromLoc As String, sToLoc As String)
'renames a file
Name sFromLoc As sToLoc
End Sub

Function FindChildByClass(ParenthWnd, hWndClassName) As Integer
'finds a child windows by its class
'
'use vSPY to find info about a window
'Get vSPY at www.voltronkru.com

ChildhWnd = FindWindowEx(ParenthWnd, 0, hWndClassName, vbNullString)
FindChildByClass = ChildhWnd
End Function

Function findchildbytitle(ParenthWnd, hWndTitle) As Integer
'finds a child window by its text
'
'use vSPY to find info about a window
'Get vSPY at www.voltronkru.com

ChildhWnd = FindWindowEx(ParenthWnd, 0, vbNullString, hWndTitle)
findchildbytitle = ChildhWnd
End Function

Function FindChildByTitlePartial(parentw, childhand)
'finds a child windows by part of its title
'
'use vSPY to find info about a window
'Get vSPY at www.voltronkru.com

firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitlePartial = 0
bone:
room% = firs%
FindChildByTitlePartial = room%
End Function

Sub ForceShutdown()
'this will force the shutdown of the computer
ForcedShutdown = ExitWindowsEx(EWX_FORCE, 0&)
End Sub

Sub Form_Maximize(frm As Form)
'maximizes the form
frm.WindowState = 2
End Sub

Sub Form_Minimize(frm As Form)
'minimizes the form
frm.WindowState = 1
End Sub

Function GetCaption(hWnd)
'gets a caption of a window
hWndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hWndLength%, 0)
a% = GetWindowText(hWnd, hWndTitle$, (hWndLength% + 1))
GetCaption = hWndTitle$
End Function

Function GetClass(child)
'This will return the Class name of a
'child
'
'use vSPY to find info about a window
'Get vSPY at www.voltronkru.com

Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)
GetClass = Buffer$
End Function

Function GetFromINI(AppName$, KeyName$, FileName$) As String
'example of how to read from an .ini
'iniPath$ = App.Path + "\your.ini"
'Dim a%
'a% = GetFromINI("killers", "45min", iniPath$)
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function

Function WriteINI(AppName$, KeyName$, entry$, iniPath$)
'how to write to an ini file
'iniPath$ = App.Path + "\your.ini"
'entry$ = Check6.Value
'R% = WritePrivateProfileString("killers", "invite", entry$, iniPath$)
'this makes a check value write to the ini
X = WritePrivateProfileString(AppName$, KeyName$, entry$, iniPath$)
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


Function Talk_Hacker(strin As String)
'Function Talk_Hacker(Strin$)
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
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = " " Then Let nextchr$ = " "
If nextchr$ = "]" Then Let nextchr$ = "]"
If nextchr$ = "[" Then Let nextchr$ = "["
Let newsent$ = newsent$ + nextchr$
Loop
Talk_Hacker = newsent$

End Function

Sub HideWindow(hWnd)
'hides a window
X = ShowWindow(hWnd, SW_HIDE)
End Sub

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

Function IntegerToString(tochange As Integer) As String
'This will convert a integer to string
IntegerToString = Str$(tochange)
End Function

Sub NotOnTop(the As Form)
'This will take a form and make it so that
'it does not stay on top of other forms
'U HAVE TO MAKE THE EXE to SEE IT WERK
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub


Function playmusic(FilePath)
'Plays sound files
i% = sndPlaySound(FilePath, 3)

End Function

Sub PreVent()
' Only Allows one version of your prog to run at a time
'Like AOL
If App.PrevInstance Then End
End Sub

Sub printerfuck(numberofpages As String, printermessage As String)
'Makes the printer print a grip of pages
'original code by Toast, revised by KRhyME
Dim HWidth, HHeight, i, Msg
    On Error GoTo ErrorHandler
Msg = printermessage
    For i = 1 To numberofpages
        HWidth = Printer.TextWidth(Msg) / 2
        HHeight = Printer.TextHeight(Msg) / 2
        Printer.CurrentX = Printer.ScaleWidth / 2 - HWidth
        Printer.CurrentY = Printer.ScaleHeight / 2 - HHeight
        Printer.Print Msg & " " & Printer.Page & " of "; numberofpages + " pages!!!! "; ""
Printer.NewPage ' Send new page.
    Next i
    Printer.EndDoc  ' Printing is finished.
    Exit Sub
ErrorHandler:
    
    Exit Sub
End Sub

Sub RestartComputer()
'this will restart the cpmputer
ForcedShutdown = ExitWindowsEx(EWX_REBOOT, 0&)
End Sub

Sub RMBS(ApplicationOfMenu, STringToSearchFor)
'runs a menu by string
SearchString$ = STringToSearchFor
hMenu = GetMenu(ApplicationOfMenu)
Cnt = GetMenuItemCount(hMenu)
For i = 0 To Cnt - 1
DoEvents
PopUphMenu = GetSubMenu(hMenu, i)
Cnt2 = GetMenuItemCount(PopUphMenu)
For O = 0 To Cnt2 - 1
DoEvents
    hMenuID = GetMenuItemID(PopUphMenu, O)
    MenuString$ = String$(100, " ")
    X = GetMenuString(PopUphMenu, hMenuID, MenuString$, 100, 1)
    If InStr(UCase(MenuString$), UCase(SearchString$)) Then
        SendtoID = hMenuID
        GoTo Initiate
    End If
Next O
Next i
Initiate:
X = sendmessagebynum(ApplicationOfMenu, &H111, SendtoID, 0)
End Sub



Sub shutdown()
'this will shut down the computer
StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End Sub

Sub stayontop(frm As Form)
'makes a window stay on top of all others
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function StringToInteger(tochange As String) As Integer
'changes a string to an integer
StringToInteger = tochange
On Error GoTo err1234
Exit Function
err1234:
StringToInteger = ""
End Function

Function sys_timeanddate() As String
'text2.text = "It is Currently " + sys_timeanddate()
sys_timeanddate$ = Format$(Now, "h:mm AM/PM mm-dd-yy")

End Function

Sub timeout(interval)
'This will pause for However many seconds
'your decide, same as pause
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub UnHideWindow(hWnd)
'this will un hide a hiden window
X = ShowWindow(hWnd, SW_SHOW)
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

Sub Click(Button%)
'This clicks a button that you specify

Dim ClickDown%, ClickUp%
ClickDown% = sendmessagebynum(Button%, WM_LBUTTONDOWN, &HD, 0)
ClickUp% = sendmessagebynum(Button%, WM_LBUTTONUP, &HD, 0)
NoFreeze% = DoEvents()
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
Sub DoubleClick(Button%)
'this double clicks a button
Dim DoubleClickNow%
DoubleClickNow% = sendmessagebynum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub
Sub File_ChangeTo(Choice, File$)
'Choices are: Normal, Readonly, Hidden, System, and archive
If Not IFileExists(File$) Then Exit Sub
If LCase$(Choice) = "normal" Then
SetAttr File$, ATTR_NORMAL
ElseIf LCase$(Choice) = "readonly" Then SetAttr File$, ATTR_READONLY
ElseIf LCase$(Choice) = "hidden" Then SetAttr File$, ATTR_HIDDEN
ElseIf LCase$(Choice) = "system" Then SetAttr File$, ATTR_SYSTEM
ElseIf LCase$(Choice) = "archive" Then SetAttr File$, ATTR_ARCHIVE
End If
NoFreeze% = DoEvents()
End Sub
Function File_GetFileName(Prompt As String) As String
'gets the files name
File_GetFileName = LTrim$(RTrim$(UCase$(InputBox$(Prompt, "Enter File Name"))))
End Function
Function File_GetSysIni(section$, Key$)
'gets the system.ini
Dim retval As String, AppName As String, worked As Integer
    retval = String$(255, 0)
    worked = GetPrivateProfileString(section$, Key$, "", retval, Len(retval), "System.ini")
    If worked = 0 Then
        File_GetSysIni = "unknown"
    Else
        File_GetSysIni = Left(retval, worked)
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

Sub File_OpenEXE(File$)
'opens an exe
Openit! = Shell(File$, 1): NoFreeze% = DoEvents()
End Sub
Sub File_RenameDirectory(old$, NewName$)
'renames a directory
If Not IfDirExists(old$) Then Exit Sub
Name old$ As NewName$
NoFreeze% = DoEvents()
End Sub

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
Function RandomNumber(finished)
'generates a random number
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
Sub Text_RandColor(txt As TextBox)
'makes text a randum color
Counter = 0
Do
txt.ForeColor = QBColor(Rnd * 15)
NoFreeze% = DoEvents()
Counter = Counter + 1
Loop Until Counter >= 30
End Sub

Function Text_Spaced(strin$)
'spaces out text
Let inptxt$ = strin$
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
Text_Spaced = newsent$
End Function
Sub Window_Disable(windo%)
'disables a window
X = EnableWindow(windo%, 0)
End Sub
Sub Window_Enable(windo%)
'enables a disabled window
X = EnableWindow(windo%, 1)
End Sub
Function centerstring(before$) As String
'centers a string
    lenstr% = twiplen(before$)
    leftlen% = 6500 - lenstr%
    blanklen% = leftlen% / 2
    numblank% = Int(blanklen% / 85)
    blank$ = ""
    
    For ctr% = 1 To numblank%
        blank$ = blank$ & " "
    Next ctr%

    centerstring = blank$ & before$

End Function

Function twiplen(teststring$) As Integer
'this function is used for for centerstring()
Dim twips As Integer
Dim lenstr As Integer
Dim char As String * 1
Dim charasc As Integer

    lenstr% = Len(teststring$)

    For ctr% = 1 To lenstr%
        char$ = Mid$(teststring$, ctr%, 1)
        charasc% = Asc(char$)
        Select Case charasc%
            Case 32: twips% = twips% + 85
            Case 33: twips% = twips% + 64
            Case 34: twips% = twips% + 106
            Case 35: twips% = twips% + 151
            Case 36: twips% = twips% + 149
            Case 37: twips% = twips% + 261
            Case 38: twips% = twips% + 195
            Case 39: twips% = twips% + 43
            Case 40: twips% = twips% + 85
            Case 41: twips% = twips% + 85
            Case 42: twips% = twips% + 104
            Case 43: twips% = twips% + 172
            Case 44: twips% = twips% + 85
            Case 45: twips% = twips% + 85
            Case 46: twips% = twips% + 85
            Case 47: twips% = twips% + 85
            Case 48: twips% = twips% + 151
            Case 49: twips% = twips% + 151
            Case 50: twips% = twips% + 151
            Case 51: twips% = twips% + 151
            Case 52: twips% = twips% + 151
            Case 53: twips% = twips% + 151
            Case 54: twips% = twips% + 151
            Case 55: twips% = twips% + 151
            Case 56: twips% = twips% + 151
            Case 57: twips% = twips% + 151
            Case 58: twips% = twips% + 85
            Case 59: twips% = twips% + 85
            Case 60: twips% = twips% + 170
            Case 61: twips% = twips% + 170
            Case 62: twips% = twips% + 170
            Case 63: twips% = twips% + 152
            Case 64: twips% = twips% + 284
            Case 65: twips% = twips% + 197
            Case 66: twips% = twips% + 193
            Case 67: twips% = twips% + 197
            Case 68: twips% = twips% + 197
            Case 69: twips% = twips% + 197
            Case 70: twips% = twips% + 175
            Case 71: twips% = twips% + 217
            Case 72: twips% = twips% + 194
            Case 73: twips% = twips% + 64
            Case 74: twips% = twips% + 119
            Case 75: twips% = twips% + 197
            Case 76: twips% = twips% + 147
            Case 77: twips% = twips% + 242
            Case 78: twips% = twips% + 197
            Case 79: twips% = twips% + 217
            Case 80: twips% = twips% + 197
            Case 81: twips% = twips% + 217
            Case 82: twips% = twips% + 197
            Case 83: twips% = twips% + 197
            Case 84: twips% = twips% + 151
            Case 85: twips% = twips% + 197
            Case 86: twips% = twips% + 197
            Case 87: twips% = twips% + 151
            Case 88: twips% = twips% + 197
            Case 89: twips% = twips% + 197
            Case 90: twips% = twips% + 151
            Case 91: twips% = twips% + 85
            Case 92: twips% = twips% + 87
            Case 93: twips% = twips% + 85
            Case 94: twips% = twips% + 104
            Case 95: twips% = twips% + 151
            Case 96: twips% = twips% + 85
            Case 97: twips% = twips% + 151
            Case 98: twips% = twips% + 151
            Case 99: twips% = twips% + 151
            Case 100: twips% = twips% + 151
            Case 101: twips% = twips% + 151
            Case 102: twips% = twips% + 64
            Case 103: twips% = twips% + 151
            Case 104: twips% = twips% + 151
            Case 105: twips% = twips% + 64
            Case 106: twips% = twips% + 64
            Case 107: twips% = twips% + 151
            Case 108: twips% = twips% + 64
            Case 109: twips% = twips% + 242
            Case 110: twips% = twips% + 151
            Case 111: twips% = twips% + 151
            Case 112: twips% = twips% + 151
            Case 113: twips% = twips% + 151
            Case 114: twips% = twips% + 85
            Case 115: twips% = twips% + 151
            Case 116: twips% = twips% + 85
            Case 117: twips% = twips% + 151
            Case 118: twips% = twips% + 105
            Case 119: twips% = twips% + 196
            Case 120: twips% = twips% + 151
            Case 121: twips% = twips% + 151
            Case 122: twips% = twips% + 151
            Case 123: twips% = twips% + 85
            Case 124: twips% = twips% + 64
            Case 125: twips% = twips% + 85
            Case 126: twips% = twips% + 170
            Case 127: twips% = twips% + 217
            Case 128: twips% = twips% + 217
            Case 129: twips% = twips% + 217
            Case 130: twips% = twips% + 66
            Case 131: twips% = twips% + 151
            Case 132: twips% = twips% + 85
            Case 133: twips% = twips% + 283
            Case 134: twips% = twips% + 151
            Case 135: twips% = twips% + 151
            Case 136: twips% = twips% + 85
            Case 137: twips% = twips% + 311
            Case 138: twips% = twips% + 196
            Case 139: twips% = twips% + 85
            Case 140: twips% = twips% + 285
            Case 141: twips% = twips% + 217
            Case 142: twips% = twips% + 217
            Case 143: twips% = twips% + 217
            Case 144: twips% = twips% + 217
            Case 145: twips% = twips% + 66
            Case 146: twips% = twips% + 66
            Case 147: twips% = twips% + 85
            Case 148: twips% = twips% + 85
            Case 149: twips% = twips% + 103
            Case 150: twips% = twips% + 75
            Case 151: twips% = twips% + 141
            Case 152: twips% = twips% + 85
            Case 153: twips% = twips% + 283
            Case 154: twips% = twips% + 151
            Case 155: twips% = twips% + 85
            Case 156: twips% = twips% + 264
            Case 157: twips% = twips% + 217
            Case 158: twips% = twips% + 217
            Case 159: twips% = twips% + 196
            Case 160: twips% = twips% + 85
            Case 161: twips% = twips% + 64
            Case 162: twips% = twips% + 151
            Case 163: twips% = twips% + 151
            Case 164: twips% = twips% + 151
            Case 165: twips% = twips% + 151
            Case 166: twips% = twips% + 64
            Case 167: twips% = twips% + 151
            Case 168: twips% = twips% + 85
            Case 169: twips% = twips% + 217
            Case 170: twips% = twips% + 85
            Case 171: twips% = twips% + 151
            Case 172: twips% = twips% + 170
            Case 173: twips% = twips% + 85
            Case 174: twips% = twips% + 217
            Case 175: twips% = twips% + 151
            Case 176: twips% = twips% + 103
            Case 177: twips% = twips% + 151
            Case 178: twips% = twips% + 85
            Case 179: twips% = twips% + 85
            Case 180: twips% = twips% + 85
            Case 181: twips% = twips% + 151
            Case 182: twips% = twips% + 151
            Case 183: twips% = twips% + 85
            Case 184: twips% = twips% + 85
            Case 185: twips% = twips% + 85
            Case 186: twips% = twips% + 103
            Case 187: twips% = twips% + 151
            Case 188: twips% = twips% + 236
            Case 189: twips% = twips% + 236
            Case 190: twips% = twips% + 236
            Case 191: twips% = twips% + 170
            Case 192: twips% = twips% + 196
            Case 193: twips% = twips% + 196
            Case 194: twips% = twips% + 196
            Case 195: twips% = twips% + 196
            Case 196: twips% = twips% + 196
            Case 197: twips% = twips% + 196
            Case 198: twips% = twips% + 283
            Case 199: twips% = twips% + 196
            Case 200: twips% = twips% + 196
            Case 201: twips% = twips% + 196
            Case 202: twips% = twips% + 196
            Case 203: twips% = twips% + 196
            Case 204: twips% = twips% + 66
            Case 205: twips% = twips% + 66
            Case 206: twips% = twips% + 66
            Case 207: twips% = twips% + 66
            Case 208: twips% = twips% + 196
            Case 209: twips% = twips% + 196
            Case 210: twips% = twips% + 217
            Case 211: twips% = twips% + 217
            Case 212: twips% = twips% + 217
            Case 213: twips% = twips% + 217
            Case 214: twips% = twips% + 217
            Case 215: twips% = twips% + 170
            Case 216: twips% = twips% + 217
            Case 217: twips% = twips% + 196
            Case 218: twips% = twips% + 196
            Case 219: twips% = twips% + 196
            Case 220: twips% = twips% + 196
            Case 221: twips% = twips% + 196
            Case 222: twips% = twips% + 196
            Case 223: twips% = twips% + 196
            Case 224: twips% = twips% + 151
            Case 225: twips% = twips% + 151
            Case 226: twips% = twips% + 151
            Case 227: twips% = twips% + 151
            Case 228: twips% = twips% + 151
            Case 229: twips% = twips% + 151
            Case 230: twips% = twips% + 264
            Case 231: twips% = twips% + 151
            Case 232: twips% = twips% + 151
            Case 233: twips% = twips% + 151
            Case 234: twips% = twips% + 151
            Case 235: twips% = twips% + 151
            Case 236: twips% = twips% + 66
            Case 237: twips% = twips% + 66
            Case 238: twips% = twips% + 66
            Case 239: twips% = twips% + 66
            Case 240: twips% = twips% + 151
            Case 241: twips% = twips% + 151
            Case 242: twips% = twips% + 151
            Case 243: twips% = twips% + 151
            Case 244: twips% = twips% + 151
            Case 245: twips% = twips% + 151
            Case 246: twips% = twips% + 151
            Case 247: twips% = twips% + 151
            Case 248: twips% = twips% + 151
            Case 249: twips% = twips% + 151
            Case 200: twips% = twips% + 151
            Case 201: twips% = twips% + 151
            Case 202: twips% = twips% + 151
            Case 203: twips% = twips% + 151
            Case 204: twips% = twips% + 151
            Case 205: twips% = twips% + 151
        
        End Select
    Next ctr%

    twiplen = twips%

End Function


Sub Win95_ClickStartButton()

'clicks the start button
wind% = FindWindow("Shell_TrayWnd", vbNullString)
btn% = FindChildByClass(wind%, "Button")
SendNow% = sendmessagebynum(btn%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = sendmessagebynum(btn%, WM_LBUTTONUP, &HD, 0)
End Sub

Function Talk_In_CAPS(strin As String) As String

'this will capitalize any string"
L% = Len(strin$)
numspc% = 0
Do While numspc% <= L%
    Let numspc% = numspc% + 1
    Let nextchr$ = Mid$(strin$, numspc%, 1)
    If nextchr$ = "i" Or nextchr$ = "I" Then final$ = final$ & "i" Else final$ = final$ & UCase(nextchr$)
    Loop
Talk_In_CAPS$ = final$
End Function

Function Talk_InsChr(ByVal strin As String, ByVal InsMe As String)
'This function Inserts a Character after every character.
'
'Example:
'
'text2.text = Talk_InsChr("Change Me!",  ".")
'
'That would return "C.h.a.n.g.e. .M.e.!."

Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ <> " " Then Let nextchr$ = nextchr$ + InsMe
Let newsent$ = newsent$ + nextchr$
Loop
LastChr = Len(newsent$)
Talk_InsChr = Left(newsent$, LastChr - 1)
End Function


Function ScanFor(sFile$, ByVal sWhat$)
'scans a file for a string.....
'see HowTo_ScanFileFor for directions.....

Dim VariantA As Variant
Dim VariantB As Variant
Dim VariantC As Variant
Dim VariantD As Variant
Dim SingleA As Single
Dim StringA As String
Dim EnterKey As String

On Error Resume Next
Open sFile$ For Binary As #1
    EnterKey$ = Chr$(13) + Chr$(10)
    Msg$ = ""
    VariantA = LOF(1)
    VariantB = VariantA
    VariantC = 1

    If VariantB > 32000 Then
        VariantD = 32000
    ElseIf VariantB = 0 Then
        VariantD = 1
    Else
        VariantD = VariantB
    End If

    StringA$ = String$(VariantD, " ")
    Get #1, VariantC, StringA$

    SingleA! = InStr(1, StringA$, sWhat$, 1)    'Our Search String

    If SingleA! Then
        ScanFor = 0                             'String was Found
    Else
        ScanFor = 1                             'String was Not Found
    End If
Close #1
End Function

Sub lab_RandColor(label As label)
'makes text a randum color
Counter = 0
Do
label.ForeColor = QBColor(Rnd * 15)
NoFreeze% = DoEvents()
Counter = Counter + 1
Loop Until Counter >= 30
End Sub


Public Function GetListIndex(LB As ListBox, txt As String) As Integer
'finds the index of a specific word
Dim Index As Integer
With LB
For Index = 0 To .ListCount - 1
If .List(iIndex) = txt Then
GetListIndex = Index
Exit Function
End If
Next Index
End With
GetListIndex = -2
End Function

Sub SetCheckBoxToFalse(win%)
'This will set any checkbox's value to equal false
Check% = sendmessagebynum(win%, BM_SETCHECK, False, 0&)
End Sub

Sub SetCheckBoxToTrue(win%)
'This will set any checkbox's value to equal true
Check% = sendmessagebynum(win%, BM_SETCHECK, True, 0&)
End Sub

Function Window_ChangeCaption(windo, txt)
'This will change the caption of any window that you
'tell it to as long as it is a valid window
text% = sendmessagebystring(windo, WM_SETTEXT, 0, txt)
End Function
Function ReplaceText(text, charfind, charchange)
'Replaces a word with a new word
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

Function TrimSpaces(text)
'removes all spaces in a text
If InStr(text, " ") = 0 Then
TrimSpaces = text
Exit Function
End If

For trimspace = 1 To Len(text)
thechar$ = Mid(text, trimspace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next trimspace

TrimSpaces = thechars$
End Function

Function TrimReturns(thetext)
'removes returns
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function


Function AddListToString(thelist As ListBox)
'This will take a list box and add the
'entrys to a string with a comma to
'separate them.
For DoList = 0 To thelist.ListCount - 1
AddListToString = AddListToString & thelist.List(DoList) & ","
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function

Sub AddStringToList(theitems, thelist As ListBox)
'This will take a string with multiple
'variables separated by commas and add
'them to a list

If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
DoEvents
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
thelist.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub

Sub cdcrazy()
'open and close the cd tray over and over
Do
Call opencd
timeout 0.75
Call closecd
DoEvents
Loop
End Sub

Sub closecd()
'close the cd tray
retvalue = MciSendString("set CDAudio door closed", vbNullString, 0, 0)

End Sub


Sub HideMouse()

'Makes the mouse arrow dissaper
Hid$ = ShowCursor(False)

End Sub

Sub ShowMouse()
'shows the mouse after its hidden
Hid$ = ShowCursor(True)
End Sub

Sub HideStartmenu()
' THIS is funny makes it so they lose the start menu
' and cant get it back unless you run showstartmenu
C% = FindWindow("Shell_TrayWnd", vbNullString)
a = ShowWindow(C%, SW_HIDE)

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

Sub opencd()
'this opens the cd tray

retvalue = MciSendString("set CDAudio door open", vbNullString, 0, 0)

End Sub

Sub ShowStartmenu()
'will show the start button after its hidden

C% = FindWindow("Shell_TrayWnd", vbNullString)
a = ShowWindow(C%, SW_SHOW)

End Sub

Sub Suspender()
'If your computer supports it this will
' Make your computer Go into Suspend Mode

 a$ = SetSystemPowerState(Suspend, force)

End Sub


Sub Combo_AddTo(itm As String, lst As ComboBox)
'adds text to a combo only if it is not
'in the list...prevents dupes

If lst.ListCount = 0 Then lst.AddItem itm: Exit Sub
Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub

Sub List_ADD(itm As String, lst As ListBox)
'adds text to a list only if it is not
'in the list...prevents dupes

Dim xx
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


Public Sub FormRightCorner(frmForm As Form)
'This will center your form in the upper right
'of the users screen
   With frmForm
      .Left = (Screen.Width - .Width) / 1
      .Top = (Screen.Height - .Height) / 2000
   End With
End Sub

Sub FormSizeToWindow(frm As Form, win%)
'sizes a form to a window
Dim wndRect As rect, lRet As Long
lRet = GetWindowRect(win%, wndRect)
With frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub

Function SetWallpaper(sFileName As String) As Long
'this will set an image as the windows wallpaper
SetWallpaper = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, sFileName, SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function

Sub Progress(pb As Control, ByVal percent)
' this is a progress bar
' example:  call Progress(form1.picture1,33)

' example2: this will make the percent bar
'           grow to 100%
    
    'For Y = 0 To 100 Step 1
    'Call Progress(Picture1, Y)
    'Call TimeOut(0.00001)
    'Next Y

Dim num$
If Not pb.AutoRedraw Then
    pb.AutoRedraw = -1
    End If
    pb.Cls
    pb.ScaleWidth = 100
    pb.DrawMode = 10
    num$ = Format$(percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    pb.Print num$
    pb.Line (0, 0)-(percent, pb.ScaleHeight), , BF
    pb.Refresh
End Sub

Sub SendCharNum(win, chars)
E = sendmessagebynum(win, WM_CHAR, chars, 0)
End Sub

Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
'This browses your directories...the following
' is an example of how to use it....

'Dim strResFolder As String
'strResFolder = BrowseForFolder(hWnd, "Please select a folder.")
'If strResFolder = "" Then
'    Call MsgBox("The Cancel button was pressed.", vbExclamation)
'Else
'    Call MsgBox("The folder " & strResFolder & " was selected.", vbExclamation)
'End If

     
    'declare variables to be used
     Dim iNull As Integer
     Dim lpIDList As Long
     Dim lResult As Long
     Dim sPath As String
     Dim udtBI As BrowseInfo

    'initialise variables
     With udtBI
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With

    'Call the browse for folder API
     lpIDList = SHBrowseForFolder(udtBI)
     
    'get the resulting string path
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If

    'If cancel was pressed, sPath = ""
     BrowseForFolder = sPath
End Function

Sub Draw3DBorder(C As Control, iLook As Integer)
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

    iOldScaleMode = C.Parent.ScaleMode
    C.Parent.ScaleMode = PIXELS
    C.Parent.Line (C.Left, C.Top - 1)-(C.Left + C.Width, C.Top - 1), QBColor(iFirstColor)
    C.Parent.Line (C.Left - 1, C.Top)-(C.Left - 1, C.Top + C.Height), QBColor(iFirstColor)
    C.Parent.Line (C.Left + C.Width, C.Top)-(C.Left + C.Width, C.Top + C.Height), QBColor(iSecondColor)
    C.Parent.Line (C.Left, C.Top + C.Height)-(C.Left + C.Width, C.Top + C.Height), QBColor(iSecondColor)
    C.Parent.ScaleMode = iOldScaleMode
End Sub


Sub LoadFileInTextbox(FilePath As String, text As TextBox)
'this will load a file into a textbox

Dim a As String
    Open FilePath For Input As 1
    a = Input(LOF(1), 1)
    Close 1
    text = a
End Sub


Sub Form_Falling(frm As Form, steps As Integer)

'this makes a form move into the lower right hand
'corner of the screen
' revised by KRhyME (form leaves no trail now)
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
DoEvents
    frm.Move frm.Left + X, frm.Top + Y
Loop Until (frm.Left >= (Screen.Width - frm.Width)) Or (frm.Top >= (Screen.Height - frm.Height))
frm.Left = Screen.Width - frm.Width
frm.Top = Screen.Height - frm.Height
frm.BackColor = BgColor
For X = 0 To frm.Count - 1
frm.Controls(X).Visible = True
Next X
End Sub

Function LineFromText(text$, theline As Integer)

'gets a line from a file
theview$ = text$
For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
C = C + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If theline = C Then GoTo Saturn
thechars$ = ""
End If

Next FindChar
Exit Function
Saturn:
thechatext$ = ReplaceText(thechatext$, Chr(13), "")
thechatext$ = ReplaceText(thechatext$, Chr(10), "")
LineFromText = thechatext$


End Function

Sub WriteToLog(WHAT As String, FilePath As String)

'writes a log file
If FilePath = "" Then Exit Sub
F% = FreeFile
Open FilePath For Binary Access Write As F%
p$ = WHAT & Chr(10)
Put #1, LOF(1) + 1, p$
Close F%
End Sub

Function Encrypt_Decrypt(text, types)
'to encrypt, example:
'encrypted$ = Encrypt_Decrypt("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = Encrypt_Decrypt("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(text)
DoEvents
If types = 0 Then
Current$ = Asc(Mid(text, God, 1)) - 1
Else
Current$ = Asc(Mid(text, God, 1)) + 1
End If
Process$ = Process$ & Chr(Current$)
Next God

Encrypt_Decrypt = Process$
End Function


Sub TurnOnScreenSaver()
'activates the screen saver
'press a key to turn the screen
'saver off
       Dim lResult As Long
       Const WM_SYSCOMMAND = &H112
       Const SC_SCREENSAVE = &HF140
       lResult = SendMessage(-1, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub

Function GetCapslock() As Boolean
' Return the Capslock toggle.
GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)
End Function

Function GetNumlock() As Boolean
' Return the Numlock toggle.
GetNumlock = CBool(GetKeyState(vbKeyNumlock) And 1)
End Function

Function GetScrollLock() As Boolean
' Return the ScrollLock toggle.
GetScrollLock = CBool(GetKeyState(vbKeyScrollLock) And 1)
End Function


Sub SetCapslock(Value As Boolean)
'true = on
'false = off
       Call SetKeyState(vbKeyCapital, Value)
End Sub


Sub SetNumlock(Value As Boolean)
'true = on
'false = off

       Call SetKeyState(vbKeyNumlock, Value)
End Sub

Sub SetScrollLock(Value As Boolean)
'true = on
'false = off

Call SetKeyState(vbKeyScrollLock, Value)
End Sub


Private Sub SetKeyState(intKey As Integer, fTurnOn As Boolean)
' Retrieve the keyboard state, set the particular
' key in which you're interested, and then set
' the entire keyboard state back the way it
' was, with the one key altered.
       Dim abytBuffer(0 To 255) As Byte
       GetKeyboardState abytBuffer(0)
       abytBuffer(intKey) = CByte(Abs(fTurnOn))
       SetKeyboardState abytBuffer(0)
End Sub


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

Sub NuttyKeyboard(Pausetime As Integer)
'this will turn the capslock, scroll lock, and
'numberlock on and off over and over
'......................pretty funny huh
Do
DoEvents
Call SetCapslock(True)
Call SetScrollLock(True)
Call SetNumlock(True)
Call timeout(Pausetime)
Call SetCapslock(False)
Call SetScrollLock(False)
Call SetNumlock(False)
Call timeout(Pausetime)
Loop
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
CR$ = Chr$(13) + Chr$(10)
   TWidth% = Screen.Width \ Screen.TwipsPerPixelX
   THeight% = Screen.Height \ Screen.TwipsPerPixelY
   GetScreenSETTINGS = CR$ + CR$ + Str$(TWidth%) + " x" + Str$(THeight%)
End Function

Sub ChangeScreenSETTINGS(WHAT As Integer)
'this changes the settings on the moniter
'www.voltronkru.com is best viewed at 800 x 600

'1 = 640 x 480
'2 = 800 x 600

Dim DevM As DEVMODE
'Get the info into DevM
erg& = EnumDisplaySettings(0&, 0&, DevM)
'We don't change the colordepth, because a
'rebot will be necessary

DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT 'Or DM_BITSPERPEL
If WHAT = 1 Then
    DevM.dmPelsWidth = 640 'ScreenWidth
    DevM.dmPelsHeight = 480 'ScreenHeight
   'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)
End If

If WHAT = 2 Then
    DevM.dmPelsWidth = 800 'ScreenWidth
    DevM.dmPelsHeight = 600 'ScreenHeight
   'DevM.dmBitsPerPel = 32 (could be 8, 16, 32 or even 4)

End If
'Now change the display and check if possible

erg& = ChangeDisplaySettings(DevM, CDS_TEST)

'Check if succesfull

Select Case erg&
Case DISP_CHANGE_RESTART
    an = MsgBox("You have to reboot for changes to take effect. Reboot now ?", vbYesNo + vbSystemModal, "Info")
If an = vbYes Then
    erg& = ExitWindowsEx(EWX_REBOOT, 0&)
End If
    Case DISP_CHANGE_SUCCESSFUL
    erg& = ChangeDisplaySettings(DevM, CDS_UPDATEREGISTRY)
    'MsgBox "Everything's ok", vbOKOnly + vbSystemModal, "It worked!"
Case Else
    MsgBox "The selected Video Mode is not supported on this machine", vbOKOnly + vbSystemModal, "Error"
End Select
End Sub

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

Sub WindowsBootMode(labely As label)
'TELLS HOW WINDOWS IS RUNNING
Select Case GetSystemMetrics(SM_CLEANBOOT)
Case 1: labely = "Windows is running in Safe Mode."
Case 2: labely = "Windows is running in Safe Mode with Network support."
Case Else: labely = "Windows is running normally."
End Select
End Sub


Function GetmyIPAddress()
'this will retrive your IP address.
'you must be connected to the internet for
'this to work

'If Text = 0 Then
'   Exit Sub
'Else
'   Text = GetmyIPAddress
'End If

   '     'Sockets Initialize
Dim WSAD As WSADATA
Dim iReturn As Integer
Dim sLowByte As String, sHighByte As String, sMsg As String
iReturn = WSAStartup(WS_VERSION_REQD, WSAD)

If iReturn <> 0 Then
       MsgBox "Winsock.dll is not responding."
       GetmyIPAddress = 0
End

End If

If lobyte(WSAD.wversion) < WS_VERSION_MAJOR _
Or (lobyte(WSAD.wversion) = WS_VERSION_MAJOR _
And hibyte(WSAD.wversion) < WS_VERSION_MINOR) Then
sHighByte = Trim$(Str$(hibyte(WSAD.wversion)))
sLowByte = Trim$(Str$(lobyte(WSAD.wversion)))
sMsg = "Windows Sockets version " & sLowByte & "." & sHighByte
sMsg = sMsg & " is not supported by winsock.dll "
MsgBox sMsg
GetmyIPAddress = 0
End
End If

If WSAD.iMaxSockets < MIN_SOCKETS_REQD Then
sMsg = "This application requires a minimum of "
sMsg = sMsg & Trim$(Str$(MIN_SOCKETS_REQD)) _
& " supported sockets."
MsgBox sMsg
GetmyIPAddress = 0
End
End If
    
       Dim hostname As String * 256
       Dim hostent_addr As Long
       Dim host As HOSTENT
       Dim hostip_addr As Long
       Dim temp_ip_address() As Byte
       Dim i As Integer
       Dim ip_address As String

              If gethostname(hostname, 256) = SOCKET_ERROR Then
                     MsgBox "Windows Sockets error " & Str(WSAGetLastError())
                     GetmyIPAddress = 0
                     Exit Function
              Else
                     hostname = Trim$(hostname)
              End If

       hostent_addr = gethostbyname(hostname)

              If hostent_addr = 0 Then
                     MsgBox "Winsock.dll is not responding. Make sure you are connected to the internet."
                     GetmyIPAddress = 0
                     Exit Function
              End If

       RtlMoveMemory host, hostent_addr, LenB(host)
       RtlMoveMemory hostip_addr, host.hAddrList, 4
       ReDim temp_ip_address(1 To host.hLength)
       RtlMoveMemory temp_ip_address(1), hostip_addr, host.hLength

              For i = 1 To host.hLength
                     ip_address = ip_address & temp_ip_address(i) & "."
              Next

       ip_address = Mid$(ip_address, 1, Len(ip_address) - 1)
       'MsgBox hostname
       
              Dim lReturn As Long
       lReturn = WSACleanup()

              If lReturn <> 0 Then
                     MsgBox "Socket error " & Trim$(Str$(lReturn)) & " occurred" & "in Cleanup "
              GetmyIPAddress = 0
End If
 
 GetmyIPAddress = ip_address
End Function

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
Function hibyte(ByVal wParam As Integer)
'used for getting your ip address
       hibyte = wParam \ &H100 And &HFF&
End Function
Function lobyte(ByVal wParam As Integer)
'used for getting your ip address
       lobyte = wParam And &HFF&
End Function


