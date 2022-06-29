Attribute VB_Name = "FrENzY32"
'this should only be seen by izekial if you
'have this file please delete it! HAHAHAHAHA
'if you have this i doubt you will, lol
Option Explicit

Public Declare Function auxSetVolume Lib "WinMM.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function auxGetVolume Lib "WinMM.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function CreateWindow Lib "user32" Alias "CreateWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function CreateWindowEx Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wflags As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetCapture Lib "user32" () As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpstring As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function mciSendString Lib "WinMM.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wflags As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "WinMM.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Window Messages
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_MOVE = &HF012
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOUSEMOVE = &H200
Public Const WM_CLEAR = &H303

'Combo Box Functions
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B

'hWnd Functions
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

'Show Window Functions
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1

'Sound Functions
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SND_LOOP = &H8


'Get Window Word Functions
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

'Virtual Key Statements
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

'Phader Color Presets
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

'Processor Types
Public Const PROCESSOR_INTEL_386 = 386
Public Const PROCESSOR_INTEL_486 = 486
Public Const PROCESSOR_INTEL_PENTIUM = 586
Public Const PROCESSOR_MIPS_R4000 = 4000
Public Const PROCESSOR_ALPHA_21064 = 21064

'Menu Functions
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_POPUP = &H10&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&

'Key Presets
Public Const ENTER_KEY = 13

'Button Messages
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

'List Box Functions
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
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

'Notify Icon Functions
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

'Windows Version Functions
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

'Byte Functions
Public Const MAX_DEFAULTCHAR = 2
Public Const MAX_LEADBYTES = 12

'Types
Public Type SYSTEM_INFO
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

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Public Type CPINFO
        MaxCharSize As Long
        DefaultChar(MAX_DEFAULTCHAR - 1) As Byte
        LeadByte(MAX_LEADBYTES - 1) As Byte
End Type

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Type COLORRGB
    Red As Long
    Green As Long
    Blue As Long
End Type

Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Sub AOLChatCmdEnterRoom()
    Dim TheLEN As Long
    If LCase(AOLChatLine$) Like LCase(".enterroom*") Then
        TheLEN& = Len(AOLChatLine)
        Call AOLEnterPrivateRoom(Right(AOLChatLine, (TheLEN& - 10)))
    End If
End Sub

Public Sub AOLChatCmdIMStatus()
    If LCase(AOLChatLine$) Like LCase(".ims off*") Then
        Call AOLInstantMessage("$IM_Off", " ")
    End If
    If LCase(AOLChatLine$) Like LCase(".ims on*") Then
        Call AOLInstantMessage("$IM_On", " ")
    End If
End Sub
Public Sub AOLChatCmdIgnoreSN()
    Dim thetext As String, TheLEN As Long
    If LCase(AOLChatLine$) Like LCase(".ignore*") Then
        TheLEN& = Len(AOLChatLine$)
        thetext$ = Right(AOLChatLine$, TheLEN - 8)
        Call ChatIgnoreByName(thetext$)
    End If
End Sub
Public Sub AOLChatCmdKeyword()
    Dim thetext As String, TheLEN As Long
    If LCase(AOLChatLine$) Like LCase(".keyword*") Then
        TheLEN& = Len(AOLChatLine)
        thetext$ = Right(AOLChatLine, TheLEN - 8)
        Call AOLKeyWord(thetext$)
    End If
End Sub
Public Sub AOLChatCmdSendIM()
    Dim TheLEN&, TempString$, TheColon&, ToWho$, Message$
    If LCase(AOLChatLine$) Like ".sendim*" Then
        TheLEN& = Len(AOLChatLine$) - 8
        TempString$ = Right(AOLChatLine$, TheLEN&)
        TheColon& = InStr(TempString$, ":")
        ToWho$ = Left(TempString$, TheColon& - 1)
        Message$ = Mid(TempString$, TheColon& + 1, Len(TempString$) - TheColon& + 1)
    End If
    Call AOLInstantMessage(ToWho$, Message$)
End Sub
Public Sub AOLChatCmdSendMail()
    Dim TheLEN&, TempString$, TheColon1&, TheColon2&, ToWho$, Message$, subject$
    If LCase(AOLChatLine$) Like ".sendmail*" Then
        TheLEN& = Len(AOLChatLine$) - 10
        TempString$ = Right(AOLChatLine$, TheLEN&)
        TheColon1& = InStr(TempString$, ":")
        TheColon2& = InStr(TheColon1& + 1, TempString$, ":")
        ToWho$ = Left(TempString$, TheColon1& - 1)
        subject$ = Mid(TempString$, TheColon1& + 1, TheColon2& - TheColon1& - 1)
        Message$ = Mid(TempString$, Len(ToWho$) + 1 + Len(subject$) + 2, Len(TempString$) - Len(ToWho$) + Len(subject$) - 1)
    End If
    Call AOLMailSend(ToWho$, subject$, Message$)
    DoEvents
End Sub
Public Sub NetZeroKillAD(NetZeroDir As String)
    Dim TheDir$
    If Dir(NetZeroDir$) = "" Then Exit Sub
    If Right(NetZeroDir$, 1) = "\" Then NetZeroDir$ = Left(NetZeroDir$, Len(NetZeroDir$) - 1)
    TheDir$ = NetZeroDir$ & "\Bin\"
    Call Kill(TheDir$ & "JdbcOdbc.dll")
    DoEvents
    Call Kill(TheDir$ & "jpeg.dll")
    DoEvents
    Call Kill(TheDir$ & "jre.exe")
    DoEvents
    Call Kill(TheDir$ & "jrew.exe")
    DoEvents
    Call Kill(TheDir$ & "math.dll")
    DoEvents
    Call Kill(TheDir$ & "mmedia.dll")
    DoEvents
    Call Kill(TheDir$ & "net.dll")
    DoEvents
    Call Kill(TheDir$ & "rmiregistry.exe")
    DoEvents
    Call Kill(TheDir$ & "symcjit.dll")
    DoEvents
    Call Kill(TheDir$ & "sysresource.dll")
    DoEvents
End Sub
Public Sub AOLChatCmdIMIgnore()
    Dim thetext As String
    If AOLChatLine Like LCase(".ignore") Then
        thetext$ = Left(AOLChatLine, 7)
        Call AOLInstantMessage("$IM_Off " & thetext$, " ")
    End If
End Sub

Public Function AOLChatLineSN() As String
    Dim Room As Long, Rich As Long, TheChatLine As String
    Room& = AOLFindChatRoom
    Rich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
    TheChatLine$ = hWndText(Rich&)
    TheChatLine$ = ReturnStringsLastLine(TheChatLine$)
    If InStr(TheChatLine$, ":") = 0 Then
        AOLChatLineSN$ = ""
        Exit Function
    End If
    AOLChatLineSN$ = Left(TheChatLine$, InStr(TheChatLine$, ":") - 1)
End Function

Public Function ReturnStringsLastLine(thestring As String) As String
    Dim Line As Long, Index As Long
    Dim Spot1 As Long, Spot2 As Long, theline As String
    If ReturnStringsLineCount(thestring$) = 1 Then ReturnStringsLastLine$ = thestring$: Exit Function
    Line& = ReturnStringsLineCount(thestring$) - 1
    Spot1& = InStr(thestring$, Chr(13))
    For Index& = 1 To Line&
        Spot2& = Spot1&
        Spot1& = InStr(Spot2& + 1, thestring$, Chr(13))
    Next Index&
    If Spot1& = 0 Then
        Spot1& = Len(thestring$)
    End If
    theline$ = Mid(thestring$, Spot2&, Spot1& - Spot2& + 1)
    theline$ = ReplaceString(theline$, Chr(13), "")
    theline$ = ReplaceString(theline$, Chr(10), "")
    ReturnStringsLastLine$ = theline$
End Function
Public Function ReturnStringsLineCount(thestring As String) As Long
    Dim Spot As Long, Count As Long
    If Len(thestring$) = 0 Then ReturnStringsLineCount = 0
    Spot& = InStr(thestring$, Chr(13))
    If Spot& <> 0& Then
        ReturnStringsLineCount& = 1
        Do
            Spot& = InStr(Spot + 1, thestring$, Chr(13))
            If Spot& <> 0& Then
                ReturnStringsLineCount& = ReturnStringsLineCount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    ReturnStringsLineCount& = ReturnStringsLineCount& + 1
End Function
Public Function AOLChatSpiral(Text As String)
    Dim theInt As Long, TheLEN As Long
    theInt& = 1
    TheLEN& = Len(Text$)
    Do
        theInt& = theInt& + 1
        Call AOLChatSend(Left(Text$, theInt&))
    Loop Until theInt& = TheLEN&
End Function
Public Sub AOLBlockBuddy(SN As String)
    Dim AOL As Long, MDI As Long, BuddyList As Long
    Dim Icon As Long, SetupScreen As Long, PrivacyPref As Long
    Dim PrivacyWin As Long, Block As Long, Who As Long, SetWho As Long
    Dim Index As Long, Edit As Long, Save As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDI Client", vbNullString)
    BuddyList& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
    Icon& = FindWindowEx(BuddyList&, 0&, "_AOL_ICON", vbNullString)
    Icon& = FindWindowEx(BuddyList&, Icon&, "_AOL_ICON", vbNullString)
    Icon& = FindWindowEx(BuddyList&, Icon&, "_AOL_ICON", vbNullString)
    ClickIcon (Icon&)
    Do
        SetupScreen& = FindWindowEx(MDI&, 0&, vbNullString, AOLUserSN & "'s Buddy Lists")
    Loop Until SetupScreen& <> 0
    PrivacyPref& = FindWindowEx(SetupScreen&, 0&, "_AOL_ICON", vbNullString)
    PrivacyPref& = FindWindowEx(SetupScreen&, PrivacyPref&, "_AOL_ICON", vbNullString)
    PrivacyPref& = FindWindowEx(SetupScreen&, PrivacyPref&, "_AOL_ICON", vbNullString)
    PrivacyPref& = FindWindowEx(SetupScreen&, PrivacyPref&, "_AOL_ICON", vbNullString)
    PrivacyPref& = FindWindowEx(SetupScreen&, PrivacyPref&, "_AOL_ICON", vbNullString)
    ClickIcon PrivacyPref&
    Call hWndClose(SetupScreen&)
    DoEvents
    Do
        PrivacyWin& = FindWindowEx(MDI&, 0&, vbNullString, "Privacy Preferences")
    Loop Until PrivacyWin& <> 0
    Block& = FindWindowEx(PrivacyWin&, 0&, vbNullString, "Block only those people whose screen names I list")
    Who& = FindWindowEx(PrivacyWin&, 0&, "_AOL_EDIT", vbNullString)
    ClickIcon (Block&)
    Call SendMessageByString(Who&, WM_SETTEXT, 0&, SN$)
    SetWho& = FindWindowEx(PrivacyWin&, 0&, "_AOL_ICON", vbNullString)
    For Index& = 1 To 21
        Edit& = FindWindowEx(PrivacyWin&, SetWho&, "_AOL_Icon", vbNullString)
    Next Index&
    ClickIcon Edit&
    ProgramPause 2
    Save& = FindWindowEx(PrivacyWin&, Edit&, "_AOL_Icon", vbNullString)
    Save& = FindWindowEx(PrivacyWin&, Save&, "_AOL_Icon", vbNullString)
    Save& = FindWindowEx(PrivacyWin&, Save&, "_AOL_Icon", vbNullString)
    ClickIcon Save&
End Sub

Public Sub AOLEnterPrivateRoom(theRoom As String)
    Call AOLKeyWord("aol://2719:2-2-" & theRoom$)
End Sub
Public Sub AOLEnterMemberRoom(Room As String)
    Call AOLKeyWord("aol://2719:61-2-" & Room$)
End Sub
Public Sub WAVPlay(Wav As String)
    Dim Check As String
    Check$ = Dir(Wav$)
    If Check$ = "" Then
        Exit Sub
    Else
        Call sndPlaySound(Wav$, SND_FLAG)
    End If
End Sub
Public Function FindWindows(Win1, Win2, Win3, Win4, Win5, Win6, Win7, Win8, Win9, Win10) As Long
'The variable Win1 etc. should be set to the
'class name of a window. Then if you don't
'have anymore windows to find just set the
'rest of the variables to ""
    Dim lWin As Long, lWin2 As Long, lWin3 As Long
    Dim lWin4 As Long, lWin5 As Long, lWin6 As Long
    Dim lWin7 As Long, lWin8 As Long, lWin9 As Long, lWin10 As Long
    lWin& = FindWindow(Win1, vbNullString)
    If Win2 = "" Then
        FindWindows = lWin&
        Exit Function
    Else
        lWin2& = FindWindowEx(lWin&, 0&, Win2, vbNullString)
        If Win3 = "" Then
            FindWindows = lWin2&
            Exit Function
        Else
            lWin3& = FindWindowEx(lWin2&, 0&, Win3, vbNullString)
            If Win4 = "" Then
                FindWindows = lWin3&
                Exit Function
            Else
                lWin4& = FindWindowEx(lWin3&, 0&, Win4, vbNullString)
                If Win5 = "" Then
                    FindWindows = lWin4&
                    Exit Function
                Else
                    lWin5& = FindWindowEx(lWin4&, 0&, Win5, vbNullString)
                    If Win6 = "" Then
                        FindWindows = lWin5&
                        Exit Function
                    Else
                        lWin6& = FindWindowEx(lWin5&, 0&, Win6, vbNullString)
                        If Win7 = "" Then
                            FindWindows = lWin6&
                            Exit Function
                        Else
                            lWin7& = FindWindowEx(lWin6&, 0&, Win7, vbNullString)
                            If Win8 = "" Then
                                FindWindows = lWin7&
                                Exit Function
                            Else
                                lWin8& = FindWindowEx(lWin7&, 0&, Win8, vbNullString)
                                If Win9 = "" Then
                                    FindWindows = lWin8&
                                    Exit Function
                                Else
                                    lWin9& = FindWindowEx(lWin8&, 0&, Win9, vbNullString)
                                    If Win10 = "" Then
                                        FindWindows = lWin9&
                                        Exit Function
                                    Else
                                        lWin10& = FindWindowEx(lWin9&, 0&, Win10, vbNullString)
                                        FindWindows = lWin10&
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
End Function
Public Function FindWindowsNext(NumberOver As Integer, Variable As Long, Parent As Long, Class As Long)
'Number over is the number of windows over your
'window is. The variable is the variable you have
'previously defined. Parent is the parent window
'and class is the class name of variable
    Do
        FindWindowsNext = FindWindowEx(Parent&, Variable&, Class&, vbNullString)
        NumberOver% = NumberOver% - 1
    Loop Until NumberOver = 0
End Function


Public Function hWndClass(hwnd As Long)
    Dim sSTR As String, ParentClass As String * 100
    sSTR$ = GetClassName(hwnd&, ParentClass$, 100)
    hWndClass = Left(sSTR$, ParentClass$)
End Function

Public Function hWndClose(hwnd As Long)
    Call SendMessage(hwnd&, WM_CLOSE, 0&, 0&)
End Function
Public Function hWndChangeCaption(hwnd As Long)
    Call SendMessageByString(hwnd&, WM_SETTEXT, 0&, 0&)
End Function
Public Function hWndIDNumber(hwnd As Long)
    Dim theID As String
    theID = GetWindowWord(hwnd, GWW_ID)
    hWndIDNumber = theID
End Function

Public Function hWndModule(hwnd As Long)
    Dim Instance As String, GetIt As String
    Dim sModuleFileName As String * 100
    Instance$ = GetWindowWord(hwnd, GWW_HINSTANCE)
    GetIt$ = GetModuleFileName(Instance$, sModuleFileName, 100)
    hWndModule = Left(sModuleFileName, GetIt$)
End Function

Public Function hWndParentClass(hwnd As Long)
    Dim sSTR As String, hWndParent As String
    Dim ParentClass As String * 100
    hWndParent$ = GetParent(hwnd&)
    sSTR$ = GetClassName(hWndParent, ParentClass, 100)
    hWndParentClass = Left(sSTR$, ParentClass)
End Function
Public Function hWndParentHandle(hwnd As Long)
    Dim hWndParent As Long
    hWndParentHandle = GetParent(hwnd)
End Function

Public Function hWndStyle(hwnd As Long)
    Dim WindowStyle As Long
    WindowStyle& = GetWindowLong(hwnd, GWL_STYLE)
    hWndStyle = WindowStyle&
End Function
Public Function hWndText(win As Long)
    Dim Buffer As String, TheLEN As Long
    TheLEN& = SendMessage(win&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(TheLEN&, 0&)
    Call SendMessageByString(win&, WM_GETTEXT, TheLEN& + 1, Buffer$)
    hWndText = Buffer$
End Function
Public Function hWndCaption(win As Long)
    Dim Buffer As String, TheLEN As Long
    TheLEN& = GetWindowTextLength(win&)
    Buffer$ = String(TheLEN&, 0&)
    Call GetWindowText(win&, Buffer$, TheLEN& + 1)
    hWndCaption = Buffer$
End Function
Public Sub ClickIcon(hwnd As Long)
    Call SendMessage(hwnd&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(hwnd&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub ClickButton(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub CDPlay()
    Call mciSendString("play cd", 0, 0, 0)
End Sub
Public Sub CDOpenDoor()
    Call mciSendString("set cd door open", 0, 0, 0)
End Sub
Public Sub CDPause()
    Call mciSendString("pause cd", 0, 0, 0)
End Sub
Public Function CDChangeTrack(Track As String)
    Call mciSendString("seek cd to " & Track$, 0, 0, 0)
End Function
Public Function CDStop()
    Call mciSendString("stop cd wait", 0, 0, 0)
End Function
Public Sub CDCloseDoor()
    Call mciSendString("set cd door closed", 0, 0, 0)
End Sub
Public Sub AOLAntiIdle()
    Dim Palette As Long, Modal As Long
    Dim Button As Long, Button2 As Long
    Palette& = FindWindow("_AOL_Palette", vbNullString)
    Button& = FindWindowEx(Palette&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (Button&)
    Modal& = FindWindow("_AOL_Modal", vbNullString)
    Button2& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (Button2&)
End Sub

Public Function ReturnStringBackwards(Text As String) As String
    'note: for vb6 users this code can be replaced with
    'ReturnStringBackwards$ = StrReverse(Text$)
    Dim TempString As String, Length As Long, Spaces As Long
    Dim NextChr As String, NewString As String
    TempString$ = Text$
    Length& = Len(TempString$)
    Do While Spaces& <= Length&
        Spaces& = Spaces& + 1
        NextChr$ = Mid$(TempString$, Spaces&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReturnStringBackwards$ = NewString$
End Function
Public Sub ChatCmdClearChat()
    If LCase$(AOLChatLine$) Like LCase$(".clear*") Then
        Call AOLChatClear
    End If
End Sub
Public Sub AOLChatClear()
    Dim ChatRoom&, ChatRoom2&
    ChatRoom& = AOLFindChatRoom&
    ChatRoom2& = FindWindowEx(ChatRoom&, 0&, "RICHCNTL", vbNullString)
    ChatRoom2& = FindWindowEx(ChatRoom&, ChatRoom2&, "RICHCNTL", vbNullString)
    Call SendMessage(ChatRoom2&, WM_CLEAR, 0&, 0&)
    Call SendMessageByString(ChatRoom2&, WM_SETTEXT, 0, "")
End Sub
Public Function ReturnStringWDots(Text As String) As String
    Dim TempString As String, Length As Long, Spaces As Long
    Dim NextChr As String, NewString As String
    TempString$ = Text$
    Length& = Len(TempString$)
    Do While Spaces& <= Length&
        Spaces& = Spaces& + 1
        NextChr$ = Mid$(TempString$, Spaces&, 1)
        NextChr$ = NextChr$ & "•"
        NewString$ = NewString$ & NextChr$
    Loop
    ReturnStringWDots$ = NewString$
End Function
Public Sub AOLFileSearch(file As String)
    Dim AOL&, MDI&, FileSrch&, Icon&, SoftSearch&, Edit&
    Call AOLKeyWord("File Search")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDI Client", vbNullString)
    Do
        FileSrch& = FindWindowEx(MDI&, 0&, vbNullString, "Filesearch")
    Loop Until FileSrch& <> 0
    Icon& = FindWindowEx(FileSrch&, 0&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(FileSrch&, Icon&, "_AOL_Icon", vbNullString)
    ClickIcon (Icon&)
    SoftSearch& = FindWindowEx(MDI&, 0&, vbNullString, "Software Search")
    Edit& = FindWindowEx(SoftSearch&, 0&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, file$)
    Call SendMessageByNum(Edit&, WM_CHAR, 0&, 13)
End Sub

Public Sub AOLWaitForMailToLoad()
    Dim AOL As Long, MDI As Long, MailBox As Long
    Dim theTree As Long, Check As Long, Check2 As Long
    Dim Check3 As Long, TabControl As Long, TabPage As Long
    Call AOLMailOpen
    Do
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDI Client", vbNullString)
        MailBox& = FindWindowEx(MDI&, 0&, vbNullString, AOLUserSN & "'s Online Mailbox")
    Loop Until MailBox& <> 0
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    theTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Do
        DoEvents
        Check& = SendMessage(theTree&, LB_GETCOUNT, 0, 0&)
        ProgramPause (2)
        Check2& = SendMessage(theTree&, LB_GETCOUNT, 0, 0&)
        ProgramPause (2)
        Check3& = SendMessage(theTree&, LB_GETCOUNT, 0, 0&)
    Loop Until Check& = Check2& And Check2& = Check3&
End Sub
Public Sub Image2HTMLPreview(TheForm As Form, theCode As String)
    Dim X As Long, TempSite As String
    TempSite$ = "C:\TempSite.html"
    Open TempSite$ For Output As #1
        Print #1, "<html><body bgcolor=""#000000"">" & theCode$ & "</body></html>"
    Close #1
    Call ShellExecute(TheForm.hwnd, "open", TempSite$, vbNullString, vbNullString, 3)
End Sub

Public Sub Image2HTML(TheForm As Form, theLoadedPicture As PictureBox, endCode As TextBox, htmlChar As TextBox, theSpeed As TextBox, thePercentage As PictureBox, Opt6 As OptionButton, Opt5 As OptionButton, Opt4 As OptionButton, Opt3 As OptionButton, Opt2 As OptionButton, Opt1 As OptionButton)
    Dim Index As Long, index2 As Long, PointColor As Long
    Dim xScale As Long, yScale As Long, TheChar As String
    Dim Quality As Long, xScale2 As Double, xScale3 As Double, Hold As String
    theLoadedPicture.Top = 0
    endCode.Text = ""
    theLoadedPicture.AutoRedraw = True
    theLoadedPicture.Appearance = 0
    theLoadedPicture.ScaleMode = 3
    If Len(htmlChar.Text) > 2 Then htmlChar.Text = "%"
    If theSpeed.Text = "Speed" Then theSpeed.Text = "0"
    TheChar$ = htmlChar.Text
    Hold$ = "<font face=""Arial"" size=2>"
    On Error GoTo ErrorHandler
    Opt1.Caption = "Worst"
    Opt2.Caption = "Very Low"
    Opt3.Caption = "Low"
    Opt4.Caption = "Medium"
    Opt5.Caption = "High"
    Opt6.Caption = "Very High"
    If Opt1.Value = True Then Quality& = 1
    If Opt2.Value = True Then Quality& = 2
    If Opt3.Value = True Then Quality& = 3
    If Opt4.Value = True Then Quality& = 4
    If Opt5.Value = True Then Quality& = 8
    If Opt6.Value = True Then Quality& = 12
    xScale& = theLoadedPicture.ScaleWidth - 1
    yScale& = theLoadedPicture.ScaleHeight - 1
    xScale2 = thePercentage.ScaleWidth / ((xScale& * yScale&) / (Quality& * Quality&))
    For Index& = 1 To yScale Step Quality&
      For index2& = 1 To xScale Step Quality&
        DoEvents
        ProgramPause theSpeed.Text
        PointColor& = theLoadedPicture.Point(Index&, index2&)
        Hold$ = Hold$ & "<font color=" & GetHTMLColor(PointColor&) & ">" & TheChar$ & "</font>"
        thePercentage.Line (xScale3, 0)-(xScale3 + xScale2, thePercentage.ScaleHeight), RGB(40, 40, 160), BF
        xScale3 = xScale3 + xScale2
      Next index2&
      Hold$ = Hold$ & "<br>"
      DoEvents
      If Len(Hold$) > 5000 Then endCode.Text = endCode.Text & Hold$: Hold$ = ""
      If Opt6.Value = True Then theLoadedPicture.Top = theLoadedPicture.Top - 1
      If Opt5.Value = True Then theLoadedPicture.Top = theLoadedPicture.Top - 1
      If Opt4.Value = True Then theLoadedPicture.Top = theLoadedPicture.Top - 1
      If Opt3.Value = True Then theLoadedPicture.Top = theLoadedPicture.Top - 3
      If Opt2.Value = True Then theLoadedPicture.Top = theLoadedPicture.Top - 8
      If Opt1.Value = True Then theLoadedPicture.Top = theLoadedPicture.Top - 8
    Next Index&
    theLoadedPicture.Top = 0
    thePercentage.Cls
    endCode.Text = endCode.Text & Hold$ & "</font>"
        Call SetWindowPos(TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    MsgBox "The decompilation of the image is done and the code has been generated! ", 32, "FrENzY32 v3"
        Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Exit Sub
ErrorHandler:
        Call SetWindowPos(TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    MsgBox "This image exceeds the maximum size for this quality level & speed, adjust the speed down a little bit and click on a lower quality code to avoid an overflow", 16, "FrENzY32 v3"
        Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    theLoadedPicture.Top = 0
    thePercentage.Cls
    Exit Sub
End Sub
Public Function Coverter(Text As String) As String
    Do While Len(Text$) < 2
      Text$ = "0" & Text$
    Loop
    Coverter = Text$
End Function
Public Function GetHTMLColor(theColor As Long) As String
    Dim Red1&, Green1&, Blue1&
    Red1& = theColor& And 255
    Green1& = theColor& \ 256 And 255
    Blue1& = theColor& \ 65536 And 255
    If Red1& = 255 Then Red1& = 254
    GetHTMLColor$ = Chr$(34) & "#" & Coverter(Hex(Red1&)) & Coverter(Hex(Green1&)) & Coverter(Hex(Blue1&)) & Chr$(34)
End Function

Public Function AOLUserSN() As String
    Dim AOL As Long, MDI As Long, WelcomeWin As Long, User As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    WelcomeWin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    User$ = hWndCaption(WelcomeWin&)
    Do
        If InStr(User$, "Welcome, ") Then
            Exit Do
        Else
            WelcomeWin& = FindWindowEx(MDI&, WelcomeWin&, "AOL Child", vbNullString)
        End If
    Loop
    AOLUserSN$ = Mid$(User$, 10, (InStr(User$, "!") - 10))
End Function

Public Function ReturnStringSpaced(Text As String)
    Dim TempString As String, Length As Long, Spaces As Long
    Dim NextChr As String, NewString As String
    TempString$ = Text$
    Length& = Len(TempString$)
    Do While Spaces& <= Length&
        Spaces& = Spaces& + 1
        NextChr$ = Mid$(TempString$, Spaces&, 1)
        NextChr$ = NextChr$ & " "
        NewString$ = NewString$ & NextChr$
    Loop
    ReturnStringSpaced = NewString$
End Function
Public Function ReturnStringLink(URL As String, Text$) As String
    ReturnStringLink$ = "<a href=" & Chr(34) & URL$ & Chr(34) & ">" & Text$ & "</a>"
End Function

Public Function ReturnStringHTML(Text As String) As String
    Dim TempString As String, Length As Long, Spaces As Long
    Dim NextChr As String, NewString As String, NumSpc As Long
    TempString$ = Text$
    Length& = Len(TempString$)
    Do While NumSpc& <= Length&
        Spaces& = Spaces& + 1
        NextChr$ = Mid$(TempString$, Spaces&, 1)
        NextChr$ = NextChr$ & "<html>"
        NewString$ = NewString$ & NextChr$
    Loop
    ReturnStringHTML$ = NewString$
End Function

Public Sub AOLBotIdle()
    Dim Palette As Long, Modal As Long
    Dim Button As Long, Button2 As Long
    Palette& = FindWindow("_AOL_Palette", vbNullString)
    Button& = FindWindowEx(Palette&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (Button&)
    Modal& = FindWindow("_AOL_Modal", vbNullString)
    Button2& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (Button2&)
End Sub
Public Sub AOLServerGetStatus(Who As String)
    Call AOLChatSend("/" & Who$ & " Send " & "Status")
End Sub
Public Sub AOLServerFindItem(Who As String, What As String)
    Call AOLChatSend("/" & Who$ & " Find " & What$)
End Sub
Public Sub AOLMailSend(Recipiants As String, subject As String, Message As String)
    Dim AOL As Long, toolbar As Long, ToolbarWin As Long
    Dim Button As Long, MDI As Long, MailWin As Long
    Dim Edit As Long, Rich As Long, Index As Long
    Dim Error As Long, Modal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    ToolbarWin& = FindWindowEx(toolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Button& = FindWindowEx(ToolbarWin&, 0&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(ToolbarWin&, Button&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(ToolbarWin&, Button&, "_AOL_Icon", vbNullString)
    ClickIcon (Button&)
    Do
        DoEvents
        MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
        MailWin& = FindWindowEx(MDI&, 0&, vbNullString, "Write Mail")
        Edit& = FindWindowEx(MailWin&, 0&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(MailWin&, 0&, "RICHCNTL", vbNullString)
        Button& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until MailWin& <> 0 And Edit& <> 0 And Rich& <> 0 And Button& <> 0
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, Recipiants$)
    Edit& = FindWindowEx(MailWin&, Edit&, "_AOL_Edit", vbNullString)
    Edit& = FindWindowEx(MailWin&, Edit&, "_AOL_Edit", vbNullString)
    Edit& = FindWindowEx(MailWin&, Edit&, "_AOL_Edit", vbNullString)
    Edit& = FindWindowEx(MailWin&, Edit&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, subject$)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    For Index& = 1 To 18
        Button& = FindWindowEx(MailWin&, Button&, "_AOL_Icon", vbNullString)
    Next Index&
    ClickIcon (Button&)
    Do
        DoEvents
        Error& = FindWindowEx(MDI&, 0&, vbNullString, "Error")
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        If MailWin& = 0 Then Exit Do
        If Modal& <> 0 Then
            Button& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
            ClickIcon (Button&)
            Call SendMessage(MailWin&, WM_CLOSE, 0, 0)
            Exit Sub
        End If
        If Error& <> 0 Then
            Call SendMessage(Error&, WM_CLOSE, 0, 0)
            Call SendMessage(MailWin&, WM_CLOSE, 0, 0)
            Exit Do
        End If
    Loop Until MailWin& = 0
End Sub
Public Sub AOLChatScrollList(List As ListBox)
    Dim Index As Long
    For Index& = 0 To List.ListCount - 1
        Call AOLChatSend(Index& & List.List(Index&))
        ProgramPause 0.3
    Next Index&
End Sub
Public Sub AOLWaitForOK()
    Dim OK As Long, OKButton As Long
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
    Loop Until OK& <> 0
    OKButton& = FindWindowEx(OK&, 0&, vbNullString, "OK")
    Call SendMessageByNum(OKButton&, WM_LBUTTONDOWN, 0, 0&)
    Call SendMessageByNum(OKButton&, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub ProgramToTray(TheForm As Form)
    Dim vbTray As NOTIFYICONDATA
    With vbTray
        .cbSize = Len(vbTray)
        .uId = vbNull
        .hwnd = TheForm.hwnd
        .ucallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .hIcon = TheForm.Icon
        .szTip = TheForm.Caption
    End With
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    TheForm.Hide
End Sub
Public Sub ProgramFromTray(TheForm As Form)
    Dim vbTray As NOTIFYICONDATA
    With vbTray
        .cbSize = Len(vbTray)
        .hwnd = TheForm.hwnd
        .uId = vbNull
    End With
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
End Sub
Public Sub AOLServerSendItem(Who As String, What As String)
    Call AOLChatSend("/" & Who$ & " Send " & What$)
End Sub

Public Sub mIRCActivate()
    Dim mIRC32 As Long, sSTR As String
    mIRC32& = FindWindow("mIRC32", vbNullString)
    sSTR$ = hWndCaption(mIRC32&)
    AppActivate (sSTR$)
End Sub

Public Function mIRCChangeCaption(NewCaption As String)
    Dim mIRC32 As Long
    mIRC32& = FindWindow("mIRC32", vbNullString)
    Call SendMessageByString(mIRC32&, WM_SETTEXT, 0, NewCaption$)
End Function
Public Function mIRCChatClear()
    Dim mIRC32 As Long, MDI As Long, Channel As Long
    Dim Edit As Long
    mIRC32& = FindWindow("mIRC32", vbNullString)
    MDI& = FindWindowEx(mIRC32&, 0, "MDIClient", vbNullString)
    Channel& = FindWindowEx(MDI&, 0, "channel", vbNullString)
    Edit& = FindWindowEx(Channel&, 0, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, "/Clear")
    Call SendMessageByNum(Edit&, WM_CHAR, 13, 0&)
End Function
Public Function mIRCChatSend(What$)
    Dim mIRC32 As Long, MDI As Long, Channel As Long
    Dim Edit As Long
    mIRC32& = FindWindow("mIRC32", vbNullString)
    MDI& = FindWindowEx(mIRC32&, 0, "MDIClient", vbNullString)
    Channel& = FindWindowEx(MDI&, 0, "channel", vbNullString)
    Edit& = FindWindowEx(Channel&, 0, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, What$)
    Call SendMessageByNum(Edit&, WM_CHAR, 13, 0&)
End Function
Public Function mIRCChatLine() As String
    Dim mIRC32 As Long, MDI As Long, Channel As Long
    mIRC32& = FindWindow("mIRC32", vbNullString)
    MDI& = FindWindowEx(mIRC32&, 0, "MDIClient", vbNullString)
    Channel& = FindWindowEx(MDI&, 0, "channel", vbNullString)
    mIRCChatLine$ = ReturnStringsLastLine(hWndText(Channel&))
End Function

Public Sub mIRCEnterRoom(theRoom As String)
    Call mIRCChatSend("/j #" & theRoom$)
End Sub
Public Sub mIRCMsgSomeone(Who As String, Message As String)
    Call mIRCChatSend("/msg " & Who$ & Message$)
End Sub
Public Function mIRCGetRoomCount() As Long
    Dim mIRC&, MDI&, Channel&, ListBox&
    mIRC& = FindWindow("mIRC32", vbNullString)
    MDI& = FindWindowEx(mIRC&, 0, "MDIClient", vbNullString)
    Channel& = FindWindowEx(MDI&, 0, "channel", vbNullString)
    ListBox& = FindWindowEx(Channel&, 0, "ListBox", vbNullString)
    mIRCGetRoomCount& = SendMessage(ListBox&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function mIRCMsgClear()
    Dim mIRC32 As Long, MDI As Long, Query As Long
    Dim Edit As Long
    mIRC32& = FindWindow("mIRC32", vbNullString)
    MDI& = FindWindowEx(mIRC32&, 0, "MDIClient", vbNullString)
    Query& = FindWindowEx(MDI&, 0, "query", vbNullString)
    Edit& = FindWindowEx(Query&, 0, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, "/Clear")
    Call SendMessageByNum(Edit&, WM_CHAR, 13, 0&)
End Function
Public Function mIRCMsgSend(What$)
    Dim mIRC32 As Long, MDI As Long, Query As Long
    Dim Edit As Long
    mIRC32& = FindWindow("mIRC32", vbNullString)
    MDI& = FindWindowEx(mIRC32&, 0, "MDIClient", vbNullString)
    Query& = FindWindowEx(MDI&, 0, "query", vbNullString)
    Edit& = FindWindowEx(Query&, 0, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, What$)
    Call SendMessageByNum(Edit&, WM_CHAR, 13, 0&)
End Function
Public Sub AOLBotFileFind(WhatFile As String)
    Call AOLChatSend("Does anyone have " & WhatFile)
    Call AOLChatSend("If you do type " & Chr(34) & "i do" & Chr(34))
    If AOLChatLine$ Like LCase("i do") Then
        Call AOLInstantMessage(AOLChatLineSN$, "Can you please send it?")
    End If
End Sub
Public Sub AOLIMMassSender(Recipiants As ListBox, Msg As String)
    Dim Index As Long
    For Index& = 0 To Recipiants.ListCount - 1
        Call AOLInstantMessage(Recipiants.List(Index&), Msg$)
    Next Index&
End Sub
Public Sub AOLPhishPhrases(Msg1 As TextBox, Msg2 As TextBox, Msg3 As TextBox)
    'I came up with these when i was bored, for more
    'check out the Frenzy32Misc bas file.
    Dim Rounder As Long
    Rounder& = Int(Rnd * 3)
    If Rounder& = 1 Then
        Msg1.Text = "Hello, I am an America Online Representative. I am sorry to inform you that we have lost critical information on your account. I am asking you to respond to this IM with your current password failure to do so will result in termination of your America Online service."
        Msg2.Text = "Thankyou. We hope that this will never happen again. We are placing 1 free month of AOL on your account."
        Msg3.Text = "Good-Bye"
        Exit Sub
    ElseIf Rounder& = 2 Then
        Msg1.Text = "America Online - Hello, We regret to inform you that we have had a recent computer failure and have misplaced your account information. Please Respond to this instant message with your password. Any errors or failure to do so will result in immediate account termination."
        Msg2.Text = "Thankyou for your cooperation. America Online will now credit your account with 1 Free Month of internet access."
        Msg3.Text = "GoodBye"
        Exit Sub
    ElseIf Rounder& = 3 Then
        Msg1.Text = "Hello, This is America Online Account Manegment. We have had a problem with one our employees here t AOL and have misplaced your account information. We ask that you respond to this instant message with your current password so that we can fix this error."
        Msg2.Text = "Thankyou, We will have your account information fixed in less than 5 min. Sorry for the inconvenience."
        Msg3.Text = "GoodBye"
        Exit Sub
    Else
        MsgBox "Error generating random phish phrase, please check your code."
        Exit Sub
    End If
End Sub

Public Sub AOLBCCMail(Recipiants As ListBox, subject As String, Message As String)
    Dim Index As Long, CurrPeep As String, thePeeps As String
    For Index& = 0 To Recipiants.ListCount - 1
        If InStr(Recipiants.List(Index&), ",") Then
            CurrPeep$ = Recipiants.List(Index&)
        Else
            CurrPeep$ = Recipiants.List(Index&) & ","
        End If
        thePeeps$ = thePeeps$ & CurrPeep$
    Next Index&
    Call AOLMailSend(thePeeps$, subject$, Message$)
End Sub

Public Sub ProgramCheckPW(CurrForm As Form, PW1 As String, PW2 As String, PW3 As String, PWInput As TextBox, PWForm1 As Form, PWForm2 As Form, PWForm3 As Form)
    'This is something i made for my program. It is a
    'password required form.
    If PWInput.Text Like LCase(PW1$) Then
        PWForm1.Show
        PWInput.Text = ""
        Unload CurrForm
    ElseIf PWInput.Text Like LCase(PW2$) Then
        PWForm2.Show
        PWInput.Text = ""
        Unload CurrForm
    ElseIf PWInput.Text Like LCase(PW3$) Then
        PWForm3.Show
        PWInput.Text = ""
        Unload CurrForm
    End If
End Sub
Public Sub AOLEchoBot(Who As String)
    If AOLChatLineSN$ Like LCase(Who$) Then
        Call AOLChatSend(AOLChatLine)
    End If
End Sub
Public Sub AOLRoomRunner(StartRoom As String)
    Dim Room As Long, OKWin As Long
    Dim OKButton As Long, theInt As Long
    Room& = AOLFindChatRoom
    theInt& = 1
    If Room& Then hWndClose (Room&)
    Do
        DoEvents
        Call AOLKeyWord("aol://2719:2-2-" & StartRoom$ & theInt&)
        Do
            DoEvents
            OKWin& = FindWindow("#32770", "America Online")
            DoEvents
        Loop Until OKWin& Or Room&
        If OKWin& <> 0 Then
            Do
                OKButton& = FindWindowEx(OKWin&, 0&, vbNullString, "OK")
                Call SendMessageByNum(OKButton&, WM_LBUTTONDOWN, 0, 0&)
                Call SendMessageByNum(OKButton&, WM_LBUTTONUP, 0, 0&)
            Loop Until OKWin& = 0
        End If
        theInt& = theInt& + 1
    Loop Until Room& <> 0
End Sub
Public Function AOLRoomBuster(theRoom As String, counter As Label)
    Dim Room As Long
    Room& = AOLFindChatRoom
    If Room& Then hWndClose (Room&)
    Do
        DoEvents
        ProgramPause 10
        Call AOLKeyWord("aol://2719:2-2-" & theRoom$)
        Call AOLOKOrChatWindow(theRoom$)
        counter = counter + 1
        If AOLFindChatRoom Then Exit Do
    Loop
End Function
Public Function AOLOKOrChatWindow(Room As String)
    'dos made this, i am too lazy to remake it
    'because it looks good and functions correctly
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    Room$ = LCase(ReplaceString(Room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = hWndCaption(AOLFindChatRoom)
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
End Function
Public Function ComputerProcessorType() As String
    Dim sSTR As String
    Dim MyVer As OSVERSIONINFO, MySys As SYSTEM_INFO
    Call GetSystemInfo(MySys)
    Select Case MySys.dwProcessorType
        Case PROCESSOR_INTEL_386
            sSTR$ = "Intel 386 DX"
        Case PROCESSOR_INTEL_486
            sSTR$ = "Intel 486 DX"
        Case PROCESSOR_INTEL_PENTIUM
            sSTR$ = "Intel Pentium Pro"
        Case PROCESSOR_MIPS_R4000
            sSTR$ = "Mips R-4000"
        Case PROCESSOR_ALPHA_21064
            sSTR$ = "Alpha 21064"
        Case Else
            sSTR$ = "Unknown Processor"
        End Select
        If MySys.dwNumberOrfProcessors > 1 Then
            sSTR$ = "Multiple " & sSTR$ & " Processors"
        Else
            sSTR$ = sSTR$ & " Processor"
        End If
        ComputerProcessorType$ = sSTR$
End Function
Public Sub PrintText(Text As String)
    Dim oldcursor&
    oldcursor& = Screen.MousePointer
    Screen.MousePointer = 11
    Printer.Print Text$
    Printer.NewPage
    Printer.EndDoc
    Screen.MousePointer = oldcursor&
End Sub
Public Function ComputerDiskInfo(theDrive As String) As String
    Dim DL&, S$, spaceloc%, FreeBytes&, TotalBytes&
    Dim SectorsPerCluster&, BytesPerSector&, NumberOfFreeClustors&, TotalNumberOfClustors&
    Dim BytesFree&, BytesTotal&, PercentFree&
    Dim Tmp1$, Tmp2$, Tmp3$, Tmp4$, Tmp5$, Tmp6$, Tmp7$
    If Right(theDrive$, 1) = "\" Then
        theDrive$ = Left(theDrive$, Len(theDrive$) - 1)
    End If
    If InStr(theDrive$, ":") Then
        theDrive$ = theDrive$
    Else
        theDrive$ = theDrive$ & ":"
    End If
    DL& = GetDiskFreeSpace(theDrive$, SectorsPerCluster, BytesPerSector, NumberOfFreeClustors, TotalNumberOfClustors)
    Tmp1$ = "Sectors Per Cluster :" & Format(SectorsPerCluster, "#,0")
    Tmp2$ = "Bytes Per Sector : " & Format(BytesPerSector, "#,0")
    Tmp3$ = "Number Of Free Clusters : " & Format(NumberOfFreeClustors, "#,0")
    Tmp4$ = "Total Number Of Clustors : " & Format(TotalNumberOfClustors, "#,0")
    Tmp5$ = TotalNumberOfClustors * SectorsPerCluster * BytesPerSector
    Tmp5$ = "Total Free Bytes : " & Format(Tmp5$, "#,0")
    Tmp6$ = NumberOfFreeClustors * SectorsPerCluster * BytesPerSector
    Tmp6$ = "Total Bytes : " & Format(Tmp6$, "#,0")
    ComputerDiskInfo$ = Tmp1$ & vbCrLf & Tmp2$ & vbCrLf & Tmp3$ & vbCrLf & Tmp4$ & vbCrLf & Tmp5$ & vbCrLf & Tmp6$
End Function
Public Sub ClickIconDouble(hwnd As Long)
    Call SendMessage(hwnd&, WM_LBUTTONDBLCLK, WM_LBUTTONDOWN, 0&)
End Sub
Public Sub ClickList(ListHandle As Long, Index As Long)
    Call SendMessage(ListHandle&, LB_SETCURSEL, CLng(Index&), 0&)
End Sub
Public Sub hWndOntop(frm As Form)
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Public Sub hWndSetText(hwnd As Long, thetext$)
    Call SendMessageByString(hwnd&, WM_SETTEXT, 0&, thetext$)
End Sub
Public Sub FormOnTop(frm As Form)
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Public Sub hWndOfftop(frm As Form)
    Call SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Public Sub FormOfftop(frm As Form)
    Call SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Public Sub hWndHide(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub
Public Function hWndExists(hwnd As Long) As Boolean
    If IsWindow(hwnd&) Then
        hWndExists = True
    Else
        hWndExists = False
    End If
End Function
Public Sub hWndFlash(hwnd As Long)
    Call FlashWindow(hwnd&, 100)
End Sub
Public Sub hWndShow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub
Public Sub hWndSpyWithTextBoxs(WinHdl As TextBox, WinClass As TextBox, WinTxT As TextBox, WinStyle As TextBox, WinIDNum As TextBox, WinPHandle As TextBox, WinPText As TextBox, WinPClass As TextBox, WinModule As TextBox)
    Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
    Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
    Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
    Dim hInstance As Long, sParentWindowText As String * 100
    Dim sModuleFileName As String * 100, r As Long
    Static hWndLast As Long
        Call GetCursorPos(pt32)
        ptx = pt32.X
        pty = pt32.Y
        hWndOver = WindowFromPointXY(ptx, pty)
        If hWndOver <> hWndLast Then
            hWndLast = hWndOver
            WinHdl.Text = "Window Handle: " & hWndOver
            r = GetWindowText(hWndOver, sWindowText, 100)
            WinTxT.Text = "Window Text: " & Left(sWindowText, r)
            r = GetClassName(hWndOver, sClassName, 100)
            WinClass.Text = "Window Class Name: " & Left(sClassName, r)
            lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
            WinStyle.Text = "Window Style: " & lWindowStyle
            hWndParent = GetParent(hWndOver)
                If hWndParent <> 0 Then
                    wID = GetWindowWord(hWndOver, GWW_ID)
                    WinIDNum.Text = "Window ID Number: " & wID
                    WinPHandle.Text = "Parent Window Handle: " & hWndParent
                    r = GetWindowText(hWndParent, sParentWindowText, 100)
                    WinPText.Text = "Parent Window Text: " & Left(sParentWindowText, r)
                    r = GetClassName(hWndParent, sParentClassName, 100)
                    WinPClass.Text = "Parent Window Class Name: " & Left(sParentClassName, r)
                Else
                    WinIDNum.Text = "Window ID Number: N/A"
                    WinPHandle.Text = "Parent Window Handle: N/A"
                    WinPText.Text = "Parent Window Text : N/A"
                    WinPClass.Text = "Parent Window Class Name: N/A"
                End If
                    hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
                    r = GetModuleFileName(hInstance, sModuleFileName, 100)
            WinModule.Text = "Module: " & Left(sModuleFileName, r)
        End If
End Sub

Public Sub AOLKeyWord(KeyWord As String)
    Dim Edit As Long
    Edit& = FindWindows("AOL Frame25", "AOL Toolbar", "_AOL_Toolbar", "_AOL_Combobox", "Edit", 0, 0, 0, 0, 0)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, KeyWord$)
    Call SendMessageLong(Edit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(Edit&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub hWndMaximize(hwnd As Long)
    Call ShowWindow(hwnd, SW_MAXIMIZE)
End Sub
Public Sub hWndMinimize(hwnd As Long)
    Call ShowWindow(hwnd&, SW_MINIMIZE)
End Sub
Public Sub hWndEnable(hwnd As Long)
    Call EnableWindow(hwnd&, 1)
End Sub
Public Sub hWndActivate(hwnd As Long)
    Call AppActivate(hwnd)
End Sub
Public Sub AOLIMIgnorer(Ignore As Boolean)
    Dim State As String
    If Ignore = True Then
        State$ = "Off"
    Else
        State$ = "On"
    End If
    Call AOLInstantMessage("$IM_" & State$, " ")
End Sub
Public Sub AOLRoomToList(thelist As ListBox, AddUser As Boolean)
    'By Dos, i was too lazy to make my own because
    'i dont really understant all of the api calls
    'and besides this one works fine.
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = AOLFindChatRoom&
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
            If ScreenName$ <> AOLUserSN$ Or AddUser = True Then
                thelist.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub AOLRoomToClipBoard(List As ListBox)
    Dim Index As Long, thelist As String
    For Index& = 0 To List.ListCount - 1
        If Index& = 0 Then
            thelist$ = List.List(Index&)
        Else
            thelist$ = thelist$ & "," & List.List(Index&)
        End If
    Next
    Clipboard.Clear
    Clipboard.SetText thelist$
End Sub

Public Function AOLIMSender() As String
    Dim AOL As Long, MDI As Long, IMWin As Long, Caption As String
    Dim theIM As Long, IMCaption As String, theSN As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    IMWin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = hWndCaption(IMWin&)
    If InStr(Caption$, "Instant Message") > 0 Then
        theIM& = IMWin&
        GoTo Found
    Else
        Do
            IMWin& = FindWindowEx(MDI&, IMWin&, "AOL Child", vbNullString)
            Caption$ = hWndCaption(IMWin&)
            If InStr(Caption$, "Instant Message") > 0 Then
                theIM& = IMWin&
                GoTo Found
            End If
        Loop Until IMWin& = 0&
    End If
Found:
    IMCaption$ = Caption$
    theSN$ = Mid(IMCaption$, InStr(IMCaption$, ":") + 2)
    AOLIMSender$ = theSN$
End Function
Public Function AOLIMLastMsg() As String
    Dim AOL As Long, MDI As Long, IMWin As Long, Caption As String
    Dim theIM As Long, IMCaption As String, theSN As String
    Dim FindMe As Long, FindMe2 As Long, MsgString As String, Rich As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    IMWin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = hWndCaption(IMWin&)
    If InStr(Caption$, "Instant Message") > 0 Then
        theIM& = IMWin&
        GoTo Found
    Else
        Do
            IMWin& = FindWindowEx(MDI&, IMWin&, "AOL Child", vbNullString)
            Caption$ = hWndCaption(IMWin&)
            If InStr(Caption$, "Instant Message") > 0 Then
                theIM& = IMWin&
                GoTo Found
            End If
        Loop Until IMWin& = 0&
    End If
Found:
    Rich& = FindWindowEx(theIM&, 0&, "RICHCNTL", vbNullString)
    MsgString$ = hWndText(Rich&)
    FindMe& = InStr(MsgString$, Chr(9))
    Do
        FindMe2& = FindMe&
        FindMe& = InStr(FindMe2& + 1, MsgString$, Chr(9))
    Loop Until FindMe& <= 0&
    MsgString$ = Right(MsgString$, Len(MsgString$) - FindMe2& - 1)
    AOLIMLastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function
Public Sub AOLIMResponder(Msg As String)
    Dim IM As Long, Rich As Long, Icon As Long
    Dim AOL As Long, MDI As Long, IMWin As Long, Caption As String
    Dim theIM As Long, IMCaption As String, theSN As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    IMWin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = hWndCaption(IMWin&)
    If InStr(Caption$, "Instant Message") > 0 Then
        theIM& = IMWin&
        GoTo Found
    Else
        Do
            IMWin& = FindWindowEx(MDI&, IMWin&, "AOL Child", vbNullString)
            Caption$ = hWndCaption(IMWin&)
            If InStr(Caption$, "Instant Message") > 0 Then
                theIM& = IMWin&
                GoTo Found
            End If
        Loop Until IMWin& = 0&
    End If
Found:
    Rich& = FindWindowEx(theIM&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(theIM&, Rich&, "RICHCNTL", vbNullString)
    Icon& = FindWindowEx(theIM&, 0&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(theIM&, Icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
    DoEvents
    Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub AOLInstantMessage(Who As String, Msg As String)
    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim Button As Long, OK As Long, Index&
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call AOLKeyWord("aol://9293:" & Who$)
    ProgramPause 4
    Do
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until IM& <> 0
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    Button& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    For Index& = 1 To 8
        Button& = FindWindowEx(IM&, Button&, "_AOL_Icon", vbNullString)
    Next Index&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
    Call SendMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Button&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0 Or IM& = 0
    If OK& <> 0 Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call hWndClose(IM&)
    End If
End Sub
Public Sub AOLChatSend(Text As String)
    Dim Room As Long, Rich As Long
    Room& = AOLFindChatRoom
    Rich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(Room&, Rich&, "RICHCNTL", vbNullString)
    Call SetFocusAPI(Rich&)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Text$)
    DoEvents
    Call SendMessageByNum(Rich&, WM_CHAR, 13, 0&)
    Call SendMessageByNum(Rich&, WM_CHAR, 13, 0&)
    ProgramPause 0.5
End Sub
Public Function AOLChatName() As String
    AOLChatName$ = hWndCaption(AOLFindChatRoom)
End Function
Public Sub FormDrag(frm As Form)
'Be Sure To Call This In The Form Drag Event
'Of The Form
    Call ReleaseCapture
    Call SendMessage(frm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub FormCenter(frm As Form)
    frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
    frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub
Public Sub FormHide(frm As Form)
    frm.Hide
End Sub
Public Sub FormShow(frm As Form)
    frm.Show
    Call FormCenter(frm)
    Call FormOnTop(frm)
End Sub
Public Sub FormExitRight(frm As Form)
    Do
        DoEvents
        frm.Left = frm.Left + 250
    Loop Until frm.Left > Screen.Width
    FormCenter frm
    FormUnload frm
End Sub
Public Sub FormLoad(frm As Form)
    frm.Load
    Call FormCenter(frm)
    Call FormOnTop(frm)
End Sub
Public Sub FormUnload(frm As Form)
    Unload frm
End Sub

Public Function MathADD(FirstNum As Integer, SecondNum As Integer)
    MathADD = FirstNum% + SecondNum%
End Function
Public Function MathSubtract(FirstNum As Integer, SecondNum As Integer)
    MathSubtract = FirstNum% - SecondNum%
End Function
Public Function MathDivide(FirstNum As Integer, SecondNum As Integer)
    MathDivide = FirstNum% / SecondNum%
End Function
Public Function MathMultiply(FirstNum As Integer, SecondNum As Integer)
    MathMultiply = FirstNum% * SecondNum%
End Function
Public Sub AsciiBuildChart(Ctrl As ListBox)
    Dim theIndex As Long
    Ctrl.Columns = 1
    For theIndex = 33 To 255
        Ctrl.AddItem Chr(theIndex)
    Next theIndex
End Sub
Public Sub AsciiScroll(Text As String)
    Dim counter As Long
    If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then Text$ = Text$ & Chr$(13) & Chr$(10)
    Do While InStr(Text$, Chr$(13)) <> 0
        counter& = counter& + 1
        Call AOLChatSend(Mid$(Text$, 1, InStr(Text$, Chr$(13)) - 1))
        If counter = 3 Then
            ProgramPause (2.9)
            counter& = 0
        End If
        Text$ = Mid$(Text$, InStr(Text$, Chr$(13) & Chr$(10)) + 2)
    Loop
End Sub
Public Sub AsciiKill(thestring As String)
    Dim KillChrs As Long
    For KillChrs& = 33 To 255
        If InStr(thestring, Chr(KillChrs&)) Then
            Call ReplaceString(thestring, Chr(KillChrs&), "")
        End If
    Next KillChrs&
End Sub
Public Function AOLFindChatRoom() As Long
    Dim AOL As Long, MDI As Long, Room As Long
    Dim List As Long, Rich As Long, counter As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Room& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    List& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Rich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
    counter& = 1
    If List& <> 0 And Rich& <> 0 Then
        AOLFindChatRoom& = Room&
    Else
        Do
            counter& = counter& + 1
            Room& = FindWindowEx(MDI&, Room&, "AOL Child", vbNullString)
            If List& <> 0 And Rich& <> 0 Then
                AOLFindChatRoom& = Room&
            End If
        Loop Until AOLFindChatRoom& = Room& Or counter& = 42
    End If
End Function
Public Function AOLIsOnline() As Boolean
    Dim AOL As Long, MDI As Long, BuddyList As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDI Client", vbNullString)
    BuddyList& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
    If BuddyList& <> 0 Then
        AOLIsOnline = True
    Else
        AOLIsOnline = False
    End If
End Function
Public Sub AOLMailOpen()
    Dim toolbar As Long, ToolbarWin As Long, Button As Long, AOL As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    ToolbarWin& = FindWindowEx(toolbar&, 0&, "_AOL_Toolbar", vbNullString)
    Button& = FindWindowEx(ToolbarWin&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (Button&)
End Sub
Public Sub WindowsClearDocuments()
    Call SHAddToRecentDocs(0, 0)
End Sub
Public Sub WindowsLoadFonts(List As Control)
    Dim Index As Long
    List.Clear
    For Index& = 1 To Screen.FontCount
        List.AddItem (Screen.Fonts(Index& - 1))
    Next Index&
End Sub
Public Sub WindowsMakeShortcut(ShortcutDir As String, ShortcutName As String, ShortcutPath As String)
    Dim WinShortcutDir As String, WinShortcutName As String, WinShortcutExePath As String, RetVal As Long
    WinShortcutDir$ = ShortcutDir$
    WinShortcutName$ = ShortcutName$
    WinShortcutExePath$ = ShortcutPath$
    RetVal& = fCreateShellLink("", WinShortcutName$, WinShortcutExePath$, "")
    Name "C:\Windows\Start Menu\Programs\" & WinShortcutName$ & ".LNK" As WinShortcutDir$ & "\" & WinShortcutName$ & ".LNK"
End Sub
Public Function WindowsVersion() As String
    Dim sSTR As String, DL As Long
    Dim MyVer As OSVERSIONINFO, MySys As SYSTEM_INFO
    #If Win32 Then
        MyVer.dwOSVersionInfoSize = 148
        DL& = GetVersionEx&(MyVer)
        If MyVer.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
            sSTR$ = "Windows95 "
    ElseIf MyVer.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            sSTR$ = "WindowsNT "
    Else
            sSTR$ = "Windows98 "
        End If
        #If Win16 Then
            sSTR$ = "Windows3.x"
            Exit Function
        #End If
    #End If
    sSTR$ = sSTR$ & MyVer.dwMajorVersion & "." & MyVer.dwMinorVersion & " Build " & MyVer.dwBuildNumber
    WindowsVersion$ = sSTR$
End Function
Public Function AOLGetVersion() As Long
    Dim AOLMenus As Long, SubMenu As Long, Item As Long, MenuStr As String
    Dim FindStr As Long, AOL As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    AOLMenus& = GetMenu(AOL&)
    SubMenu& = GetSubMenu(AOLMenus&, 0)
    Item& = GetMenuItemID(SubMenu&, 8)
    MenuStr$ = String$(100, " ")
    FindStr& = GetMenuString(SubMenu&, Item&, MenuStr$, 100, 1)
    If UCase(MenuStr$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
        AOLGetVersion& = 3
    Else
        AOLGetVersion& = 4
    End If
End Function

Public Sub AOLRoomToCombo(List As ListBox, Combo As ComboBox)
    Dim X As Long
    Call AOLRoomToList(List, False)
    For X = 0 To List.ListCount
        Combo.AddItem (List.List(X))
    Next X
End Sub
Public Sub AOLRoomToTextBox(List As ListBox, Text As TextBox)
    Dim SN As String, X As Long
    Call AOLRoomToList(List, False)
    For X = 0 To List.ListCount - 1
        SN$ = SN$ & List.List(X)
    Next X
    If Text.Text = "" Then
        Text.Text = SN$ & Chr(13) & Chr(10)
    Else
        Text.Text = Text.Text & SN$ & Chr(13) & Chr(10)
    End If
End Sub
Public Function AOLChatLine() As String
    Dim Room As Long, Rich As Long, TheChatLine As String
    Room& = AOLFindChatRoom&
    Rich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
    TheChatLine$ = hWndText(Rich&)
    TheChatLine$ = ReturnStringsLastLine(TheChatLine$)
    If InStr(TheChatLine$, Chr(9)) = 0 Then
        AOLChatLine$ = ""
        Exit Function
    End If
    AOLChatLine$ = Right(TheChatLine$, Len(TheChatLine$) - InStr(TheChatLine$, Chr(9)))
End Function
Public Sub ProgramPause(Length As Long)
    Dim Current As Long
    Current& = Timer
    Do Until Timer - Current& >= Length&
        DoEvents
    Loop
End Sub
Public Sub ProgramRunOnStartup()
    Call INIWrite("Windows", "Load", App.Path & "\" & App.EXEName, "C:\Window\Win.ini")
End Sub
Public Sub AOLClearLocations()
    Dim AOL As Long, toolbar As Long, ToolbarWin As Long
    Dim TheCombo As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    ToolbarWin& = FindWindowEx(toolbar&, 0&, "_AOL_Toolbar", vbNullString)
    TheCombo& = FindWindowEx(ToolbarWin&, 0&, "_AOL_Combobox", vbNullString)
    Call SendMessage(TheCombo&, CB_RESETCONTENT, 0, 0)
End Sub
Public Sub ComboRemoveItem(ComboWin As Long, thestring As String)
    Dim FindIt As Long, DeleteIt As Long
    FindIt& = SendMessageByString(ComboWin&, CB_FINDSTRINGEXACT, -1, thestring$)
    If FindIt& <> -1 Then
        Call SendMessageByString(ComboWin&, CB_DELETESTRING, FindIt&, 0)
    End If
End Sub
Public Sub LogWriteTo(What As String, LoGPath As String)
    Dim file As String, sSTR As String
    If InStr(LoGPath, ".") <= 0 Then Exit Sub
    file$ = FreeFile
    Open LoGPath For Binary Access Write As #file$
        sSTR$ = What$ & Chr(10)
        Put #1, LOF(1) + 1, sSTR$
    Close file$
End Sub
Public Sub AOLMailReadCurrent()
    Dim AOL As Long, MDI As Long, MailWin As Long
    Dim TabControl As Long, TabPage As Long, theTree As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    MailWin& = FindWindowEx(MDI&, 0&, "New Mail", vbNullString)
    TabControl& = FindWindowEx(MailWin&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    theTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessage(theTree&, WM_KEYDOWN, VK_RETURN, 0)
    Call SendMessage(theTree&, WM_KEYUP, VK_RETURN, 0)
End Sub
Public Sub AOLMailForward()
    Dim AOL As Long, MDI As Long, MailWin As Long, Button As Long, Index As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    MailWin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Button& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
    For Index& = 1 To 8
        Button& = FindWindowEx(MailWin&, Button&, "_AOL_Icon", vbNullString)
    Next Index&
    DoEvents
    ClickIcon (Button&)
End Sub
Public Sub AOLMassMailer(Recipiants As ListBox)
    'This is just a simple example, it will work
    'however, it is not error proof im sure.
    Dim Index As Long, People As String, AOL As Long
    Dim MDI As Long, MailWin As Long, Icon As Long, theCount&
    Dim TabControl As Long, TabPage As Long, theTree As Long
    For Index& = 0 To Recipiants.ListCount - 1
        People$ = People$ & Recipiants.List(Index&)
    Next Index&
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    MailWin& = FindWindowEx(MDI&, 0&, vbNullString, AOLUserSN & "'s Online Mailbox")
    TabControl& = FindWindowEx(MailWin&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    theTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Icon& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
    Do
        Call AOLMailOpen
        Call AOLWaitForMailToLoad
        theCount& = AOLMailCountNew&
        Call SendMessage(theTree&, LB_SETCURSEL, 0, 0&)
        Call AOLMailReadCurrent
        ProgramPause 4
        Call AOLMailForward
        Call AOLMailForwardAndSend(People$)
    Loop Until theCount& = 0
End Sub
Public Sub AOLMailKeepAsNew()
    Dim AOL As Long, MDI As Long, MailBox As Long
    Dim Button As Long, Index As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    MailBox& = FindWindowEx(MDI&, 0&, vbNullString, AOLUserSN & "'s Online Mailbox")
    Button& = FindWindowEx(MDI&, MailBox&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(MDI&, Button&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(MDI&, Button&, "_AOL_Icon", vbNullString)
    Call SendMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Button&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function AOLMailFindBox() As Long
    Dim AOL As Long, MDI As Long, MailBox As Long, Thebox As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Thebox& = FindWindowEx(MDI&, 0&, vbNullString, AOLUserSN & "'s Online Mailbox")
    If AOLMailFindBox& <= 0 Then
        Call AOLMailOpen
    End If
    AOLMailFindBox& = Thebox&
End Function
Public Sub AOLMailReadNext()
    Dim AOL As Long, MDI As Long, MailWin As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDI Client", vbNullString)
    MailWin& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Button& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(MailWin&, Button&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(MailWin&, Button&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(MailWin&, Button&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(MailWin&, Button&, "_AOL_Icon", vbNullString)
    Button& = FindWindowEx(MailWin&, Button&, "_AOL_Icon", vbNullString)
    If Button& <= 0 Then MsgBox "There is no more mail in your box"
    ClickIcon (Button&)
End Sub
Public Sub AOLMailClickReadButton()
    Dim AOL As Long, MDI As Long, Box As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Box& = FindWindowEx(MDI&, 0&, vbNullString, AOLUserSN & "'s Online Mailbox")
    Button& = FindWindowEx(Box&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (Button&)
End Sub
Public Sub AOLMailForwardAndSend(Recipiants As String)
    Dim AOL As Long, MDI As Long, Mail As Long, Edit As Long
    Dim RichText As Long, Button As Long, Index As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    Do
        DoEvents
        MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
        Mail& = FindWindowEx(MDI&, 0&, vbNullString, "Fwd: ")
        Edit& = FindWindowEx(Mail&, 0&, "_AOL_Edit", vbNullString)
        RichText& = FindWindowEx(Edit&, 0&, "RICHCNTL", vbNullString)
        Button& = FindWindowEx(Mail&, 0&, "_AOL_Icon", vbNullString)
    Loop Until Mail& <> 0 And Edit& <> 0 And RichText& <> 0 And Button& <> 0
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, Recipiants$)
    For Index& = 1 To 14
        Button& = FindWindowEx(Mail&, Button&, "_AOL_Icon", vbNullString)
    Next Index&
    ClickIcon (Button&)
    Do
        DoEvents
        Mail& = FindWindowEx(MDI&, 0&, vbNullString, "Fwd: ")
        Edit& = FindWindowEx(Mail&, 0&, "_AOL_Edit", vbNullString)
    Loop Until Edit& = 0
End Sub
Public Function AOLMailCountFlash() As Long
    Dim AOL As Long, MDI As Long, FlashWin As Long
    Dim Tree As Long, Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    FlashWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    Tree& = FindWindowEx(FlashWin&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(Tree&, LB_GETCOUNT, 0&, 0&)
    AOLMailCountFlash& = Count&
End Function
Public Function AOLMailCountNew() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim theTree As Long, Count As Long
    MailBox& = AOLMailFindBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    theTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(theTree&, LB_GETCOUNT, 0&, 0&)
    AOLMailCountNew& = Count&
End Function
Public Function AOLMailCountOld() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim theTree As Long, Count As Long
    MailBox& = AOLMailFindBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    theTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(theTree&, LB_GETCOUNT, 0&, 0&)
    AOLMailCountOld& = Count&
End Function


Public Sub ProgramDecompileProtect(ExePath As String)
    Dim TheFile As String
    On Error Resume Next
    If Not InStr(ExePath, "\") Then MsgBox "Executable File Not Found", vbOKOnly, "FrENzY32.bas": Exit Sub
    TheFile$ = FreeFile
    Open ExePath$ For Binary As #TheFile$
        Seek #TheFile, 25
        Put #TheFile, , "."
    Close #1
    If Err Then MsgBox "Not A Visual Basic Made File!", vbOKOnly, "Error In File": Exit Sub
    MsgBox "You're File Has Been Protected", vbOKOnly, "FrENzY32.bas"
End Sub
Public Sub BasHomepage()
    Call AOLKeyWord("http://www.come.to/izekial83/")
End Sub
Public Sub BasMailMaker(Msg As String)
    Call AOLMailSend("funkdemon@yahoo.com", "«ð»«ð» FrENzY32 «ð»«ð»", Msg$)
End Sub
Public Sub BasIMMaker(Msg As String)
    Call AOLInstantMessage("izekial83", Msg$)
End Sub
Public Sub MIDIPlay(Midi As String)
    Dim file As String
    file$ = Dir(Midi$)
    If file$ <> "" Then
        Call mciSendString("play " & Midi$, 0&, 0, 0)
    End If
End Sub
Public Sub MIDIStop(Midi As String)
    Dim file As String
    file$ = Dir(Midi$)
    If file$ <> "" Then
        Call mciSendString("stop " & Midi$, 0&, 0, 0)
    End If
End Sub
Public Sub ListRemoveItem(ListWin As Long, thestring As String)
    Dim FindIt As Long, DeleteIt As Long
    FindIt& = SendMessageByString(ListWin&, LB_FINDSTRINGEXACT, -1, thestring$)
    If FindIt& <> -1 Then
        Call SendMessageByString(ListWin&, LB_DELETESTRING, FindIt&, 0)
    End If
End Sub
Public Function ProgramPath() As String
    ProgramPath$ = App.Path & "\" & App.EXEName
End Function
Public Function FormSetAsAOLBarChild(Form As Form, XPosition As Long, YPosition As Long)
    Dim AOL As Long, toolbar As Long
    Form.Top = YPosition&
    Form.Left = XPosition&
    AOL& = FindWindow("AOL Frame25", vbNullString)
    toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Call SetParent(Form.hwnd, toolbar&)
    Call ShowWindow(AOL&, 2)
    Call ShowWindow(AOL&, 3)
End Function
Public Sub ControlMake3D(TheForm As Form, TheControl As Control)
    Dim OldMode As Long
    If TheForm.AutoRedraw = False Then
        OldMode = TheForm.ScaleMode
            TheForm.ScaleMode = 3
            TheForm.AutoRedraw = True
            TheForm.CurrentX = TheControl.Left - 1
            TheForm.CurrentY = TheControl.Top + TheControl.Height
            TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
            TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
            TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
            TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
            TheForm.AutoRedraw = False
        TheForm.ScaleMode = OldMode
    End If
    If TheForm.AutoRedraw = True Then
        OldMode = TheForm.ScaleMode
            TheForm.ScaleMode = 3
            TheForm.CurrentX = TheControl.Left - 1
            TheForm.CurrentY = TheControl.Top + TheControl.Height
            TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
            TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
            TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
            TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
        TheForm.ScaleMode = OldMode
    End If
End Sub
Public Function ReturnStringCoded(sSTR As String) As String
    Dim TempString As String, StringLength As Long, NumSpaces As Long
    Dim NextChr As String, NewString As String
    TempString$ = sSTR$
    StringLength& = Len(TempString$)
    Do While NumSpaces& <= StringLength&
        NumSpaces& = NumSpaces& + 1
        NextChr$ = Mid$(TempString$, NumSpaces&, 1)
        If NextChr$ = "A" Then NextChr$ = "š"
        If NextChr$ = "B" Then NextChr$ = "œ"
        If NextChr$ = "C" Then NextChr$ = "¢"
        If NextChr$ = "D" Then NextChr$ = "¤"
        If NextChr$ = "E" Then NextChr$ = "±"
        If NextChr$ = "F" Then NextChr$ = "°"
        If NextChr$ = "G" Then NextChr$ = "²"
        If NextChr$ = "H" Then NextChr$ = "³"
        If NextChr$ = "I" Then NextChr$ = "µ"
        If NextChr$ = "J" Then NextChr$ = "ª"
        If NextChr$ = "K" Then NextChr$ = "¹"
        If NextChr$ = "L" Then NextChr$ = "º"
        If NextChr$ = "M" Then NextChr$ = "Ÿ"
        If NextChr$ = "N" Then NextChr$ = "í"
        If NextChr$ = "O" Then NextChr$ = "î"
        If NextChr$ = "P" Then NextChr$ = "ï"
        If NextChr$ = "Q" Then NextChr$ = "ð"
        If NextChr$ = "R" Then NextChr$ = "ñ"
        If NextChr$ = "S" Then NextChr$ = "ò"
        If NextChr$ = "T" Then NextChr$ = "ó"
        If NextChr$ = "U" Then NextChr$ = "ô"
        If NextChr$ = "V" Then NextChr$ = "õ"
        If NextChr$ = "W" Then NextChr$ = "ö"
        If NextChr$ = "X" Then NextChr$ = "ø"
        If NextChr$ = "Y" Then NextChr$ = "ù"
        If NextChr$ = "Z" Then NextChr$ = "ú"
        If NextChr$ = " " Then NextChr$ = " "
        If NextChr$ = "a" Then NextChr$ = "'"
        If NextChr$ = "b" Then NextChr$ = "û"
        If NextChr$ = "c" Then NextChr$ = "ü"
        If NextChr$ = "d" Then NextChr$ = "ý"
        If NextChr$ = "e" Then NextChr$ = "þ"
        If NextChr$ = "f" Then NextChr$ = "Æ"
        If NextChr$ = "g" Then NextChr$ = "Ç"
        If NextChr$ = "h" Then NextChr$ = "Ì"
        If NextChr$ = "i" Then NextChr$ = "Í"
        If NextChr$ = "j" Then NextChr$ = "Î"
        If NextChr$ = "k" Then NextChr$ = "Ï"
        If NextChr$ = "l" Then NextChr$ = "Ø"
        If NextChr$ = "m" Then NextChr$ = "Þ"
        If NextChr$ = "n" Then NextChr$ = "ß"
        If NextChr$ = "o" Then NextChr$ = "†"
        If NextChr$ = "p" Then NextChr$ = "ƒ"
        If NextChr$ = "q" Then NextChr$ = "Œ"
        If NextChr$ = "r" Then NextChr$ = "Š"
        If NextChr$ = "s" Then NextChr$ = "‡"
        If NextChr$ = "t" Then NextChr$ = "¡"
        If NextChr$ = "u" Then NextChr$ = "£"
        If NextChr$ = "v" Then NextChr$ = "§"
        If NextChr$ = "w" Then NextChr$ = "ì"
        If NextChr$ = "x" Then NextChr$ = "ë"
        If NextChr$ = "y" Then NextChr$ = "ê"
        If NextChr$ = "z" Then NextChr$ = "é"
        If NextChr$ = "1" Then NextChr$ = "è"
        If NextChr$ = "2" Then NextChr$ = "ç"
        If NextChr$ = "3" Then NextChr$ = "æ"
        If NextChr$ = "4" Then NextChr$ = "á"
        If NextChr$ = "5" Then NextChr$ = "å"
        If NextChr$ = "6" Then NextChr$ = "â"
        If NextChr$ = "7" Then NextChr$ = "ã"
        If NextChr$ = "8" Then NextChr$ = "ä"
        If NextChr$ = "9" Then NextChr$ = "à"
        If NextChr$ = "0" Then NextChr$ = "×"
        NewString$ = NewString$ & NextChr$
    Loop
    ReturnStringCoded$ = NewString$
End Function

Public Sub FormStepDown(TheForm As Form, Steps As Long)
    On Error Resume Next
    Dim AddX, AddY As Boolean
    Dim theBackColor&, Index&, theX&, theY&
    theBackColor& = TheForm.BackColor
    TheForm.BackColor = RGB(0, 0, 0)
    For Index& = 0 To TheForm.Count - 1
        TheForm.Controls(Index&).Visible = False
    Next Index&
    AddX = True: AddY = True
    TheForm.Show
    theX& = ((Screen.Width - TheForm.Width) - TheForm.Left) / Steps&
    theY& = ((Screen.Height - TheForm.Height) - TheForm.Top) / Steps&
    Do
        TheForm.Move TheForm.Left + theX&, TheForm.Top + theY&
    Loop Until (TheForm.Left >= (Screen.Width - TheForm.Width)) Or (TheForm.Top >= (Screen.Height - TheForm.Height))
    TheForm.Left = Screen.Width - TheForm.Width
    TheForm.Top = Screen.Height - TheForm.Height
    TheForm.BackColor = theBackColor&
    For Index& = 0 To TheForm.Count - 1
        TheForm.Controls(Index&).Visible = True
    Next Index&
End Sub
Public Function ReturnStringElite(Text$) As String
    Dim TempString As String, Length As Long, Spaces As Long
    Dim NextChr As String, NewString As String, NumSpc As Long
    TempString$ = Text$
    Length& = Len(TempString$)
    Do While NumSpc& <= Length&
        NumSpc& = NumSpc& + 1
        NextChr$ = Mid$(TempString$, NumSpc&, 1)
        If NextChr$ = "a" Then NextChr$ = "â"
        If NextChr$ = "b" Then NextChr$ = "b"
        If NextChr$ = "c" Then NextChr$ = "ç"
        If NextChr$ = "e" Then NextChr$ = "ë"
        If NextChr$ = "i" Then NextChr$ = "î"
        If NextChr$ = "j" Then NextChr$ = "j"
        If NextChr$ = "n" Then NextChr$ = "ñ"
        If NextChr$ = "o" Then NextChr$ = "õ"
        If NextChr$ = "s" Then NextChr$ = "š"
        If NextChr$ = "t" Then NextChr$ = "†"
        If NextChr$ = "u" Then NextChr$ = "ü"
        If NextChr$ = "w" Then NextChr$ = "vv"
        If NextChr$ = "y" Then NextChr$ = "ÿ"
        If NextChr$ = "0" Then NextChr$ = "Ø"
        If NextChr$ = "A" Then NextChr$ = "Ã"
        If NextChr$ = "B" Then NextChr$ = "ß"
        If NextChr$ = "C" Then NextChr$ = "Ç"
        If NextChr$ = "D" Then NextChr$ = "Ð"
        If NextChr$ = "E" Then NextChr$ = "Ë"
        If NextChr$ = "I" Then NextChr$ = "Í"
        If NextChr$ = "N" Then NextChr$ = "Ñ"
        If NextChr$ = "O" Then NextChr$ = "Õ"
        If NextChr$ = "S" Then NextChr$ = "Š"
        If NextChr$ = "U" Then NextChr$ = "Û"
        If NextChr$ = "W" Then NextChr$ = "VV"
        If NextChr$ = "Y" Then NextChr$ = "Ý"
        NewString$ = NewString$ & NextChr$
    Loop
    ReturnStringElite$ = NewString$
End Function
Public Sub AOLChatAnnoy(Times As Long)
    Dim theNum As Long
    Do
        theNum& = theNum& + 1
        Call AOLChatSend("{s *a:\spinning}")
    Loop Until theNum& = Times&
End Sub
Public Function ReturnStringDeCoded(sSTR As String) As String
    Dim thestring As String, StringLength As Long, NumSpaces As Long
    Dim NextChr As String, NewString As String
    thestring$ = sSTR$
    StringLength = Len(thestring$)
    Do While NumSpaces& <= StringLength
        NumSpaces& = NumSpaces& + 1
        NextChr$ = Mid$(thestring$, NumSpaces&, 1)
        If NextChr$ = "š" Then NextChr$ = "A"
        If NextChr$ = "œ" Then NextChr$ = "B"
        If NextChr$ = "¢" Then NextChr$ = "C"
        If NextChr$ = "¤" Then NextChr$ = "D"
        If NextChr$ = "±" Then NextChr$ = "E"
        If NextChr$ = "°" Then NextChr$ = "F"
        If NextChr$ = "²" Then NextChr$ = "G"
        If NextChr$ = "³" Then NextChr$ = "H"
        If NextChr$ = "µ" Then NextChr$ = "I"
        If NextChr$ = "ª" Then NextChr$ = "J"
        If NextChr$ = "¹" Then NextChr$ = "K"
        If NextChr$ = "º" Then NextChr$ = "L"
        If NextChr$ = "Ÿ" Then NextChr$ = "M"
        If NextChr$ = "í" Then NextChr$ = "N"
        If NextChr$ = "î" Then NextChr$ = "O"
        If NextChr$ = "ï" Then NextChr$ = "P"
        If NextChr$ = "ð" Then NextChr$ = "Q"
        If NextChr$ = "ñ" Then NextChr$ = "R"
        If NextChr$ = "ò" Then NextChr$ = "S"
        If NextChr$ = "ó" Then NextChr$ = "T"
        If NextChr$ = "ô" Then NextChr$ = "U"
        If NextChr$ = "õ" Then NextChr$ = "V"
        If NextChr$ = "ö" Then NextChr$ = "W"
        If NextChr$ = "ø" Then NextChr$ = "X"
        If NextChr$ = "ù" Then NextChr$ = "Y"
        If NextChr$ = "ú" Then NextChr$ = "Z"
        If NextChr$ = " " Then NextChr$ = " "
        If NextChr$ = "'" Then NextChr$ = "a"
        If NextChr$ = "û" Then NextChr$ = "b"
        If NextChr$ = "ü" Then NextChr$ = "c"
        If NextChr$ = "ý" Then NextChr$ = "d"
        If NextChr$ = "þ" Then NextChr$ = "e"
        If NextChr$ = "Æ" Then NextChr$ = "f"
        If NextChr$ = "Ç" Then NextChr$ = "g"
        If NextChr$ = "Ì" Then NextChr$ = "h"
        If NextChr$ = "Í" Then NextChr$ = "i"
        If NextChr$ = "Î" Then NextChr$ = "j"
        If NextChr$ = "Ï" Then NextChr$ = "k"
        If NextChr$ = "Ø" Then NextChr$ = "l"
        If NextChr$ = "Þ" Then NextChr$ = "m"
        If NextChr$ = "ß" Then NextChr$ = "n"
        If NextChr$ = "†" Then NextChr$ = "o"
        If NextChr$ = "ƒ" Then NextChr$ = "p"
        If NextChr$ = "Œ" Then NextChr$ = "q"
        If NextChr$ = "Š" Then NextChr$ = "r"
        If NextChr$ = "‡" Then NextChr$ = "s"
        If NextChr$ = "¡" Then NextChr$ = "t"
        If NextChr$ = "£" Then NextChr$ = "u"
        If NextChr$ = "§" Then NextChr$ = "v"
        If NextChr$ = "ì" Then NextChr$ = "w"
        If NextChr$ = "ë" Then NextChr$ = "x"
        If NextChr$ = "ê" Then NextChr$ = "y"
        If NextChr$ = "é" Then NextChr$ = "z"
        If NextChr$ = "è" Then NextChr$ = "1"
        If NextChr$ = "ç" Then NextChr$ = "2"
        If NextChr$ = "æ" Then NextChr$ = "3"
        If NextChr$ = "á" Then NextChr$ = "4"
        If NextChr$ = "å" Then NextChr$ = "5"
        If NextChr$ = "â" Then NextChr$ = "6"
        If NextChr$ = "ã" Then NextChr$ = "7"
        If NextChr$ = "ä" Then NextChr$ = "8"
        If NextChr$ = "à" Then NextChr$ = "9"
        If NextChr$ = "×" Then NextChr$ = "0"
        NewString$ = NewString$ & NextChr$
    Loop
    ReturnStringDeCoded$ = NewString$
End Function
Public Function FileExists(FilePath As String) As Boolean
    Dim Check As Long
    If InStr(FilePath$, ".") = 0 Then
        FileExists = False
    End If
    Check& = Len(Dir$(FilePath$))
    If Check& = 0 Then
        FileExists = False
        Exit Function
    Else
        FileExists = True
    End If
End Function
Public Sub FileCopy(file$, destination$)
    If Not FileExists(file$) Then Exit Sub
    If InStr(file$, ".") = 0 Then Exit Sub
    If InStr(destination$, "\") = 0 Then Exit Sub
    Call FileCopy(file$, destination$)
End Sub
Public Sub FileDelete(file$)
    Dim NoFreeze As Long
    If Not FileExists(file$) Then Exit Sub
    If InStr(file$, ".") = 0 Then Exit Sub
    Call Kill(file$)
    NoFreeze& = DoEvents()
End Sub
Public Sub FileRename(file$, Path$, NewName$)
    Dim NoFreeze As Long
    If InStr(file$, ".") = 0 Then Exit Sub
    Name Path$ & file$ As NewName$
    NoFreeze& = DoEvents()
End Sub
Public Function FileGetAttributes(FilePath As String) As Integer
    Dim Check As Long
    Check& = Dir(FilePath$)
    If Check& = 0 Then MsgBox "File Not Found", vbCritical, "FrENzY32.bas": Exit Function
    FileGetAttributes = GetAttr(FilePath$)
End Function
Public Function INIRead(Section As String, Key As String, FullPath As String)
   Dim Buffer As String
   Buffer$ = String(750, Chr(0))
   INIRead = Left(Buffer, GetPrivateProfileString(Section$, ByVal LCase(Key$), "", Buffer, Len(Buffer), FullPath$))
End Function
Public Sub INIWrite(Section As String, Key As String, KeyValue As String, FullPath As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, FullPath$)
End Sub
Public Sub AOLRunMenu(MainMenu As Long, TheSubMenu As Long)
    Dim AOL As Long, Menu As Long
    Dim SubMenu As Long, MenuID As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    Menu& = GetMenu(AOL&)
    SubMenu& = GetSubMenu(Menu&, MainMenu&)
    MenuID& = GetMenuItemID(SubMenu&, TheSubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, MenuID&, 0&)
End Sub
Public Sub WAVStop()
    Call sndPlaySound("", SND_FLAG)
End Sub

Public Sub WAVLoop(file As String)
    Dim SoundName$, theFlags&
    SoundName$ = file$
    theFlags& = SND_ASYNC Or SND_LOOP
    Call sndPlaySound(SoundName$, theFlags&)
End Sub
Public Sub WAVLoopStop()
    Dim theFlags As Long
    theFlags& = SND_ASYNC Or SND_LOOP
    Call sndPlaySound("", theFlags&)
End Sub

Public Sub WindowsShutdown()
    Dim EWX_SHUTDOWN, GetMsg As Long
    GetMsg& = MsgBox("Do you really want to Shut Down Windows 95/98?", vbYesNo Or vbQuestion)
    If GetMsg& = vbNo Then
        Exit Sub
    Else
        Call ExitWindowsEx(EWX_SHUTDOWN, 0)
    End If
End Sub
Public Function ListFindString(List As ListBox, FindString As String) As Long
    Dim Index As Long
    If List.ListCount = 0 Then MsgBox "There is no item in " & List & ", that matches your specifications"
    For Index& = 0 To List.ListCount - 1
        List.ListIndex = Index&
        If UCase(List.Text) = UCase(FindString$) Then
            ListFindString& = Index&
            Exit Function
            If Err Then MsgBox "There is no item in " & List & ", that matches your specifications"
        End If
    Next Index&
End Function
Public Function INIPath(ININame)
    If InStr(ININame, ".ini") = 0 Then
        INIPath = App.Path & "\" & ININame & ".ini"
    Else
        INIPath = App.Path & "\" & ININame
    End If
End Function

Public Sub AOLRunMenuByString(SearchString As String)
    Dim AOL As Long, Menu As Long, MenuCount As Long
    Dim X As Long, SubMenu As Long, SubMenuID As Long
    Dim SubMenuCount As Long, I As Long, Stringer As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    Menu& = GetMenu(AOL&)
    MenuCount& = GetMenuItemCount(Menu&)
    For X = 0 To MenuCount& - 1
        SubMenu& = GetSubMenu(Menu&, X&)
        SubMenuCount& = GetMenuItemCount(SubMenu&)
        For I = 0 To SubMenuCount& - 1
            SubMenuID& = GetMenuItemID(SubMenu&, I)
            Stringer$ = String$(100, " ")
            Call GetMenuString(SubMenu&, SubMenuID&, Stringer$, 100&, 1)
            If InStr(LCase(Stringer$), LCase(SearchString$)) Then
                Call SendMessageLong(AOL&, WM_COMMAND, SubMenuID&, 0&)
                Exit Sub
            End If
        Next I
    Next X
End Sub

Public Sub AOLSNListLoad(List As ListBox, CommonDlg As Control)
    Dim sSNList As String, sChar As String, sSN As String, lPos As Long
    With CommonDlg
        .DialogTitle = "Load SN List"
        .CancelError = True
        .Filter = "Text File (*.txt)|*.txt"
        .FilterIndex = 0
        .ShowOpen
    End With
    Open CommonDlg.FileName For Input As #1
        sSNList$ = Input(LOF(1), 1)
    Close #1
    Let sSN$ = ""
    For lPos& = 1 To Len(sSNList$)
        sChar$ = Mid$(sSNList$, lPos&, 1)
        If sChar$ = "," Then
            List.AddItem sSN$
            Let sSN$ = ""
        Else
             sSN$ = sSN$ & sChar$
        End If
    Next lPos&
    Exit Sub
End Sub
Public Sub AOLSNListSave(List As ListBox, CommonDlg As Control)
    Dim sList As String, lSN As Long
    With CommonDlg
        .CancelError = True
        .DialogTitle = "Save SNs"
        .Filter = "Text Files (*.txt)|*.txt"
        .FilterIndex = 0
        .ShowSave
    End With
    sList$ = ""
    For lSN& = 0 To List.ListCount - 1
        If lSN& = 0 Then
            sList$ = List.List(lSN&)
        Else
            sList$ = sList$ & "," & List.List(lSN&)
        End If
    Next lSN&
    Open CommonDlg.FileName For Output As #1
        Print #1, sList$
    Close #1
    Exit Sub
End Sub

Public Sub ScanForPWS(FilePath$, FileName$)
    Dim FileLen As Long, FileInfo As String, NumOne, GenOiZBack, GenOziDe, TheFileInfo$, PWS, PWS2, PWS3, VirusedFile, LengthOfFile, TotalRead, TheTab, TheMSg, TheMsg2, TheMsg3, TheMsg4, TheMsg5, TheDots, StopPWScanner As Boolean, PentiumRest As Long
    If FileName$ = "" Then Exit Sub
    If Right(FilePath$, 1) = "\" Then
        FileName$ = FilePath$ & FileName$
    Else
        FileName$ = FilePath$ & "\" & FileName$
    End If
    If Not FileExists(FileName$) Then MsgBox "File Not Found", 16, "Error": Exit Sub
    FileLen& = Len(FileName$)
    Open FileName$ For Binary As #1
        If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": Exit Sub
        FileInfo$ = String(32000, 0)
        Get #1, 1, FileInfo$
    Close #1
    Open FileName$ For Binary As #2
        If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": Exit Sub
        If InStr(1, LCase$(FileInfo$), "main.idx" & Chr(0), 1) Then
            MsgBox "This file is infected with a PWS"
        End If
End Sub

Public Function ReplaceString(thestring As String, ReplaceWhat As String, WithWhat As String)
    Dim Position As Long
    Do While InStr(1, thestring$, ReplaceWhat$)
        DoEvents
        Position& = InStr(1, thestring$, ReplaceWhat$)
        thestring$ = Left(thestring$, (Position& - 1)) & WithWhat$ & Right(thestring$, Len(thestring$) - (Position& + Len(ReplaceWhat$) - 1))
    Loop
    ReplaceString = thestring$
End Function


Public Sub ScanForDeltree(FilePath$, FileName$)
    Dim FileLen As Long, FileInfo As String
    If FileName$ = "" Then Exit Sub
    If Right(FilePath$, 1) = "\" Then
        FileName$ = FilePath$ & FileName$
    Else
        FileName$ = FilePath$ & "\" & FileName$
    End If
    If Not FileExists(FileName$) Then MsgBox "File Not Found", 16, "Error": Exit Sub
    FileLen& = Len(FileName$)
    Open FileName$ For Binary As #1
        If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": Exit Sub
        FileInfo$ = String(32000, 0)
        Get #1, 1, FileInfo$
    Close #1
    Open FileName$ For Binary As #2
        If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": Exit Sub
        If InStr(1, LCase$(FileInfo$), "deltree" & Chr(0), 1) Then
            MsgBox "This file is infected with a deltree"
        End If
        If InStr(1, LCase$(FileInfo$), "kill" & Chr(0), 1) Then
            MsgBox "This file is infected with a deltree"
        End If
End Sub
