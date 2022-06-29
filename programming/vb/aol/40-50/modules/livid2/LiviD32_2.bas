Attribute VB_Name = "LiviD32_2"
'       ***************************
'       *'livid32.bas v2 by livid'*
'       ***************************
'mail me at Lividx@aol.com
'32 bit module made for aol 4.0
Option Explicit
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function auxGetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long
Public Declare Function auxSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CascadeWindows Lib "user32" (ByVal hWndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpkids As Long) As Integer
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function CreateMDIWindow Lib "user32" Alias "CreateMDIWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hInstance As Long, ByVal lParam As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOPENFILENAME As OPENFILENAME) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function TileWindows Lib "user32" (ByVal hWndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpkids As Long) As Integer
Public Declare Sub RtlMoveMemory Lib "kernel32" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_ENABLE = &HA
Public Const WM_GETFONT = &H31
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFOCUS = &H7
Public Const WM_SETFONT = &H30
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SYSCOMMAND = &H112
Public Const WM_USER = &H400
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETSEL = &HB0
Public Const EM_LIMITTEXT = &HC5
Public Const EM_REPLACESEL = &HC2
Public Const EN_CHANGE = &H300
Public Const LB_SETCURSEL = &H186
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT = &H18B
Public Const GW_CHILD = 5
Public Const GW_hWndFIRST = 0
Public Const GW_hWndLAST = 1
Public Const GW_hWndNEXT = 2
Public Const GW_hWndPREV = 3
Public Const GWL_WNDPROC = -4
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
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
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const SC_MOVE = &HF011
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NODEFAULT = &H2
Public Const SND_NOSTOP = &H10
Public Const SND_NOWAIT = &H2000
Public Const SND_PURGE = &H40
Public Const SND_SYNC = &H0
Public Const SND_VALIDFLAGS = &H17201F
Public Const SW_HIDE = 0
Public Const SW_Maximize = 3
Public Const SW_Minimize = 6
Public Const SW_NORMAL = 1
Public Const SW_Restore = 9
Public Const SW_SHOW = 5
Public Const SPI_SETDOUBLECLICKTIME = 32
Public Const SPI_SETCURSORS = 87
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SPI_GETSCREENSAVEACTIVE = 16
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOREDRAW = &H8
Public Const VK_DOWN = &H28
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Global PrevWndProc As Long
Global ghWnd As Long
Global StopScrolling As Boolean
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
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
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public nid As NOTIFYICONDATA

Public Sub Pause(Seconds As Long)
    Dim a As Long, b As Long
    a = Timer
    Do
        DoEvents
        b = Timer
        If b - a >= Seconds Then
            Exit Do
        End If
    Loop
End Sub
Public Function ReplaceChar(ToSearch As String, ToReplace As String, ToReplaceWith As String)
    Dim a As Long, b As Long
    Dim c As String, d As String
    Dim X As Long
    b = 1
    d = ""
    For X = 1 To Len(ToSearch)
        c = Mid$(ToSearch, X, 1)
        If c = ToReplace Then
            d = d + ToReplaceWith
        Else
            d = d + c
        End If
    Next X
    ReplaceChar = d
End Function
Public Function AOLWindow() As Long
    AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function
Public Function AOLMDI() As Long
    AOLMDI = FindWindowEx(AOLWindow, 0, "MDIClient", vbNullString)
End Function
Public Sub SetCaption(hwnd As Long, NewTitle As String)
    Call SetWindowText(hwnd, NewTitle)
End Sub
Public Sub SetText(hwnd As Long, ToSet As String)
    Call SendMessageByString(hwnd, WM_SETTEXT, 0, "")
    Call SendMessageByString(hwnd, WM_SETTEXT, 0, ToSet)
End Sub
Public Function GetCaption(hwnd As Long) As String
    Dim a As Long, b As String
    a = GetWindowTextLength(hwnd)
    b = String(a, " ")
    GetCaption = GetWindowText(hwnd, b, 0)
End Function
Public Function GetText(hwnd As Long)
    Dim a As Long, b As String, c As Long
    a = SendMessageByNum(hwnd, WM_GETTEXTLENGTH, 0, 0)
    b = Space$(a)
    c = SendMessageByString(hwnd, WM_GETTEXT, a + 1, b)
    GetText = b
End Function
Public Sub Click(hwnd As Long)
    Call SendMessage(hwnd, WM_LBUTTONDOWN, 0, 0)
    Call SendMessage(hwnd, WM_LBUTTONUP, 0, 0)
End Sub

Public Sub IMSend(ToWho As String, ToSay As String)
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long
    'this finds the im then fills it out..
    'THIS DOESNT OPEN THE Instant Message
    If IsUserOnline = False Then Exit Sub
    Do
        a = FindWindowEx(AOLMDI, 0, "AOL Child", "Send Instant Message")
        b = FindWindowEx(a, 0, "_AOL_Edit", vbNullString)
        c = FindWindowEx(a, 0, "RICHCNTL", vbNullString)
        DoEvents
    Loop Until a <> 0 And b <> 0 And c <> 0
    Call SetText(b, ToWho)
    Call SetText(c, ToSay)
    d = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    For e = 1 To 8
        d = FindWindowEx(a, d, "_AOL_Icon", vbNullString)
    Next e
    Call Click(d)
    Do
        DoEvents
        a = FindWindowEx(AOLMDI, 0, "AOL Child", "Send Instant Message")
        e = FindWindow("#32770", "America Online")
        If e <> 0 Then
            CloseWindow (e)
            CloseWindow (a)
        End If
    Loop Until a = 0
    If Left(TrimSpaces(LCase(ToWho)), 7) = "$im_off" Then Exit Sub
    If Left(TrimSpaces(LCase(ToWho)), 6) = "$im_on" Then Exit Sub
    Do
        DoEvents
        a = FindChildByTitle(AOLMDI, "  Instant Message To: ")
    Loop Until a <> 0
    Call CloseWindow(a)
End Sub
Public Function TrimSpaces(ToTrim As String) As String
    Dim a As Long, b As String, c As String
    For a = 1 To Len(ToTrim)
        DoEvents
        b = Mid$(ToTrim, a, 1)
        If b = " " Then GoTo NowNext
        c = c + b
NowNext:
    Next a
    TrimSpaces = c
End Function

Public Function ChatFindRoom()
    Dim a As Long, b As Long
    Dim c As Long, d As Long
    Dim e As Long, f As Long
    Dim g As Long, h As Long
    h = FindWindowEx(AOLMDI, b, "AOL Child", vbNullString)
    For a = 0 To GetAOLChildCount + 10
        b = h
        c = FindWindowEx(b, 0, "RICHCNTL", vbNullString)
        d = FindWindowEx(b, c, "RICHCNTL", vbNullString)
        e = FindWindowEx(b, 0, "_AOL_ListBox", vbNullString)
        f = FindWindowEx(b, 0, "_AOL_Icon", vbNullString)
        h = GetWindow(b, GW_hWndNEXT)
        If b <> 0 And c <> 0 And d <> 0 And e <> 0 And f <> 0 Then GoTo FoundIt
    Next a
FoundIt:
    ChatFindRoom = b
End Function
Public Sub ChatSend(ToSend As String)
    Dim a As Long, b As String
    Dim c As String, d As Long
    If AOLWindow = 0 Then Exit Sub
    If ChatFindRoom = 0 Then Exit Sub
    a = ChatSendBox
    d = GetFocus
    b = GetText(a)
    Call SetText(a, ToSend)
    Do
        Call SendMessageByNum(a, WM_CHAR, 13, 0)
        c = GetText(a)
        Call Pause(0.1)
    Loop Until c <> ToSend
    Call SetText(a, b)
    Call SetFocus(d)
End Sub
Public Sub MailSend(ToWho As String, subject As String, ToSay As String)
    Dim a As Long, f As Long, g As Long
    If AOLWindow = 0 Then Exit Sub
    If IsUserOnline = False Then Exit Sub
    a = GetToolBarIcon(2)
    Call Click(a)
    Dim b As Long, c As Long, d As Long, e As Long, h As Long, i As Long
    Do
        DoEvents
        b = FindWindowEx(AOLMDI, 0, "AOL Child", "Write Mail")
        c = FindWindowEx(b, 0, "_AOL_Edit", vbNullString)
        h = FindWindowEx(b, c, "_AOL_Edit", vbNullString)
        i = FindWindowEx(b, h, "_AOL_Edit", vbNullString)
        d = FindWindowEx(b, 0, "RICHCNTL", vbNullString)
        e = FindWindowEx(b, 0, "_AOL_Icon", vbNullString)
    Loop Until b <> 0 And c <> 0 And e <> 0 And h <> 0 And i <> 0 And d <> 0
    Call SetText(c, ToWho)
    Call SetText(i, subject)
    Call SetText(d, ToSay)
    Dim j As Long, k As Long
    Do
        b = FindWindowEx(AOLMDI, 0, "AOL Child", "Write Mail")
        e = 0
        For j = 1 To 14
            e = FindWindowEx(b, e, "_AOL_Icon", vbNullString)
        Next j
        Call Click(e)
        DoEvents
        Pause 0.2
        k = FindWindowEx(AOLMDI, 0, "AOL Child", "Error")
        If k <> 0 Then
            Call MsgBoxOnTop("The Mail Can't Be Sent Due to People Who Can't Recieve Mail From You.", vbOKOnly + vbCritical, App.Title)
            Exit Sub
        End If
    Loop Until b = 0
End Sub
Public Function GetToolBarIcon(Index As Long) As Long
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindWindowEx(AOLWindow, 0, "AOL Toolbar", vbNullString)
    b = FindWindowEx(a, 0, "_AOL_Toolbar", vbNullString)
    c = 0
    For d = 1 To Index
        c = FindWindowEx(b, c, "_AOL_Icon", vbNullString)
    Next d
    GetToolBarIcon = c
End Function
    
Public Function GetWinParent(hwnd As Long) As Long
    Dim a As Long
    a = GetParent(hwnd)
    GetWinParent = a
End Function

Public Function ChatLastLine() As String
    Dim a As String, b As String
    a = ChatText
    b = GetLastLine(a)
    ChatLastLine = b
End Function

Public Sub Keyword(ToWhere As String)
    Dim a As Long, b As Long
    Dim c As Long, d As Long
    a = FindWindowEx(AOLWindow, 0, "AOL Toolbar", vbNullString)
    b = FindWindowEx(a, 0, "_AOL_Toolbar", vbNullString)
    c = FindWindowEx(b, 0, "_AOL_Combobox", vbNullString)
    d = FindWindowEx(c, 0, "Edit", vbNullString)
    Call SetText(d, ToWhere)
    Call SendMessageByNum(d, WM_CHAR, Asc(" "), 0)
    Call SendMessageByNum(d, WM_CHAR, 13, 0)
End Sub
Public Sub RunPopUpMenu(IconNum As Long, Letter As String)
    'u lazy bastards can thank DeLTa for this one
    Dim a As Long, b As Long
    Dim c As Long
    a = FindWindow("#32768", vbNullString)
    b = GetToolBarIcon(IconNum)
    Call PostMessage(b, WM_LBUTTONDOWN, 0, 0)
    Call PostMessage(b, WM_LBUTTONUP, 0, 0)
    Call SetCursorPos(0, 0)
    Do
        DoEvents
        c = FindWindow("#32768", vbNullString)
    Loop Until c <> a
    Call SetCursorPos(0, 0)
    Call PostMessage(c, WM_CHAR, Asc(Letter), 0)
End Sub
Public Sub IMKeyword(ToWho As String, ToSay As String)
    Call Keyword("aol://9293:")
    Call IMSend(ToWho, ToSay)
End Sub
Public Sub IMIgnore(ToIgnore As String)
    Call IMKeyword("$IM_OFF " + ToIgnore, Chr(9))
End Sub
Public Sub IMUnIgnore(ToUnignore As String)
    Call IMKeyword("$IM_ON " + ToUnignore, Chr(9))
End Sub
Public Sub IMsOn()
    Call IMKeyword("$IM_ON", Chr(9))
End Sub
Public Sub IMsOff()
    Call IMKeyword("$IM_OFF", Chr(9))
End Sub
Public Sub IMPopUpMenu(ToWho As String, ToSay As String)
    Call RunPopUpMenu(10, "i")
    Call IMSend(ToWho, ToSay)
End Sub
Public Sub IMBuddy(ToWho As String, ToSay As String)
    Dim a As Long, b As Long, c As Long
Start:
    DoEvents
    a = FindWindowEx(AOLMDI, 0, "AOL Child", "Buddy List Window")
    If a = 0 Then Call OpenBL: GoTo Start
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    c = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    Call Click(c)
    Call IMSend(ToWho, ToSay)
End Sub
Public Sub CloseWindow(hwnd As Long)
    Call SendMessage(hwnd, WM_CLOSE, 0, 0)
End Sub
    
Public Sub APIListToVBList(hwnd As Long, Lst As ListBox, TrimForMail As Boolean, AddNumbers As Boolean)
    Dim a As Long, b As Long, c As String
    Dim d As Long, e As Long, f As String
    Dim g As Long, h As String
    a = SendMessage(hwnd, LB_GETCOUNT, 0, 0)
    For b = 0 To a - 1 Step 1
        DoEvents
        d = SendMessage(hwnd, LB_GETTEXTLEN, b, 0)
        c = Space(d + 1)
        Call SendMessageByString(hwnd, LB_GETTEXT, b, c)
        If TrimForMail = True Then
            e = InStr(1, c, Chr(9))
            f = InStr(e + 1, c, Chr(9)) + 1
            c = Mid$(c, f)
        End If
        If AddNumbers = True Then
            c = "(" + TrimSpaces(Str(b + 1)) + ") " + c
        End If
        Lst.AddItem c
    Next b
End Sub
    
Public Function GetLastLine(OfWhat As String) As String
    Dim a As String, b As Long, c As String
    c = OfWhat
    b = GetLineCount(c)
    a = GetLine(c, b)
    GetLastLine = a
End Function
Public Sub INIWrite(Section As String, Keyword As String, NewWord As String, Path As String)
    Call WritePrivateProfileString(Section, Keyword, NewWord, Path)
End Sub
    
Public Function INIRead(Section As String, Keyword As String, Path As String) As String
    Dim a As String
        a = String(255, Chr(0))
        INIRead = Left(a, GetPrivateProfileString(Section, Keyword, "", a, Len(a), Path))
End Function
    
Public Sub StayOnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub
    
Public Sub NotOnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
End Sub
    
    
Function GetClass(hwnd)
    Dim a As String, b As Long
    a = String$(250, 0)
    b = GetClassName(hwnd, a, 250)
    GetClass = a
End Function
Public Sub DisableCtrlAltDel()
    Dim a As Long
    a = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, 0&, 0)
End Sub
Public Sub EnableCtrlAltDel()
    Dim a As Long
    a = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, 0&, 0)
End Sub
Public Sub MoveForm(Frm As Form)
    Call ReleaseCapture
    Dim a As Long
    a = SendMessage(Frm.hwnd, WM_SYSCOMMAND, SC_MOVE, 0)
End Sub
Public Sub FormMove(Frm As Form)
    Call ReleaseCapture
    Dim a As Long
    a = SendMessage(Frm.hwnd, WM_SYSCOMMAND, SC_MOVE, 0)
End Sub
Public Function FindChildByTitle(hwnd As Long, ChildTitle As String)
    Dim a As Long, b As Long, c As Long
    a = GetWindow(hwnd, GW_CHILD)
    If UCase(GetText(a)) Like UCase(ChildTitle) Then GoTo FoundChild
    a = GetWindow(hwnd, GW_CHILD)
    Do While a <> 0
        b = GetWindow(hwnd, GW_CHILD)
        If UCase(GetText(b)) Like UCase(ChildTitle) & "*" Then GoTo FoundChild
        a = GetWindow(a, GW_hWndNEXT)
        If UCase(GetText(a)) Like UCase(ChildTitle) & "*" Then GoTo FoundChild
    Loop
    FindChildByTitle = 0
FoundChild:
    c = a
    FindChildByTitle = c
End Function
Public Function AOLUser()
    On Error Resume Next
    Dim a As Long, b As String
    a = FindChildByTitle(AOLMDI, "Welcome, ")
    If a = 0 Then
        AOLUser = ""
        Exit Function
    End If
    b = GetText(a)
    b = Mid$(b, 10)
    b = Left(b, Len(b) - 1)
    AOLUser = b
End Function
Public Sub IMCloseIMTo()
    'you can put this in a timer and it will kill all the
    'Instant Message To: windows
    'use this in a timer:
    'Call IMCloseIMTo
    Dim a As Long
KillIMTo:
    a = FindChildByTitle(AOLMDI, "  Instant Message To: ")
    If a = 0 Then Exit Sub
    Call CloseWindow(a)
    GoTo KillIMTo
End Sub
Public Function GetAPILineCount(hwnd As Long) As Long
    GetAPILineCount = SendMessage(hwnd, EM_GETLINECOUNT, 0, 0)
End Function
Public Function CountChr(ToCountIn As String, ToCount As String, CaseSensitive As Boolean) As Long
    Dim a As Long, b As String, c As Long
    c = 0
    For a = 1 To Len(ToCountIn)
        b = Mid$(ToCountIn, a, 1)
        If CaseSensitive = True Then
            If b = ToCount Then
                c = c + 1
            End If
        ElseIf CaseSensitive = False Then
            If LCase(b) = LCase(ToCount) Then
                c = c + 1
            End If
        End If
    Next a
    CountChr = c
End Function
Public Function GetLineCount(ToCountIn As String) As Long
    GetLineCount = CountChr(ToCountIn, Chr(13), False) + 1
End Function
Public Function GetAPILine(hwnd As Long, LineNumber As Long) As String
    Dim a As String, b As Long
    a = String(250, " ")
    b = SendMessage(hwnd, EM_GETLINE, LineNumber, a)
    GetAPILine = a
End Function
Public Sub SysTrayAdd(FormToMin As Form, ToolTip As String)
    FormToMin.Show
    FormToMin.WindowState = 1
    FormToMin.Refresh
    With nid
        .cbSize = Len(nid)
        .hwnd = FormToMin.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = FormToMin.Icon
        .szTip = ToolTip & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
    FormToMin.Hide
End Sub
Public Sub SysTrayRemove(FormToKill As Form)
    Shell_NotifyIcon NIM_DELETE, nid
    FormToKill.Visible = True
    FormToKill.WindowState = 0
End Sub
Public Function FileExist(ThePath As String) As Boolean
    Dim a As String
    If TrimSpaces(ThePath) = "" Then FileExist = False: Exit Function
    a = Dir(ThePath)
    If Len(a) = 0 Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function
Public Sub FileDelete(Path As String)
    If FileExist(Path) = True Then Call Kill(Path)
End Sub
Public Sub KillWait()
    Dim a As Long, b As Long
    Call RunMenu(AOLWindow, 4, 10)
    Do
        a = FindWindow("_AOL_Modal", vbNullString)
        b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
        DoEvents
    Loop Until b <> 0
    Do
        a = FindWindow("_AOL_Modal", vbNullString)
        b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
        Call Click(b)
        DoEvents
    Loop Until a = 0
End Sub
Public Sub RunMenu(hwnd As Long, FirstMenu As Long, SecondMenu As Long)
    Dim a As Long, b As Long, c As Long
    a = GetMenu(hwnd)
    b = GetSubMenu(a, FirstMenu)
    c = GetMenuItemID(b, SecondMenu)
    Call SendMessageByNum(hwnd, WM_COMMAND, c, 0&)
End Sub
Public Sub RunMenuByString(hwnd As Long, TheString As String)
    'i suggest using regular runmenu, its faster
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As String, h As Long
    Dim i As Long
    a = GetMenu(hwnd)
    b = GetMenuItemCount(a)
    For c = 0 To b
        d = GetSubMenu(a, c)
        e = GetMenuItemCount(d)
        For f = 0 To d
            g = String(250, " ")
            h = GetMenuItemID(d, f)
            i = GetMenuString(d, h, g, 250, 1)
            If LCase(Left(g, Len(TheString))) = LCase(TheString) Then
                Call SendMessageByNum(hwnd, WM_COMMAND, h, 0)
                Exit Sub
            End If
        Next f
    Next c
End Sub
Public Sub ChatAddRoomList(Lst As ListBox, AddUserSn As Boolean)
    On Error Resume Next
    Dim a As Long, b As Long, c As String, d As Long
    Dim e As Long, f As Long, g As Long, h As Long
    Dim i As Long, j As Long, k As Long
    g = ChatFindRoom
    If g = 0 Then Exit Sub
    h = FindWindowEx(g, 0, "_AOL_Listbox", vbNullString)
    i = GetWindowThreadProcessId(h, a)
    j = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, a)
    For f = 0 To SendMessage(h, LB_GETCOUNT, 0, 0) - 1
        c = String(4, vbNullChar)
        b = SendMessage(h, LB_GETITEMDATA, ByVal CLng(f), ByVal 0) + 24
        Call ReadProcessMemory(j, b, c, 4, e)
        Call RtlMoveMemory(d, ByVal c, 4)
        d = d + 6
        c = String(16, vbNullChar)
        Call ReadProcessMemory(j, d, c, Len(c), e)
        k = InStr(c, vbNullChar)
        c = Left(c, k - 1)
        If TrimSpaces(c) = TrimSpaces(AOLUser) Then
            If AddUserSn = True Then
                Lst.AddItem c
            End If
        Else
            Lst.AddItem c
        End If
    Next f
    Call CloseHandle(j)
End Sub
Public Sub FormCenter(Frm As Form)
    Dim a As Long
    a = (Screen.Height / 2) - (Frm.Height / 2)
    Frm.Top = a
    a = (Screen.Width / 2) - (Frm.Width / 2)
    Frm.Left = a
End Sub
Public Sub ChatAddRoomCombo(Comb As ComboBox, AddUserSn As Boolean)
    On Error Resume Next
    Dim a As Long, b As Long, c As String, d As Long
    Dim e As Long, f As Long, g As Long, h As Long
    Dim i As Long, j As Long, k As Long
    g = ChatFindRoom
    If g = 0 Then Exit Sub
    h = FindWindowEx(g, 0, "_AOL_Listbox", vbNullString)
    i = GetWindowThreadProcessId(h, a)
    j = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, a)
    If j = 0 Then Exit Sub
    For f = 0 To SendMessage(h, LB_GETCOUNT, 0, 0) - 1
        c = String(4, vbNullChar)
        b = SendMessage(h, LB_GETITEMDATA, ByVal CLng(f), ByVal 0) + 24
        Call ReadProcessMemory(j, b, c, 4, e)
        Call RtlMoveMemory(d, ByVal c, 4)
        d = d + 6
        c = String(16, vbNullChar)
        Call ReadProcessMemory(j, d, c, Len(c), e)
        k = InStr(c, vbNullChar)
        c = Left(c, k - 1)
        If c = AOLUser Then
            If AddUserSn = True Then
                Comb.AddItem c
            End If
        End If
    Next f
    Call CloseHandle(j)
End Sub
    
Public Sub KillListDupes(Lst As ListBox)
    Dim a As Long, b As String, c As Long
    If Lst.ListCount = 0 Then Exit Sub
    If Lst.ListCount = 1 Then Exit Sub
Start:
    For a = 0 To Lst.ListCount - 1
        b = Lst.List(a)
        If a = Lst.ListCount - 1 Then Exit Sub
        For c = (a + 1) To Lst.ListCount - 1
            If b = Lst.List(c) Then
                Lst.RemoveItem c
                GoTo Start
            End If
        Next c
    Next a
End Sub
Public Sub MailSetPrefs()
    Call RunPopUpMenu(3, "P")
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long
    DoEvents
    Do
        DoEvents
        a = FindWindow("_AOL_Modal", "Mail Preferences")
        b = FindWindowEx(a, 0, "_AOL_Checkbox", vbNullString)
        c = FindWindowEx(a, b, "_AOL_Checkbox", vbNullString)
        d = FindWindowEx(a, c, "_AOL_Checkbox", vbNullString)
        d = FindWindowEx(a, d, "_AOL_Checkbox", vbNullString)
        d = FindWindowEx(a, d, "_AOL_Checkbox", vbNullString)
        d = FindWindowEx(a, d, "_AOL_Checkbox", vbNullString)
    Loop Until a <> 0 And b <> 0 And c <> 0 And d <> 0
    DoEvents
    Call SendMessage(b, BM_SETCHECK, False, 0)
    Call SendMessage(c, BM_SETCHECK, True, 0)
    Call SendMessage(d, BM_SETCHECK, False, 0)
    Do
        DoEvents
        a = FindWindow("_AOL_Modal", "Mail Preferences")
        e = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
        Call Click(e)
    Loop Until a = 0
End Sub
Public Sub MailRemoveFwd()
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As String
    a = FindChildByTitle(AOLMDI, "Fwd: ")
    b = FindWindowEx(a, 0, "_AOL_Edit", vbNullString)
    c = FindWindowEx(a, b, "_AOL_Edit", vbNullString)
    d = FindWindowEx(a, c, "_AOL_Edit", vbNullString)
    e = GetText(d)
    If InStr(1, e, "Fwd: ") = 0 Then Exit Sub
    e = Mid$(e, 6)
    Call SetText(d, "")
    Call SetText(d, e)
End Sub
Public Sub AOLTileWindows()
    Call RunMenu(AOLWindow, 2, 1)
End Sub
Public Sub AOLCascadeWindows()
    Call RunMenu(AOLWindow, 2, 0)
End Sub
Public Function SignOnScreen() As Long
    Dim a As Long, b As String, c As Long
    a = FindWindowEx(AOLMDI, 0, "AOL Child", vbNullString)
    b = LCase(GetText(a))
    If b = "sign on" Or b = "goodbye from america online!" Then
        GoTo FoundScreen
    Else
        For c = 0 To GetAOLChildCount + 5
            a = FindWindowEx(AOLMDI, a, "AOL Child", vbNullString)
            b = LCase(GetText(a))
            If b = "sign on" Or b = "goodbye from america online!" Then
                GoTo FoundScreen
            End If
        Next c
    End If
FoundScreen:
    SignOnScreen = a
End Function
Public Sub SelectGuest()
    'this will select guest from the login menu
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long
    a = SignOnScreen
    b = FindWindowEx(a, 0, "_AOL_Combobox", vbNullString)
    Call SetCursorPos(0, 0)
    Call PostMessage(b, WM_LBUTTONDOWN, 0, 0)
    Call PostMessage(b, WM_LBUTTONUP, 0, 0)
    Do
        DoEvents
        c = FindWindow("#32769", vbNullString)
        d = FindWindowEx(d, 0, "ComboLBox", vbNullString)
    Loop Until d <> 0
    Call SetCursorPos(0, 0)
    For e = 1 To 8
        Call PostMessage(d, WM_KEYDOWN, VK_DOWN, 0)
        Call PostMessage(d, WM_KEYUP, VK_DOWN, 0)
    Next e
    Call Click(b)
End Sub
Public Function CheckPhish(TheName As String, ThePW As String) As Boolean
    Dim a As Long, b As Long, c As Long
    Dim d As Long, e As Long, f As Long
    Dim g As Long, h As Long, i As Long
    Dim j As String, k As Long, l As String
    Dim m As Long
    If SignOnScreen = 0 Then Exit Function
    If IsUserOnline = True Then Exit Function
    Call SelectGuest
    a = SignOnScreen
    c = 0
    For b = 1 To 4
        c = FindWindowEx(a, c, "_AOL_Icon", vbNullString)
    Next b
    Call Click(c)
    Do
        DoEvents
        d = FindWindow("_AOL_Modal", vbNullString)
        e = FindWindowEx(d, 0, "_AOL_Icon", vbNullString)
        f = FindWindowEx(d, 0, "_AOL_Static", vbNullString)
        g = FindWindowEx(d, 0, "_AOL_Edit", vbNullString)
        h = FindWindowEx(d, g, "_AOL_Edit", vbNullString)
    Loop Until d <> 0 And e <> 0 And f <> 0 And g <> 0 And h <> 0
    Call SetText(g, TheName)
    Call SetText(h, ThePW)
    Call Click(e)
    Do
        DoEvents
        i = FindWindow("#32770", "America Online")
        j = AOLUser
    Loop Until i <> 0 Or j <> ""
    If j <> "" Then
        CheckPhish = True
    ElseIf i <> 0 Then
        k = FindWindowEx(i, 0, "Static", vbNullString)
        k = FindWindowEx(i, k, "Static", vbNullString)
        l = LCase(GetText(f))
        If InStr(l, "incorrect") <> 0 Then
              Call CloseWindow(i)
              m = FindWindowEx(d, e, "_AOL_Icon", vbNullString)
              Call Click(m)
              CheckPhish = False
        ElseIf InStr(l, "signed on") <> 0 Then
              Call CloseWindow(i)
              m = FindWindowEx(d, e, "_AOL_Icon", vbNullString)
              Call Click(m)
              CheckPhish = True
        End If
    End If
End Function
Public Sub ChatIgnoreByIndex(Index As Long)
    Dim a As Long, b As Long, c As String, d As Long
    Dim e As Long, f As Long
    a = ChatFindRoom
    b = FindWindowEx(a, 0, "_AOL_Listbox", vbNullString)
    Call SendMessage(b, LB_SETCURSEL, Index, 0)
    Call PostMessage(b, WM_LBUTTONDBLCLK, 0, 0)
    c = ChatGetListName(Index)
    Do
        DoEvents
        d = FindWindowEx(AOLMDI, 0, "AOL Child", c)
        e = FindWindowEx(d, 0, "_AOL_Checkbox", vbNullString)
    Loop Until d <> 0 And e <> 0
    Do
        f = SendMessage(e, BM_GETCHECK, 0, 0)
        Call Click(e)
    Loop Until f <> 0
    Call CloseWindow(d)
End Sub
Public Function ChatGetListName(Index As Long) As String
    On Error Resume Next
    Dim a As Long, b As Long, c As String, d As Long
    Dim e As Long, f As Long, g As Long, h As Long
    Dim i As Long, j As Long
    g = ChatFindRoom
    If g = 0 Then Exit Function
    h = FindWindowEx(g, 0, "_AOL_Listbox", vbNullString)
    i = GetWindowThreadProcessId(h, a)
    j = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, a)
    c = String(4, vbNullChar)
    b = SendMessage(h, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0) + 24
    Call ReadProcessMemory(j, b, c, 4, e)
    Call RtlMoveMemory(d, ByVal c, 4)
    d = d + 6
    c = String(16, vbNullChar)
    Call ReadProcessMemory(j, d, c, Len(c), e)
    f = InStr(c, vbNullChar)
    c = Left(c, f - 1)
    ChatGetListName = c
    Call CloseHandle(j)
End Function
Public Sub ChatIgnorebyString(ToIgnore As String)
    On Error Resume Next
    Dim a As Long, b As String
     For a = 0 To 25
        b = ChatGetListName(a)
        If TrimSpaces(LCase(b)) = TrimSpaces(LCase(ToIgnore)) Then
            Call ChatIgnoreByIndex(a)
            Exit Sub
        End If
    Next a
End Sub
Public Sub MailRemoveErrorNames(Lst As ListBox)
    'this will remove the names off a list that
    'cant recieve mail, their box is full, etc...
    'good for mmers
    Dim a As Long, b As Long, c As String, d As Long
    Dim e As Long, f As String, g As Long, h As String
    Dim i As Long
    a = FindWindowEx(AOLMDI, 0, "AOL Child", "Error")
    If a = 0 Then Exit Sub
    b = FindWindowEx(a, 0, "_AOL_View", vbNullString)
    c = GetText(b)
    d = GetLineCount(c)
    For e = 3 To d
        f = GetLine(c, e)
        g = InStr(1, f, " - ")
        If g = 0 Then GoTo NextE
        h = Left(f, g - 1)
        i = GetListIndex(Lst, h)
        Lst.RemoveItem i
NextE:
    Next e
End Sub
Public Function GetLine(ToGetFrom As String, LineNumber As Long) As String
    Dim a As Long, b As Long, c As String
    On Error Resume Next
    c = ToGetFrom
    If LineNumber = 1 Then
        b = InStr(c, Chr(13))
        c = Left(c, b - 1)
        GetLine = c
        Exit Function
    End If
    For a = 1 To LineNumber - 1
        b = InStr(c, Chr(13))
        c = Mid$(c, b + 2)
    Next a
    b = InStr(c, Chr(13))
    c = Left(c, b - 1)
    GetLine = c
End Function
Public Function GetListIndex(Lst As ListBox, ToFind As String) As Long
    Dim a As Long, b As Long
    If Lst.ListCount = 0 Then GetListIndex = -1
    b = -1
    For a = 0 To Lst.ListCount - 1
        If TrimSpaces(LCase(Lst.List(a))) = TrimSpaces(LCase(ToFind)) Then
            b = a
            GoTo Found
        End If
    Next a
Found:
    GetListIndex = b
    Exit Function

End Function
Public Sub KillModal()
    Dim a As Long
    a = FindWindow("_AOL_Modal", vbNullString)
    If a = 0 Then Exit Sub
    Do
        a = FindWindow("_AOL_Modal", vbNullString)
        CloseWindow (a)
        DoEvents
    Loop Until a = 0
End Sub
Public Function GetModalCount() As Long
    'counts amount of modal windows open
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As String
    a = FindWindow("_AOL_Modal", vbNullString)
    If a = 0 Then GetModalCount = 0: Exit Function
    d = 1
    For b = 1 To GetWindowCount
        c = GetWindow(a, GW_hWndNEXT)
        e = GetClass(c)
        If LCase(e) = "_aol_modal" Then
            d = d + 1
        End If
    Next b
    GetModalCount = d
End Function
Public Function GetWindowCount() As Long
    'returns amount of windows are open
    'this doesnt count their children
    Dim a As Long, b As Long
    b = 0
    a = FindWindow(vbNullString, vbNullString)
    If a = 0 Then GetWindowCount = 0: Exit Function
    b = 1
    Do
        a = GetWindow(a, GW_hWndNEXT)
        If a <> 0 Then b = b + 1
    Loop Until a = 0
    GetWindowCount = b
End Function
Public Function GetAOLChildCount() As Long
    'returns amount of children are open in aol
    Dim a As Long, b As Long
    b = 0
    a = FindWindowEx(AOLMDI, 0, "AOL Child", vbNullString)
    If a = 0 Then GetAOLChildCount = 0: Exit Function
    b = 1
    Do
        a = GetWindow(a, GW_hWndNEXT)
        If a <> 0 Then b = b + 1
    Loop Until a = 0
    GetAOLChildCount = b
End Function
Public Sub UpChatON()
    Dim a As Long, b As Long, c As Long, d As String
    Dim e As String, f As Long, g As Long
    g = 0
    a = FindWindow("_AOL_Modal", vbNullString)
    If a = 0 Then Exit Sub
    For b = 1 To GetWindowCount
        c = GetWindow(a, GW_hWndNEXT)
        e = GetClass(c)
        If LCase(e) = "_aol_modal" Then
            d = GetCaption(c)
            f = InStr(1, d, " - ")
            If f = 0 Then GoTo NextB
            d = Left(d, f - 1)
            If LCase(TrimSpaces(d)) = "filetransfer" Then
                g = c
                GoTo FoundModal
            End If
        End If
NextB:
    Next b
    If g = 0 Then Exit Sub
FoundModal:
    Call EnableWindow(g, 0)
    Call EnableWindow(AOLWindow, 1)
End Sub
Public Sub UpChatOFF()
    Dim a As Long, b As Long, c As Long, d As String
    Dim e As String, f As Long, g As Long
    g = 0
    a = FindWindow("_AOL_Modal", vbNullString)
    If a = 0 Then Exit Sub
    For b = 1 To GetWindowCount
        c = GetWindow(a, GW_hWndNEXT)
        e = GetClass(c)
        If LCase(e) = "_aol_modal" Then
            d = GetCaption(c)
            f = InStr(1, d, " - ")
            If f = 0 Then GoTo NextB
            d = Left(d, f - 1)
            If LCase(TrimSpaces(d)) = "filetransfer" Then
                g = c
                GoTo FoundModal
            End If
        End If
NextB:
    Next b
    If g = 0 Then Exit Sub
FoundModal:
    Call EnableWindow(g, 1)
    Call EnableWindow(AOLWindow, 0)
End Sub
Public Sub MailOpenNew()
    Dim a As Long
    a = GetToolBarIcon(1)
    Call Click(a)
    Do
        DoEvents
        a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    Loop Until a <> 0
End Sub
Public Sub MailOpenOld()
    Call RunPopUpMenu(3, "O")
End Sub
Public Sub MailOpenSent()
    Call RunPopUpMenu(3, "S")
End Sub
Public Sub MailOpenFlash()
    Dim a As Long, b As Long, c As Long
    a = FindWindow("#32768", vbNullString)
    b = GetToolBarIcon(3)
    Call PostMessage(b, WM_LBUTTONDOWN, 0, 0)
    Call PostMessage(b, WM_LBUTTONUP, 0, 0)
    Call SetCursorPos(0, 0)
    Do
        DoEvents
        c = FindWindow("#32768", vbNullString)
    Loop Until c <> a
    Call SetCursorPos(0, 0)
    Call PostMessage(c, WM_CHAR, Asc("d"), 0)
    Do
        DoEvents
        a = FindWindow("#32768", vbNullString)
    Loop Until a <> c
    Call SetCursorPos(0, 0)
    Call PostMessage(c, WM_CHAR, Asc("I"), 0)
    Do
        DoEvents
        a = FindWindowEx(AOLMDI, 0, "AOL Child", "Incoming/Saved Mail")
    Loop Until a <> 0
End Sub

Public Sub RunPopUpMenu2(IconNum As Long, MenuPosition As Long)
    'this was written by keg ;]
    'how to use:
    '   iconnum = the number of the aol icon on the toolbar
    '   menuposition = the position of the menuitem you
    '                          want to run
    'for example, you can do:
    '   call runpopupmenu2(3,3)
    'to open a write mail window
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindWindow("#32768", vbNullString)
    b = GetToolBarIcon(IconNum)
    Call PostMessage(b, WM_LBUTTONDOWN, 0, 0)
    Call PostMessage(b, WM_LBUTTONUP, 0, 0)
    Call SetCursorPos(0, 0)
    Do
        DoEvents
        c = FindWindow("#32768", vbNullString)
    Loop Until c <> a
    Call SetCursorPos(0, 0)
    If MenuPosition = 1 Then
        Call PostMessage(c, WM_KEYDOWN, VK_RETURN, 0)
        Call PostMessage(c, WM_KEYUP, VK_RETURN, 0)
    Else
        For d = 1 To MenuPosition
            Call PostMessage(c, WM_KEYDOWN, VK_DOWN, 0)
            Call PostMessage(c, WM_KEYUP, VK_DOWN, 0)
        Next d
        Call PostMessage(c, WM_KEYDOWN, VK_RETURN, 0)
        Call PostMessage(c, WM_KEYUP, VK_RETURN, 0)
    End If
End Sub
Public Sub MailListNew(Lst As ListBox, AddNumbers As Boolean)
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_TabControl", vbNullString)
    c = FindWindowEx(b, 0, "_AOL_TabPage", vbNullString)
    d = FindWindowEx(c, 0, "_AOL_Tree", vbNullString)
    Call APIListToVBList(d, Lst, True, AddNumbers)
End Sub
Public Sub MailListOld(Lst As ListBox, AddNumbers As Boolean)
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_TabControl", vbNullString)
    c = FindWindowEx(b, 0, "_AOL_TabPage", vbNullString)
    c = FindWindowEx(b, c, "_AOL_TabPage", vbNullString)
    d = FindWindowEx(c, 0, "_AOL_Tree", vbNullString)
    Call APIListToVBList(d, Lst, True, AddNumbers)
End Sub
Public Sub MailListSent(Lst As ListBox, AddNumbers As Boolean)
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_TabControl", vbNullString)
    c = FindWindowEx(b, 0, "_AOL_TabPage", vbNullString)
    c = FindWindowEx(b, c, "_AOL_TabPage", vbNullString)
    c = FindWindowEx(b, c, "_AOL_TabPage", vbNullString)
    d = FindWindowEx(c, 0, "_AOL_Tree", vbNullString)
    Call APIListToVBList(d, Lst, True, AddNumbers)
End Sub
Public Sub MailListFlash(Lst As ListBox, AddNumbers As Boolean)
    Dim a As Long, b As Long
    a = FindWindowEx(AOLMDI, 0, "AOL Child", "Incoming/Saved Mail")
    b = FindWindowEx(a, 0, "_AOL_Tree", vbNullString)
    Call APIListToVBList(b, Lst, True, AddNumbers)
End Sub
Public Sub MailSelectNew(Index As Long)
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_TabControl", vbNullString)
    c = FindWindowEx(b, 0, "_AOL_TabPage", vbNullString)
    d = FindWindowEx(c, 0, "_AOL_Tree", vbNullString)
    Call SendMessage(d, LB_SETCURSEL, Index, 0)
End Sub
Public Sub MailSelectOld(Index As Long)
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_TabControl", vbNullString)
    c = FindWindowEx(b, 0, "_AOL_TabPage", vbNullString)
    c = FindWindowEx(b, c, "_AOL_TabPage", vbNullString)
    d = FindWindowEx(c, 0, "_AOL_Tree", vbNullString)
    Call SendMessage(d, LB_SETCURSEL, Index, 0)
End Sub
Public Sub MailSelectSent(Index As Long)
    Dim a As Long, b As Long, c As Long, d As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_TabControl", vbNullString)
    c = FindWindowEx(b, 0, "_AOL_TabPage", vbNullString)
    c = FindWindowEx(b, c, "_AOL_TabPage", vbNullString)
    c = FindWindowEx(b, c, "_AOL_TabPage", vbNullString)
    d = FindWindowEx(c, 0, "_AOL_Tree", vbNullString)
    Call SendMessage(d, LB_SETCURSEL, Index, 0)
End Sub
Public Sub MailSelectFlash(Index As Long)
    Dim a As Long, b As Long
    a = FindWindowEx(AOLMDI, 0, "AOL Child", "Incoming/Saved Mail")
    b = FindWindowEx(a, 0, "_AOL_Tree", vbNullString)
    Call SendMessage(b, LB_SETCURSEL, Index, 0)
End Sub
Public Sub MailClickOpen()
    Dim a As Long, b As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    Call Click(b)
End Sub

Public Sub MailClickKeepAsNew()
    Dim a As Long, b As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    Call Click(b)
End Sub
Public Sub MailClickDelete()
    Dim a As Long, b As Long
    a = FindChildByTitle(AOLMDI, AOLUser + "'s Online Mailbox")
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    Call Click(b)
End Sub
Public Sub MailClickOpenFlash()
    Dim a As Long, b As Long
    a = FindWindowEx(AOLMDI, 0, "AOL Child", "Incoming/Saved Mail")
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    Call Click(b)
End Sub
Public Sub MailFwdNewbyIndex(Index As Long, ToWho As String, ToSay As String)
    Dim a As Long, b As Long, c As Long, d As Long
    Do
        DoEvents
        Call MailSelectNew(Index)
        Call MailClickOpen
        a = MailFindOpen
    Loop Until a <> 0
    Do
        DoEvents
        Call MailClickFwd(a)
        b = FindChildByTitle(AOLMDI, "Fwd: ")
    Loop Until b <> 0
    DoEvents
    Call MailRemoveFwd
    c = FindWindowEx(b, 0, "_AOL_Edit", vbNullString)
    Call SetText(c, ToWho)
    c = FindWindowEx(b, 0, "RICHCNTL", vbNullString)
    Call SetText(c, ToSay)
    c = 0
    For d = 1 To 12
        c = FindWindowEx(b, c, "_AOL_Icon", vbNullString)
    Next d
    Do
        DoEvents
        Call Click(c)
        Pause 0.3
        b = FindChildByTitle(AOLMDI, "Fwd: ")
        c = FindWindowEx(AOLMDI, 0, "AOL Child", "Error")
    Loop Until b = 0 Or c <> 0
    Call CloseWindow(a)
End Sub
Public Sub MailFwdFlashbyIndex(Index As Long, ToWho As String, ToSay As String)
    Dim a As Long, b As Long, c As Long, d As Long
    Do
        DoEvents
        a = MailFindOpen
        CloseWindow (a)
    Loop Until a = 0
    Do
        DoEvents
        Call MailSelectFlash(Index)
        Call MailClickOpenFlash
        a = MailFindOpen
    Loop Until a <> 0
    Do
        DoEvents
        Call MailClickFwd(a)
        b = FindChildByTitle(AOLMDI, "Fwd: ")
    Loop Until b <> 0
    DoEvents
    Call MailRemoveFwd
    c = FindWindowEx(b, 0, "_AOL_Edit", vbNullString)
    Call SetText(c, ToWho)
    c = FindWindowEx(b, 0, "RICHCNTL", vbNullString)
    Call SetText(c, ToSay)
    c = 0
    For d = 1 To 12
        c = FindWindowEx(b, c, "_AOL_Icon", vbNullString)
    Next d
    Do
        DoEvents
        Call Click(c)
        Pause 0.3
        b = FindChildByTitle(AOLMDI, "Fwd: ")
        c = FindWindowEx(AOLMDI, 0, "AOL Child", "Error")
    Loop Until b = 0 Or c <> 0
    Call CloseWindow(a)
End Sub
Public Function MailFindOpen() As Long
    Dim a As Long, b As Long
    Dim c As Long, d As Long
    Dim e As Long, f As Long
    Dim g As Long, h As Long
    Dim i As String, j As Long
    Dim k As Long
    h = FindWindowEx(AOLMDI, b, "AOL Child", vbNullString)
    For a = 0 To GetAOLChildCount + 10
        b = h
        d = FindWindowEx(b, 0, "RICHCNTL", vbNullString)
        e = FindWindowEx(b, 0, "_AOL_Static", vbNullString)
        f = FindWindowEx(b, 0, "_AOL_Icon", vbNullString)
        c = FindWindowEx(b, f, "_AOL_Icon", vbNullString)
        i = GetText(d)
        j = InStr(1, i, "From")
        k = InStr(1, i, "Subj")
        h = GetWindow(b, GW_hWndNEXT)
        If b <> 0 And c <> 0 And d <> 0 And e <> 0 And f <> 0 And j <> 0 And k <> 0 And i <> "" Then GoTo FoundIt
    Next a
FoundIt:
    MailFindOpen = b
End Function
Public Sub MailClickFwd(Mail As Long)
    Dim a As Long, b As Long, c As Long
    b = FindWindowEx(Mail, 0, "_AOL_Icon", vbNullString)
    For c = 1 To 6
        b = FindWindowEx(Mail, b, "_AOL_Icon", vbNullString)
    Next c
    Call Click(b)
End Sub
Public Function MsgBoxOnTop(Prompt As String, Optional Style As VbMsgBoxStyle, Optional Title As String) As VbMsgBoxResult
    Dim a As VbMsgBoxResult
    If Style = 0 Then
        Style = vbOKOnly
    End If
    If Title = vbNullString Then
        Title = App.Title
    End If
    a = MsgBox(Prompt, Style + vbSystemModal, Title)
    MsgBoxOnTop = a
End Function
Public Function ChatSendBox() As Long
    Dim a As Long, b As Long
    a = ChatFindRoom
    b = FindWindowEx(a, 0, "RICHCNTL", vbNullString)
    b = FindWindowEx(a, b, "RICHCNTL", vbNullString)
    ChatSendBox = b
End Function
Public Function ChatSendButton() As Long
    Dim a As Long, b As Long
    a = ChatFindRoom
    b = FindWindowEx(a, 0, "RICHCNTL", vbNullString)
    b = FindWindowEx(a, b, "RICHCNTL", vbNullString)
    b = GetWindow(b, GW_hWndNEXT)
    ChatSendButton = b
End Function
Public Function ChatText() As String
    Dim a As Long, b As Long
    a = ChatFindRoom
    b = FindWindowEx(a, 0, "RICHCNTL", vbNullString)
    ChatText = GetText(b)
End Function
Public Function ChatTextBox() As String
    Dim a As Long, b As Long
    a = ChatFindRoom
    b = FindWindowEx(a, 0, "RICHCNTL", vbNullString)
    ChatTextBox = b
End Function
Public Sub ChatClear()
    Call SetText(ChatTextBox, "")
End Sub
Public Sub ChatLink(URL As String, LinkText As String)
    Call ChatSend("<a href=" + Chr(34) + Chr(34) + "><a href=" + Chr(34) + Chr(34) + "><a href=" + Chr(34) + URL + Chr(34) + ">" + LinkText + "<font color=" + Chr(34) + "#FEFEFE" + Chr(34) + "></a>")
End Sub
Public Sub ChatRoom(RoomName As String)
    Call Keyword("aol://2719:2-2-" & RoomName)
End Sub
Public Sub ChatScrollMacro(ToScroll As String)
    'to make it quit before the end of scrolling the lines,
    'just do StopScrolling = True
    Dim a As Long, b As Long, c As String
    StopScrolling = False
    a = GetLineCount(ToScroll)
    For b = 1 To a
        If StopScrolling = True Then Exit Sub
        c = GetLine(ToScroll, b)
        Call ChatSend(c)
        Pause 0.75
    Next b
End Sub
Public Sub ChatScroll(ToScroll As String, AmountOfTimes As Long, ToPauseFor As Long)
    Dim a As Long
    'to make it quit before the end of scrolling the lines,
    'just do StopScrolling = True
    StopScrolling = False
    For a = 1 To AmountOfTimes
        If StopScrolling = True Then Exit Sub
        Call ChatSend(ToScroll)
        Call Pause(ToPauseFor)
    Next a
End Sub
Public Sub ChatScroll2(ToScroll As String, ToPauseFor As Long)
    Dim a As Long
    'this one you HAVE to do StopScrolling = true to stop
    StopScrolling = False
    Do
        If StopScrolling = True Then Exit Sub
        Call ChatSend(ToScroll)
        Call Pause(ToPauseFor)
    Loop
End Sub
Public Function RandomNumber(Range As Long) As Long
    Dim a As Long
    Randomize
    a = Fix((Range * Rnd) + 1)
    RandomNumber = a
End Function
Public Sub ClipCopy(ToCopy As String)
    Clipboard.Clear
    Clipboard.SetText (ToCopy)
End Sub
Public Function ClipPaste() As String
    ClipPaste = Clipboard.GetText
End Function
Public Sub ChatLastLineSNandMSG(SnTextBox As TextBox, MSGTextBox As TextBox)
    Dim a As String, b As Long, c As String, d As String
    a = ChatLastLine
    b = InStr(1, a, ":")
    If b = 0 Then
        SnTextBox.Text = ""
        MSGTextBox.Text = ""
        Exit Sub
    End If
    c = Left(a, b - 1)
    d = Mid(a, b + 3)
    SnTextBox.Text = c
    MSGTextBox.Text = d
End Sub
Public Sub FormOutline(TheForm As Form)
    TheForm.AutoRedraw = True
    TheForm.Line (0, 0)-(TheForm.Width, 0), RGB(0, 0, 0)
    TheForm.Line (0, 0)-(0, TheForm.Height), RGB(0, 0, 0)
    TheForm.Line (TheForm.Width - 10, 0)-(TheForm.Width - 10, TheForm.Height), RGB(0, 0, 0)
    TheForm.Line (0, TheForm.Height - 10)-(TheForm.Width, TheForm.Height - 10), RGB(0, 0, 0)
End Sub
Public Function IsUserOnline() As Boolean
    Dim a As String
    a = AOLUser
    If a = "" Then
        IsUserOnline = False
    Else
        IsUserOnline = True
    End If
End Function
Public Sub GhostOn()
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, h As Long
    Dim i As Long
    a = FindWindowEx(AOLMDI, 0, "AOL Child", "Buddy List Window")
    If a = 0 Then
        Call OpenBL
        Do
            DoEvents
            a = FindWindowEx(AOLMDI, 0, "AOL Child", "Buddy List Window")
            b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
            b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
            b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
        Loop Until a <> 0 And b <> 0
    End If
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    Do
        DoEvents
        Call Click(b)
        Call Pause(1)
        c = FindChildByTitle(AOLMDI, AOLUser + "'s Buddy List")
        d = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
    Loop Until c <> 0 And d <> 0
    Do
        DoEvents
        Call Click(d)
        Call Pause(1)
        e = FindWindowEx(AOLMDI, 0, "AOL Child", "Privacy Preferences")
        f = FindWindowEx(e, 0, "_AOL_Checkbox", vbNullString)
        f = FindWindowEx(e, f, "_AOL_Checkbox", vbNullString)
        f = FindWindowEx(e, f, "_AOL_Checkbox", vbNullString)
        f = FindWindowEx(e, f, "_AOL_Checkbox", vbNullString)
        f = FindWindowEx(e, f, "_AOL_Checkbox", vbNullString)
        g = FindWindowEx(e, f, "_AOL_Checkbox", vbNullString)
        g = FindWindowEx(e, g, "_AOL_Checkbox", vbNullString)
    Loop Until e <> 0 And g <> 0 And f <> 0
    For h = 1 To 10
        Call Click(f)
        Call Click(g)
        Call Pause(0.1)
    Next h
    i = FindWindowEx(e, 0, "_AOL_Icon", vbNullString)
    i = FindWindowEx(e, i, "_AOL_Icon", vbNullString)
    i = FindWindowEx(e, i, "_AOL_Icon", vbNullString)
    i = FindWindowEx(e, i, "_AOL_Icon", vbNullString)
    Do
        Call Click(f)
        Call Click(g)
        Call Click(i)
        e = FindWindowEx(AOLMDI, 0, "AOL Child", "Privacy Preferences")
        Call Pause(1)
    Loop Until e = 0
    Do
        DoEvents
        e = FindWindow("#32770", "America Online")
    Loop Until e <> 0
    Do
        DoEvents
        e = FindWindow("#32770", "America Online")
        Call CloseWindow(e)
    Loop Until e = 0
    Call CloseWindow(c)
End Sub
Public Sub GhostOff()
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, h As Long
    Dim i As Long
    a = FindWindowEx(AOLMDI, 0, "AOL Child", "Buddy List Window")
    If a = 0 Then
        Call OpenBL
        Do
            DoEvents
            a = FindWindowEx(AOLMDI, 0, "AOL Child", "Buddy List Window")
            b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
            b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
            b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
        Loop Until a <> 0 And b <> 0
    End If
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    b = FindWindowEx(a, b, "_AOL_Icon", vbNullString)
    Do
        DoEvents
        Call Click(b)
        Call Pause(1)
        c = FindChildByTitle(AOLMDI, AOLUser + "'s Buddy List")
        d = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
        d = FindWindowEx(c, d, "_AOL_Icon", vbNullString)
    Loop Until c <> 0 And d <> 0
    Do
        DoEvents
        Call Click(d)
        Call Pause(1)
        e = FindWindowEx(AOLMDI, 0, "AOL Child", "Privacy Preferences")
        f = FindWindowEx(e, 0, "_AOL_Checkbox", vbNullString)
        g = FindWindowEx(e, f, "_AOL_Checkbox", vbNullString)
        g = FindWindowEx(e, g, "_AOL_Checkbox", vbNullString)
        g = FindWindowEx(e, g, "_AOL_Checkbox", vbNullString)
        g = FindWindowEx(e, g, "_AOL_Checkbox", vbNullString)
        g = FindWindowEx(e, g, "_AOL_Checkbox", vbNullString)
    Loop Until e <> 0 And g <> 0 And f <> 0
    For h = 1 To 10
        Call Click(f)
        Call Click(g)
        Call Pause(0.1)
    Next h
    i = FindWindowEx(e, 0, "_AOL_Icon", vbNullString)
    i = FindWindowEx(e, i, "_AOL_Icon", vbNullString)
    i = FindWindowEx(e, i, "_AOL_Icon", vbNullString)
    i = FindWindowEx(e, i, "_AOL_Icon", vbNullString)
    Do
        Call Click(f)
        Call Click(g)
        Call Click(i)
        e = FindWindowEx(AOLMDI, 0, "AOL Child", "Privacy Preferences")
        Call Pause(1)
    Loop Until e = 0
    Do
        DoEvents
        e = FindWindow("#32770", "America Online")
    Loop Until e <> 0
    Do
        DoEvents
        e = FindWindow("#32770", "America Online")
        Call CloseWindow(e)
    Loop Until e = 0
    Call CloseWindow(c)
End Sub
Public Function CheckIfOnline(TheSN As String) As Boolean
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As String
    If IsUserOnline = False Then Exit Function
    Call Keyword("aol://9293:")
    Do
        a = FindWindowEx(AOLMDI, 0, "AOL Child", "Send Instant Message")
        b = FindWindowEx(a, 0, "_AOL_Edit", vbNullString)
        DoEvents
    Loop Until a <> 0 And b <> 0
    Call SetText(b, TheSN)
    d = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    For e = 1 To 9
        d = FindWindowEx(a, d, "_AOL_Icon", vbNullString)
    Next e
    Call Click(d)
    Do
        DoEvents
        e = FindWindow("#32770", "America Online")
    Loop Until e <> 0
    c = FindWindow("#32770", "America Online")
    f = FindWindowEx(c, 0, "Static", vbNullString)
    f = FindWindowEx(c, f, "Static", vbNullString)
    g = GetText(f)
    If InStr(g, "able to receive") <> 0 Then
        CheckIfOnline = True
    Else
        CheckIfOnline = False
    End If
    Call CloseWindow(c)
    Call CloseWindow(a)
End Function
Public Function CheckIfAlive(TheSN As String) As Boolean
    If IsUserOnline = False Then
        CheckIfAlive = False
        Exit Function
    End If
    Dim a As Long, b As Long, c As Long, d As Long
    Dim e As Long, f As Long, g As Long, h As String
    Dim i As Long, j As String
    a = GetToolBarIcon(2)
    Call Click(a)
    Do
        b = FindWindowEx(AOLMDI, 0, "AOL Child", "Write Mail")
        c = FindWindowEx(b, 0, "_AOL_Edit", vbNullString)
        d = FindWindowEx(b, 0, "_AOL_Icon", vbNullString)
    Loop Until b <> 0 And c <> 0 And d <> 0
    Call SetText(c, "*, " + TheSN)
    d = 0
    For e = 1 To 14
        d = FindWindowEx(b, d, "_AOL_Icon", vbNullString)
    Next e
    Call Click(d)
    Do
        DoEvents
        f = FindWindowEx(AOLMDI, 0, "AOL Child", "Error")
        g = FindWindowEx(f, 0, "_AOL_View", vbNullString)
        h = GetText(g)
    Loop Until f <> 0 And g <> 0 And h <> ""
    Call MsgBoxOnTop(h)
    i = InStr(1, TrimSpaces(LCase(h)), TrimSpaces(LCase(TheSN)))
    If i = 0 Then
        CheckIfAlive = True
    Else
        CheckIfAlive = False
    End If
    Call CloseWindow(f)
    Call PostMessage(b, WM_CLOSE, 0, 0)
    Do
        b = FindWindow("#32770", "America Online")
        c = FindWindowEx(b, 0, "Button", "&No")
        DoEvents
    Loop Until b <> 0
    Call PostMessage(c, WM_KEYDOWN, VK_SPACE, 0)
    Call PostMessage(c, WM_KEYUP, VK_SPACE, 0)
End Function
Public Sub OpenBL()
    'opens buddy list
    Call Keyword("bv")
End Sub

Public Sub xHook()
    PrevWndProc = SetWindowLong(ghWnd, GWL_WNDPROC, AddressOf xNewMsg)
End Sub
Public Sub xUnhook()
    Dim temp As Long
    temp = SetWindowLong(ghWnd, GWL_WNDPROC, PrevWndProc)
End Sub
Public Function xNewMsg(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'this is like the event on a subclassing ocx when
    'a new msg is sent from the subclassed window
    
    'dont edit the following line
    xNewMsg = CallWindowProc(PrevWndProc, hwnd, msg, wParam, lParam)
    'you can edit the following lines to do different things
    'when different messages are sent
    If msg = WM_LBUTTONDOWN Then
        Call MsgBoxOnTop("Left Button Clicked", vbOKOnly + vbExclamation, "Subclassed ;]")
    End If
End Function
Public Sub xREAD_ME_FIRST()
'these "x" functions are subclassing functions that can only
'subclass your own forms
'vb isnt powerful enough for subclassing other windows other
'then its own (without DLL's and OCX's)

'1 FORM MAY BE SUBCLASSED AT A TIME....
'U HAVE TO PUT "ghWnd = me.hwnd" IN THE FORM'S
'LOAD PROCEDURE!
'SUBCLASSING A FORM TWICE WILL RESULT IN HAVING
'TO CLOSE VB
'YOU HAVE TO UNSUBCLASS BEFORE YOU TAKE YOU
'END YOUR PROGRAM OR VB WILL CRASH (SAVE OFTEN!)
End Sub
Public Sub RunProgram(ThePath As String, Optional Style As VbAppWinStyle)
    If Style = 0 Then
        Style = vbNormalFocus
    End If
    Call Shell(ThePath, Style)
End Sub
Function TrimTimer()
    Dim a As Long, b As String, c As Long
    a = Timer
    b = Str(a)
    c = InStr(b, ".")
    If c = 0 Then GoTo NoPeriod
    b = Left$(b, c + 1)
NoPeriod:
    TrimTimer = b
End Function
Public Function TrimDate() As String
    Dim a As String, b As Long, c As String
    Dim d As String, e As Long, f As String
    'this function is y2k compliant =D
    a = Date
    b = InStr(a, "/")
    c = Left(a, b - 1)
    If c = "1" Then
        d = "January"
    ElseIf c = "2" Then
        d = "February"
    ElseIf c = "3" Then
        d = "March"
    ElseIf c = "4" Then
        d = "April"
    ElseIf c = "5" Then
        d = "May"
    ElseIf c = "6" Then
        d = "June"
    ElseIf c = "7" Then
        d = "July"
    ElseIf c = "8" Then
        d = "August"
    ElseIf c = "9" Then
        d = "September"
    ElseIf c = "10" Then
        d = "October"
    ElseIf c = "11" Then
        d = "November"
    ElseIf c = "12" Then
        d = "December"
    End If
    e = InStr(b + 1, a, "/")
    c = Mid(a, b + 1, e - b - 1)
    d = d + " " + c + ","
    c = Mid(a, e + 1)
    If Len(c) = 2 Then
        c = "19" + c
    End If
    d = d + " " + c
    TrimDate = d
End Function
Public Function TrimTime() As String
    Dim a As String, b As String, c As Long
    Dim d As String
    a = Time
    b = Right(a, 2)
    c = InStr(a, ":")
    c = InStr(c + 1, a, ":")
    d = Left(a, c - 1)
    d = d + " " + b
    TrimTime = d
End Function
Function FileOpen(ThePath As String) As String
    Dim a As Long, b As String
    a = FreeFile
    Open ThePath For Binary As #a
    b = String$(LOF(a), " ")
    Get #a, , b
    FileOpen = b
    Close #a
End Function
Function FileSearch(ThePath As String, ToFind As String) As Boolean
    Dim a As Long, b As Long
    a = FreeFile
    Open ThePath For Binary As #a
    b = InStr(1, a, ToFind, vbBinaryCompare)
    Close #a
    If b = 0 Then
        FileSearch = False
    Else
        FileSearch = True
    End If
End Function
Sub FileSave(ThePath As String, ToSave As String)
    Dim a As Long
    On Error GoTo ErrorStop
    If FileExist(ThePath) = True Then
        Call Kill(ThePath)
    End If
    a = FreeFile
    Open ThePath For Binary As #a
    Put #a, 1, ToSave
ErrorStop:
    Close #a
End Sub
Public Function ChatCount() As Long
    Dim a As Long, b As Long
    a = FindWindowEx(ChatFindRoom, 0, "_AOL_Listbox", vbNullString)
    b = SendMessage(a, LB_GETCOUNT, 0, 0)
    ChatCount = b
End Function
Public Sub FileSetAttribute(ThePath As String, TheAttribute As VbFileAttribute)
    If FileExist(ThePath) = False Then Exit Sub
    Call SetAttr(ThePath, TheAttribute)
End Sub
Public Function TextHacker(ToHackerize As String) As String
    Dim a As String
    a = UCase(ToHackerize)
    a = ReplaceChar(ToHackerize, "A", "a")
    a = ReplaceChar(ToHackerize, "E", "e")
    a = ReplaceChar(ToHackerize, "I", "i")
    a = ReplaceChar(ToHackerize, "O", "o")
    a = ReplaceChar(ToHackerize, "U", "u")
    TextHacker = a
End Function
Public Sub TextPrint(ToPrint As String)
    On Error GoTo Oops
    Printer.Font = "Arial"
    Printer.FontSize = 10
    Printer.Print ToPrint
    Printer.EndDoc
    Exit Sub
Oops:
End Sub
Public Function TextSpace(ToSpace As String) As String
    Dim a As Long, b As String, c As String
    c = ""
    For a = 1 To Len(ToSpace)
        b = Mid(ToSpace, a, 1)
        c = c + b + " "
    Next a
    c = Left(c, Len(c) - 1)
    TextSpace = c
End Function
Public Function TextVb32(ToVb32 As String) As String
    Dim a As Long, b As String, c As String
    Dim d As String, e As String
    If Len(ToVb32) = 0 Or TrimSpaces(ToVb32) = "" Then Exit Function
    b = " " + ToVb32
    For a = 1 To Len(b)
        c = Mid(b, a, 1)
        If a = Len(b) Then GoTo LastLetter
        If c = " " Then
            d = Mid(b, a + 1, 1)
            e = e + " <b>" + d + "</b>"
            a = a + 1
        Else
LastLetter:
            e = e + c
        End If
        DoEvents
    Next a
    e = Mid(e, 2)
    TextVb32 = e
End Function
Public Function TextElite(ToElite As String) As String
    Dim a As String
    a = LCase(ToElite)
    a = ReplaceChar(a, "a", "â")
    a = ReplaceChar(a, "b", "|›")
    a = ReplaceChar(a, "c", "¢")
    a = ReplaceChar(a, "d", "‹|")
    a = ReplaceChar(a, "e", "ë")
    a = ReplaceChar(a, "f", "ƒ")
    a = ReplaceChar(a, "h", "|-|")
    a = ReplaceChar(a, "i", "ï")
    a = ReplaceChar(a, "k", "|‹")
    a = ReplaceChar(a, "l", "£")
    a = ReplaceChar(a, "m", "^^")
    a = ReplaceChar(a, "n", "ñ")
    a = ReplaceChar(a, "o", "ø")
    a = ReplaceChar(a, "p", "þ")
    a = ReplaceChar(a, "s", "š")
    a = ReplaceChar(a, "t", "†")
    a = ReplaceChar(a, "u", "Ü")
    a = ReplaceChar(a, "v", "\/")
    a = ReplaceChar(a, "w", "vv")
    a = ReplaceChar(a, "x", "×")
    a = ReplaceChar(a, "y", "ÿ")
    TextElite = a
End Function
Public Sub Anti45MinTimer()
    Dim a As Long, b As Long
    a = FindWindow("_AOL_Palette", vbNullString)
    If a <> 0 Then
        Do
            DoEvents
            Pause 0.4
            b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
            Call Click(b)
        Loop Until b = 0
    End If
End Sub
Public Sub AntiIdle()
    Dim a As Long, b As Long, c As Long, d As String
    a = FindWindow("_AOL_Modal", vbNullString)
    b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
    c = FindWindowEx(a, 0, "_AOL_Static", vbNullString)
    d = Str(c)
    If InStr(1, LCase(d), "You have been idle") = 0 Then Exit Sub
    If a <> 0 Then
        Do
            DoEvents
            Pause 0.4
            b = FindWindowEx(a, 0, "_AOL_Icon", vbNullString)
            Call Click(b)
        Loop Until b = 0
    End If
End Sub

Public Function ShowOpen(Owner As Long, Title As String, Optional StartDir As String, Optional Filter As String) As String
    'use the following example of the filter variable
    'Filter = "Programs" & vbNullChar & "*.EXE" & vbNullChar & vbNullChar
    Dim udtFile As OPENFILENAME, lResult As Long, nNullPos As Integer
    Dim sFile As String, sFileTitle As String
    udtFile.lStructSize = Len(udtFile)
    udtFile.hwndOwner = Owner
    udtFile.FLAGS = OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST + OFN_HIDEREADONLY
    udtFile.lpstrFile = Space$(255)
    udtFile.nMaxFile = 255
    udtFile.lpstrFileTitle = Space$(255)
    udtFile.nMaxFileTitle = 255
    udtFile.lpstrInitialDir = StartDir
    udtFile.lpstrFilter = Filter
    udtFile.nFilterIndex = 1
    udtFile.lpstrTitle = Title
    lResult = GetOpenFileName(udtFile)
    If lResult <> 0 Then
        nNullPos = InStr(udtFile.lpstrFileTitle, vbNullChar)
        If nNullPos > 0 Then
            sFileTitle = Left$(udtFile.lpstrFileTitle, nNullPos - 1)
        End If
        nNullPos = InStr(udtFile.lpstrFile, vbNullChar)
        If nNullPos > 0 Then
            sFile = Left$(udtFile.lpstrFile, nNullPos - 1)
        End If
        ShowOpen = sFile
    End If
End Function

Public Function ShowSave(Owner As Long, Title As String, Optional StartDir As String, Optional Filter As String) As String
    'use the following example of the filter variable
    'Filter = "Programs" & vbNullChar & "*.EXE" & vbNullChar & vbNullChar
    Dim udtFile As OPENFILENAME, lResult As Long, nNullPos As Integer
    Dim sFile As String, sFileTitle As String
    udtFile.lStructSize = Len(udtFile)
    udtFile.hwndOwner = Owner
    udtFile.FLAGS = OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST + OFN_HIDEREADONLY
    udtFile.lpstrFile = Space$(255)
    udtFile.nMaxFile = 255
    udtFile.lpstrFileTitle = Space$(255)
    udtFile.nMaxFileTitle = 255
    udtFile.lpstrInitialDir = StartDir
    udtFile.lpstrFilter = Filter
    udtFile.nFilterIndex = 1
    udtFile.lpstrTitle = Title
    lResult = GetSaveFileName(udtFile)
    If lResult <> 0 Then
        nNullPos = InStr(udtFile.lpstrFileTitle, vbNullChar)
        If nNullPos > 0 Then
            sFileTitle = Left$(udtFile.lpstrFileTitle, nNullPos - 1)
        End If
        nNullPos = InStr(udtFile.lpstrFile, vbNullChar)
        If nNullPos > 0 Then
            sFile = Left$(udtFile.lpstrFile, nNullPos - 1)
        End If
        ShowSave = sFile
    End If
End Function

Public Sub FormGrow(TheForm As Form)
    Dim a As Long, b As Long, c As Long
    TheForm.Show
    a = TheForm.Width
    b = TheForm.Height
    TheForm.Width = 0
    TheForm.Height = 0
    For c = 0 To a
        DoEvents
        Pause 0.01
        TheForm.Width = c
    Next c
    For c = 0 To b
        DoEvents
        Pause 0.01
        TheForm.Height = c
    Next c
End Sub
Public Sub FormFadeBG(TheForm As Form, Red1 As Long, Red2 As Long, Green1 As Long, Green2 As Long, Blue1 As Long, Blue2 As Long)
    Dim frmWidth As Long, frmHeight As Long, Fade As Long
    Dim NextRed As Long, NextGreen As Long, NextBlue As Long
    Dim NextRGB As Long
    TheForm.AutoRedraw = True
    frmWidth = TheForm.Width
    frmHeight = TheForm.Height
    For Fade = 0 To frmHeight Step 5
        DoEvents
        NextRed = (Red2 - Red1) / frmHeight * Fade + Red1
        NextGreen = (Green2 - Green1) / frmHeight * Fade + Green1
        NextBlue = (Blue2 - Blue1) / frmHeight * Fade + Blue1
        NextRGB = RGB(NextRed, NextGreen, NextBlue)
        TheForm.Line (0, Fade)-(frmWidth, Fade), NextRGB
    Next Fade
End Sub
Public Sub FadePicBox(PicBox As PictureBox, Red1 As Long, Red2 As Long, Green1 As Long, Green2 As Long, Blue1 As Long, Blue2 As Long)
    Dim PBWidth As Long, PBHeight As Long, Fade As Long
    Dim NextRed As Long, NextGreen As Long, NextBlue As Long
    Dim NextRGB As Long
    PBWidth = PicBox.Width
    PBHeight = PicBox.Height
    For Fade = 0 To PBWidth Step 1
        NextRed = (Red2 - Red1) / PBWidth * Fade + Red1
        NextGreen = (Green2 - Green1) / PBWidth * Fade + Green1
        NextBlue = (Blue2 - Blue1) / PBWidth * Fade + Blue1
        NextRGB = RGB(NextRed, NextGreen, NextBlue)
        PicBox.Line (Fade, 0)-(Fade, PBHeight), NextRGB
    Next Fade
End Sub
Public Function Fade2(TheText As String, Red1 As Long, Red2 As Long, Green1 As Long, Green2 As Long, Blue1 As Long, Blue2 As Long, wavy As Boolean)
    Dim txtLen As Long, Fade As Long, nString As String
    Dim NextRed As Long, NextGreen As Long, NextBlue As Long
    Dim NextRGB As Long, NextHex As String, nChar As String
    txtLen = Len(TheText)
    nString = ""
    nChar = ""
    For Fade = 1 To txtLen
        DoEvents
        nChar = Mid(TheText, Fade, 1)
        NextRed = (Red2 - Red1) / txtLen * Fade + Red1
        NextGreen = (Green2 - Green1) / txtLen * Fade + Green1
        NextBlue = (Blue2 - Blue1) / txtLen * Fade + Blue1
        NextRGB = RGB(NextRed, NextGreen, NextBlue)
        NextHex = FadeRGBtoHEX(NextRGB)
        nString = nString + "<Font Color=" + Chr(34) + NextHex + Chr(34) + ">" + nChar
    Next Fade
    Fade2 = nString
End Function
Public Function FadeRGBtoHEX(RGBColor As Long) As String
    Dim a As String, b As Long, c As String
    a = Hex(RGBColor)
    b = 6 - Len(a)
    c = String(b, "0")
    a = a + c
    FadeRGBtoHEX = a
End Function

Public Sub ListToList(StartList As ListBox, DestList As ListBox)
    Dim a As Long, b As String
    For a = 0 To StartList.ListCount - 1
        b = StartList.List(a)
        DestList.AddItem b
    Next a
End Sub

Public Sub FormUnloadRight(TheForm As Form)
    Dim a As Long, b As Long
    a = TheForm.Left
    For b = a To Screen.Width Step 100
        DoEvents
        Pause 0.0001
        TheForm.Left = b
    Next b
    Unload TheForm
End Sub
Public Sub FormUnloadLeft(TheForm As Form)
    Dim a As Long, b As Long
    a = TheForm.Left
    For b = a To (0 - a) Step -100
        DoEvents
        Pause 0.0001
        TheForm.Left = b
    Next b
    Unload TheForm
End Sub
Public Sub FormExplode(TheForm As Form)
    Dim a As Long, b As Long, c As Long
    Dim hDone As Boolean, wDone As Boolean
    TheForm.Show
    a = TheForm.Width
    b = TheForm.Height
    hDone = False
    hDone = False
    Let TheForm.Width = 0
    Let TheForm.Height = 0
    Do
        DoEvents
        Call FormCenter(TheForm)
        TheForm.Height = TheForm.Height + 100
        TheForm.Width = TheForm.Width + 100
        If TheForm.Width > a Then
            TheForm.Width = a
            wDone = True
        End If
        If TheForm.Height > b Then
            TheForm.Height = b
            hDone = True
        End If
    Loop Until hDone = True And wDone = True
End Sub

Public Sub RunWebPage(YourForm As Form, URL As String)
    Call ShellExecute(YourForm.hwnd, "Open", URL, "", "", SW_NORMAL)
End Sub
Public Function Enter() As String
    'this adds the enter key
    'to go to the next line
    Enter = Chr(13) + Chr(10)
End Function
Public Function Qu() As String
    'this is the " character
    Qu = Chr(34)
End Function
Public Function Decrypt(ToChange As String, Key As Long) As String
    'this can be used for an e-chat
    On Error Resume Next
    Dim a As Long, b As Long, c As String, d As Long
    Dim e As String, f As String
    a = Len(ToChange)
    For b = 1 To a Step 2
        c = Mid(ToChange, b, 1)
        d = Asc(c)
        d = d - Key
        If d < 0 Then
            d = 255 + d
        End If
        e = Chr(d)
        f = f + e
    Next b
    Decrypt = f
End Function
Public Function Encrypt(ToChange As String, Key As Long) As String
    'this can be used for an e-chat
    On Error Resume Next
    Dim a As Long, b As Long, c As String, d As Long
    Dim e As String, f As String
    a = Len(ToChange)
    For b = 1 To a
        c = Mid(ToChange, b, 1)
        d = Asc(c)
        d = d + Key
        If d > 255 Then
            d = d - 255
        End If
        e = Chr(d)
        f = f + e + " "
    Next b
    f = Left(f, Len(f) - 1)
    Encrypt = f
End Function

