Attribute VB_Name = "Timeonlinebas"
'All subs from this .bas were taken from Chronx.bas v3.0 BETA
'If you use these subs or this form please add me to your greets
'Chron__x@hotmail.com
'http://chronx.cjb.net



Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Public Declare Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As String
Public Declare Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetProfileString& Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long)
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer

Public Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Public Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function MciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal Ch As Byte, ByVal dwhkl As Long) As Integer

Public Declare Function WindowFromPointXy Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
 ByVal bRedraw As Boolean) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const Cb_AddString& = &H143
Public Const Cb_DeleteString& = &H144
Public Const Cb_FindStringExact& = &H158
Public Const Cb_GetCount& = &H146
Public Const Cb_GetItemData = &H150
Public Const Cb_GetLbText& = &H148
Public Const Cb_ResetContent& = &H14B
Public Const Cb_SetCursel& = &H14E

Public Const Em_GetLineCount& = &HBA
Public Const ENTER_KEY = 13

Public Const Hwnd_NotTopMost = -2
Public Const HWND_TOPMOST = -1

Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_INSERTSTRING = &H181
Public Const Lb_ResetContent& = &H184
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const MF_BYPOSITION = &H400&

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const Snd_Flags = SND_ASYNC Or SND_NODEFAULT
Public Const Snd_Flags2 = SND_ASYNC Or SND_LOOP

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE& = 3
Public Const SW_MINIMIZE& = 6
Public Const SW_RESTORE& = 9
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Swp_Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const Sys_Add = &H0
Public Const Sys_Delete = &H2
Public Const Sys_Message = &H1
Public Const Sys_Icon = &H2
Public Const Sys_Tip = &H4

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
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
Public Const WM_MOVE = &HF012
Public Const Wm_RButtonDblClk& = &H206
Public Const Wm_RButtonDown& = &H204
Public Const Wm_RButtonUp& = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_USER& = &H400

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const Op_Flags = PROCESS_READ Or RIGHTS_REQUIRED

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public SysTray As NOTIFYICONDATA

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public Function Timeonline() As String
    Dim AolModal As Long, AolIcon As Long, AolStatic As Long
    Call PopUpIcon(5, "O")
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", "America Online")
        AolIcon& = FindWindowEx(AolModal&, 0, "_AOL_Icon", vbNullString)
        AolStatic& = FindWindowEx(AolModal&, 0, "_AOL_Static", vbNullString)
    Loop Until AolModal& <> 0& And AolIcon& <> 0& And AolStatic& <> 0&
    Timeonline$ = GetText(AolStatic&)
    Call PostMessage(AolIcon&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(AolIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function


Public Sub PopUpIcon(IconNumber As Long, Character As String)
    Dim Message1 As Long, Message2 As Long, AolFrame As Long
    Dim AolToolbar As Long, Toolbar As Long, AolIcon As Long
    Dim NextOfClass As Long, AscCharacter As Long
    Message1& = FindWindow("#32768", vbNullString)
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    For NextOfClass& = 1 To IconNumber&
        AolIcon& = GetWindow(AolIcon&, 2)
    Next NextOfClass&
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Message2& = FindWindow("#32768", vbNullString)
    Loop Until Message2& <> Message1&
    AscCharacter& = Asc(Character$)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
End Sub

Public Function GetText(WinHandle As Long) As String
    Dim Buffer As String, TextLen As Long
    TextLen& = SendMessageByNum(WinHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(TextLen&, 0&)
    Call SendMessageByString(WinHandle&, WM_GETTEXT, TextLen& + 1, Buffer$)
    GetText$ = Buffer$
End Function
