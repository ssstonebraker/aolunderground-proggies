Attribute VB_Name = "Server_Bas"

Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SystemInfo)
Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lpReserved As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
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
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
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
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function enablewindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal Cmd As Long) As Long

Public Const SPI_SCREENSAVERRUNNING = 97

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

Public Const LBN_DBLCLK = 2

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
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

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

' INI File KeyNames
'
Global Const APP_NAME = "X-Treme Server '98 By TiTo"
Global Const APP_PREFERENCES = "Preferences"
Global Const KEY_INI = "INI Server"


Global Const twips = 1
Global Const pixels = 3
Global Const RES_INFO = 2
Global Const MINIMIZED = 1

Type MYVERSION
    lMajorVersion As Long
    lMinorVersion As Long
    lExtraInfo As Long
End Type

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type

Public Type SystemInfo
    dwOemId As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
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

Global Const VER_PLATFORM_WIN32s = 0
Global Const VER_PLATFORM_WIN32_WINDOWS = 1
Global Const VER_PLATFORM_WIN32_NT = 2

Global Const WF_CPU286 = &H2&
Global Const WF_CPU386 = &H4&
Global Const WF_CPU486 = &H8&
Global Const WF_STANDARD = &H10&
Global Const WF_ENHANCED = &H20&
Global Const WF_80x87 = &H400&

Global Const SM_MOUSEPRESENT = 19

Global Const GFSR_SYSTEMRESOURCES = &H0
Global Const GFSR_GDIRESOURCES = &H1
Global Const GFSR_USERRESOURCES = &H2


' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Dim WinVersion As Integer, SoundAvailable As Integer
Global VisibleFrame As Frame


'Server 'Variables
'
Global FirstMail As Integer
Global LastMail As Integer
Global RoomBust
Global PromptValue
Global MChatBot
Global IgnoreBot
Global ServerBot
Global mComm$
Global mChat$
Global RequestBot
Global StopBot
Global mChatText
Public m_MailList() As String 'The maillist vector
Public m_PeopleRequest As New Collection
Public m_PeopleFind As New Collection
Public m_PeopleList As New Collection 'The people in the list.
Public m_MailSubjectPrefix As String 'Prefix string
Public m_MailText As String 'The body text of a mail
Public m_DeleteFwd As Boolean 'This sets whether to take off the fwd
Public m_BlindCarbonCopied As Boolean 'This sets whether the names are blind carbon copied
Public m_ErrorPeople As New Collection 'The people who were taken off due to errors

Public Sub MailToListFlash(thelist As ListBox)
    
    Dim AOL As Long, mdi As Long, fMail As Long, fList As Long
    Dim Count As Long, MyString As String, AddMails As Long
    Dim sLength As Long, Spot As Long
    
    INI_FILENAME = App.Path + "\Settings\Server.ini"
    SERVER_FILENAME = App.Path + "\Server.dat"
    SERVER_FIND_FILENAME = App.Path + "\ServerFind.dat"
    Do
     DoEvents
     AOL& = FindWindow("AOL Frame25", vbNullString)
     mdi& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
     fMail& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
     fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    Loop Until fList& <> 0
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MyString$ = String(255, 0)
    On Error Resume Next
    Screen.MousePointer = 11
    Kill SERVER_FILENAME
    Open SERVER_FILENAME For Binary Access Write As #1
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(fList&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(fList&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        MyString$ = ReplaceString(MyString$, Chr(0), "")
        thelist.AddItem MyString$
        Server.Status = "Working On (Mail #" & Trim(Str$(AddMails&)) & " of " & Trim(Str$(Count&)) & ")"
        P$ = "(" & Trim(Str$(AddMails&)) & ")" & Chr(9) & MyString$ & Chr$(13) & Chr$(10)
        Put #1, LOF(1) + 1, P$
        Done% = AddMails&
        Call PercentBar(Server.Picture1, Done%, Count&)
    Next AddMails&
    Server.Picture1.Visible = False
    Screen.MousePointer = 0
Close #1
Server.Status = "Closing Incoming/Saved Mails Window."
Timeout 0.5
G = ShowWindow(fMail&, 2)
Status = "Preparing [" & Trim(Str$(List2.ListCount) - 1 & "] Mails")
Server.Status = "Ready."
End Sub
Public Sub MailToListNew(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long

    INI_FILENAME = App.Path + "\Settings\Server.ini"
    SERVER_FILENAME = App.Path + "\Server.dat"
    SERVER_FIND_FILENAME = App.Path + "\ServerFind.dat"
    Do
     DoEvents
      MailBox& = FindMailBox
      TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
      TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
      mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Loop Until mTree& <> 0
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = 11
    Kill SERVER_FILENAME
    Open SERVER_FILENAME For Binary Access Write As #1
    For AddMails& = 0 To Count& - 1
            DoEvents
           StringSpace$ = String(255, 0)
           Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, StringSpace$)
           StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
           StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
           StringSpace$ = Left(StringSpace$, InStr(1, StringSpace$, Chr$(0)) - 1)
           thelist.AddItem StringSpace$
           Server.Status = "Working On (Mail #" & Trim(Str$(AddMails&)) & " of " & Trim(Str$(Count&)) & ")"
           P$ = "(" & Trim(Str$(AddMails&)) & ")" & Chr(9) & StringSpace$ & Chr$(13) & Chr$(10)
           Put #1, LOF(1) + 1, P$
           Done% = AddMails&
           Call PercentBar(Server.Picture1, Done%, Count&)
    Next AddMails&
    Server.Picture1.Visible = False
   Screen.MousePointer = 0
Close #1
Server.Status = "Closing Mail Box Window."
Timeout 0.5
G = ShowWindow(MailBox&, 2)
Status = "Preparing [" & Trim(Str$(List2.ListCount) - 1 & "] Mails")
Server.Status = "Ready."
Call RunMenuByString("S&top Incoming Text")
End Sub
Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)
'This is used like:
'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)
'where Label1 is how many mails have
'already been forwarded, and Label2 is
'how many total mails there are.
On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = True
X = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), QBColor(0), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = QBColor(15) 'RGB(255, 0, 0)
Shape.Print Percent(Done, Total, 100) & "%"
End Sub
Public Sub PreVent()
If App.PrevInstance Then End
End Sub
Public Sub CenterForm(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub
Public Function MsgMessage() As String
Dim oWindow As Long, oButton As Long
Dim oStatic As Long, oString As String
    oWindow& = FindWindow("#32770", "America Online")
    oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    If oButton& <> 0& Then
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
        MsgMessage$ = oString$
    End If
End Function
Public Sub MailToListOld(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long

    
    INI_FILENAME = App.Path + "\Settings\Server.ini"
    SERVER_FILENAME = App.Path + "\Server.dat"
    SERVER_FIND_FILENAME = App.Path + "\ServerFind.dat"
    Do
     DoEvents
     MailBox& = FindMailBox
     TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
     TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
     TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
     mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Loop Until mTree& <> 0
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = 11
    Kill SERVER_FILENAME
    Open SERVER_FILENAME For Binary Access Write As #1
    For AddMails& = 0 To Count& - 1
        DoEvents
        StringSpace$ = String(255, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, StringSpace$)
        StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
        StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
        StringSpace$ = Left(StringSpace$, InStr(1, StringSpace$, Chr$(0)) - 1)
        thelist.AddItem StringSpace$
        Server.Status = "Working On (Mail #" & Trim(Str$(AddMails&)) & " of " & Trim(Str$(Count&)) & ")"
        P$ = "(" & Trim(Str$(AddMails&)) & ")" & Chr(9) & MyString$ & Chr$(13) & Chr$(10)
        Put #1, LOF(1) + 1, P$
        Done% = AddMails&
    Call PercentBar(Server.Picture1, Done%, Count&)
    Next AddMails&
    Server.Picture1.Visible = False
    Screen.MousePointer = 0
Close #1
Server.Status = "Closing Mail Box Window."
Timeout 0.5
G = ShowWindow(MailBox&, 2)
Status = "Preparing [" & Trim(Str$(List2.ListCount) - 1 & "] Mails")
Server.Status = "Ready."
Call RunMenuByString("S&top Incoming Text")
End Sub
Public Sub MailToListSent(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long

    
    
    INI_FILENAME = App.Path + "\Settings\Server.ini"
    SERVER_FILENAME = App.Path + "\Server.dat"
    SERVER_FIND_FILENAME = App.Path + "\ServerFind.dat"
    Do
    DoEvents
     MailBox& = FindMailBox
     TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
     TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
     TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
     TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
     mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Loop Until mTree& <> 0&
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    On Error Resume Next
    Screen.MousePointer = 11
    Kill SERVER_FILENAME
    Open SERVER_FILENAME For Binary Access Write As #1
    For AddMails& = 0 To Count& - 1
        DoEvents
        DoEvents
        StringSpace$ = String(255, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, StringSpace$)
        StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
        StringSpace$ = Mid$(StringSpace$, InStr(1, StringSpace$, Chr$(9)) + 1)
        StringSpace$ = Left(StringSpace$, InStr(1, StringSpace$, Chr$(0)) - 1)
        thelist.AddItem StringSpace$
        Server.Status = "Working On (Mail #" & Trim(Str$(AddMails&)) & " of " & Trim(Str$(Count&)) & ")"
        P$ = "(" & Trim(Str$(AddMails&)) & ")" & Chr(9) & MyString$ & Chr$(13) & Chr$(10)
        Put #1, LOF(1) + 1, P$
    Done% = AddMails&
    Call PercentBar(Server.Picture1, Done%, Count&)
    Next AddMails&
    Server.Picture1.Visible = False
    Screen.MousePointer = 0
Close #1
Server.Status = "Closing Mail Box Window."
Timeout 0.5
G = ShowWindow(MailBox&, 2)
Status = "Preparing [" & Trim(Str$(List2.ListCount) - 1 & "] Mails")
Server.Status = "Ready."
Call RunMenuByString("S&top Incoming Text")
End Sub
Public Sub LogLine(L As String)
If MenuForm.itemLog.Checked = False Then Exit Sub
On Error Resume Next
INI_FILENAME = App.Path + "\Settings\Server.ini"
SERVER_FILENAME = App.Path + "\Server.dat"
SERVER_FIND_FILENAME = App.Path + "\ServerFind.dat"
SERVER_LOG = App.Path + "\Server.txt"
F = FreeFile
Server.Status = "Login InFo To File"
Open SERVER_LOG For Binary Access Write As F
P$ = L & Chr$(13) & Chr$(10)
Put #1, LOF(1) + 1, P$
Close F
End Sub
Public Sub LogDead(What As String)
On Error Resume Next
SERVER_DEAD_LOG = App.Path + "\Dead.Log"
F = FreeFile
Open SERVER_DEAD_LOG For Binary Access Write As F
P$ = What$ & Chr$(13) & Chr$(10)
Put #1, LOF(1) + 1, P$
Close F
End Sub
Public Sub SendChat(text As String)
'This is a sub that sends text to the chat room.
Dim Holder As Long
   Holder = FindWindowEx(AOLFindChatRoom, 0, "RICHCNTL", vbNullString)
    Holder = FindWindowEx(AOLFindChatRoom, Holder, "RICHCNTL", vbNullString)
    SendMessageByString Holder, WM_SETTEXT, 0, text
    SendMessage Holder, WM_CHAR, VK_RETURN, 0
End Sub
Public Sub Timeout(Duration)
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop
End Sub
Public Sub BoldSendChat(BoldChat)
SendChat ("<b>" & BoldChat & "</b>")
End Sub
Public Sub CenterFormTop(frm As Form)
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub
Public Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub AddRoomToListbox(ListBox As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Dim Room As Long
thelist.Clear

Room = AOLFindChatRoom
aolhandle = FindChildByClass(Room&, "_AOL_Listbox")

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
If Person$ = AOLUserSN Then GoTo Na
ListBox.AddItem Person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Public Function AOLFindChatRoom() As Long
    Dim AOL As Long, mdi As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        AOLFindChatRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                AOLFindChatRoom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    AOLFindChatRoom& = child&
End Function
