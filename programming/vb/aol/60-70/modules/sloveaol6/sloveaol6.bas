Attribute VB_Name = "aol6"
'sloves aol6.0 module
'made almost entirely by slove
'some contributions by:
'turtle - ringmaster - nut - progee - masta
'66 subs/functions
'
'this is just a beta, i will be adding
'on more and it will be bigger and better

'greets: turtle, ringmaster, nut, nautica
'xtc, kazan, amy, fallen, notorious,
'jmaan, magus, nk, ignite, nofx, jmaan
'masta, lithe, hail and all my friends
'i better not catch digital stealing
'my coding and taking credit for it
'transcend 2000

'------Coding Begins------'

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
'declares for system tray start
'Constants for the ACTION PROPERTY
Public Const sys_Add = 0       'Specifies that an icon is being add
Public Const sys_Modify = 1    'Specifies that an icon is being modified
Public Const sys_Delete = 2    'Specifies that an icon is being deleted
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
'Constants for ERROR MESSAGE
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETLINE = &HC4
Public Const EM_LINELENGTH = &HC1


'Constants for the ERROR EVENT
Public Const errUnableToAddIcon = 1    'Icon can not be added to system tray
Public Const errUnableToModifyIcon = 2 'System tray icon can not be modified
Public Const errUnableToDeleteIcon = 3 'System tray icon can not be deleted
Public Const errUnableToLoadIcon = 4   'Icon could not be loaded (occurs while using icon property)

'Constants for MOUSE RELATETED EVENTS
Public Const vbLeftButton = 1     'Left button is pressed
Public Const vbRightButton = 2    'Right button is pressed
Public Const vbMiddleButton = 4   'Middle button is pressed

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)


Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Const cdlOpen = 1
Public Const cdlSave = 2
Public Const cdlColor = 3
Public Const cdlPrint = 4
Public Const cdlOFNReadOnly = 1             'Checks Read-Only check box for Open and Save As dialog boxes.
Public Const cdlOFNOverwritePrompt = 2      'Causes the Save As dialog box to generate a message box if the selected file already exists.
Public Const cdlOFNHideReadOnly = 4         'Hides the Read-Only check box.
Public Const cdlOFNNoChangeDir = 8          'Sets the current directory to what it was when the dialog box was invoked.
Public Const cdlOFNHelpButton = 10          'Causes the dialog box to display the Help button.
Public Const cdlOFNNoValidate = 100         'Allows invalid characters in the returned filename.
Public Const cdlOFNAllowMultiselect = 200   'Allows the File Name list box to have multiple selections.
Public Const cdlOFNExtensionDifferent = 400 'The extension of the returned filename is different from the extension set by the DefaultExt property.
Public Const cdlOFNPathMustExist = 800      'User can enter only valid path names.
Public Const cdlOFNFileMustExist = 1000     'User can enter only names of existing files.
Public Const cdlOFNCreatePrompt = 2000      'Sets the dialog box to ask if the user wants to create a file that doesn't currently exist.
Public Const cdlOFNShareAware = 4000        'Sharing violation errors will be ignored.
Public Const cdlOFNNoReadOnlyReturn = 8000  'The returned file doesn't have the Read-Only attribute set and won't be in a write-protected directory.
Public Const cdlOFNExplorer = 8000          'Use the Explorer-like Open A File dialog box template.  (Windows 95 only.)
Public Const cdlOFNNoDereferenceLinks = 100000
Public Const cdlOFNLongNames = 200000
'declares for system tray end

'dim varibles
Dim m_sLineString As String * 1056
Dim m_lngRet As Long
Dim m_sRetString As String
'dim varibles end

Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

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
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long



Private Type OPENFILENAME
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
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const CB_SETCURSEL = &H14E
Public Const CB_GETCOUNT = &H146
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_FINDSTRING = &H18F
Global Const LB_FINDSTRINGEXACT = &H1A2&
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_NORMAL = 1
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3

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


Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const ewx_shutdown = 1

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000



Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type


Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const LB_DIR = &H18D
       Public Const DDL_READWRITE = &H0
       Public Const DDL_READONLY = &H1
       Public Const DDL_HIDDEN = &H2
       Public Const DDL_SYSTEM = &H4
       Public Const DDL_DIRECTORY = &H10
       Public Const DDL_ARCHIVE = &H20
       Public Const DDL_DRIVES = &H4000
       Public Const DDL_EXCLUSIVE = &H8000
       Public Const DDL_POSTMSGS = &H2000
       Public Const DDL_FLAGS = DDL_ARCHIVE Or DDL_DIRECTORY

Public Type POINTAPI
        X As Long
        Y As Long
        i As Long
End Type
Sub RunMenu(MenuNumber, CommandNumber)
'written by progee
'Menu Numbers -----------------
'1-mail, 2-people, 3-aol services, 4-settings, 5-favorites

'Command Numbers ----------
'mail (1)-----------------------------
'new mail - 7170
'old mail - 7171
'sent mail - 7172
'write mail - 7173
'address book - 7175
'mail center - 7176
'recently deleted mail - 7178
'pfc - 7179
'mail waiting to be sent - 7180
'auto aol - 7181
'mail signatures - 7183
'mail controls - 7184
'mail preferences - 7185
'greetings & mail extras - 7186
'people (2)-----------------------------
'im - 7169
'people connection - 7170
'chat now - 7172
'find a chat - 7173
'start your own chat - 7174
'live events - 7175
'buddy list - 7177
'get profile - 7178
'locate - 7179
'send message to pager - 7180
'sign on a friend - 7181
'invitations - 7185
'member directory - 7186
'aol services (3)-----------------------------
'internet connection - 7171
'go to the web - 7173
'search the web - 7174
'ftp - 7177
'add to my calender - 7179
'aol help - 7180
'download center - 7184
'settings (4)-----------------------------
'my aol - 7169
'parental controls - 7172
'my profile - 7173
'screennames - 7174
'passwords - 7175
'billing center - 7177
'favorites (5)-----------------------------
'favorite places - 7188
'preferences - 7171
'add top window to favorites - 7190
'go to kw - 7191
'edit hotkeys - 7194
'hotkey1 - 7196
'hotkey2 - 7197
'hotkey3 - 7198
'hotkey4 - 7199
'hotkey5 - 7200
'hotkey6 - 7201
'hotkey7 - 7202
'hotkey8 - 7203
'hotkey9 - 7204
'hotkey0 - 7205

a = FindWindow("aol frame25", vbNullString)
B = FindWindowEx(a, 0, "AOL Toolbar", vbNullString)
c = FindWindowEx(B, 0, "_AOL_Toolbar", vbNullString)

Select Case MenuNumber
Case 1
   D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
Case 2
   D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
   For X = 1 To 3
   D = GetWindow(D, 2)
   Next X
Case 3
   D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
   For X = 1 To 6
   D = GetWindow(D, 2)
   Next X
Case 4
   D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
   For X = 1 To 9
   D = GetWindow(D, 2)
   Next X
Case 5
   D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
   For X = 1 To 11
   D = GetWindow(D, 2)
   Next X
End Select

PostMessage D, &H201, 0, 0
PostMessage D, &H202, 0, 0
PostMessage D, &H201, 0, 0
PostMessage D, &H202, 0, 0
Do
e = FindWindow("#32768", vbNullString)
DoEvents
Loop Until e > 0
PostMessage a, 273, CommandNumber, 0
End Sub
Public Sub AOL6_ViewBuddy()
'written by slove
RunMenu "2", "7177"
End Sub


Sub AddRoom(lst As listbox)
'for aol5 written by progee
Room& = FindRoom()
ListX& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
Count& = SendMessageByNum(ListX&, &H18B, 0, 0)
Buffer$ = Space(255)
For Counter& = 0 To Count& - 1
List& = AOLGetList(ListX&, Counter&, Buffer$)
For e = 0 To lst.ListCount
    If lst.List(e) = Buffer$ Then
    GoTo Here:
    End If
Next e
If Buffer$ = GetUser Then GoTo Here
lst.AddItem (Buffer$)
Here:
Next Counter&
End Sub


Function AOLGetList(LBHandle, Index, Buffer As String)
'written by progee
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
AOLThread = GetWindowThreadProcessId(LBHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or &HF0000, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(LBHandle, &H199, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function



Sub Ghost_On()
'written by progee
a = FindWindow("aol frame25", vbNullString)
B = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(B, 0, "AOL Child", "Buddy List")
D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
For e = 1 To 6
D = GetWindow(D, 2)
Next e
IconClick D
Do
f = FindWindowEx(B, 0, "AOL Child", "Buddy List Setup")
g = FindWindowEx(f, 0, "_AOL_Edit", vbNullString)
DoEvents
Loop Until g > 0
g = GetWindow(g, 3)
g = GetWindow(g, 3)
IconClick g
Do
H = FindWindowEx(B, 0, "AOL Child", "Buddy List Preferences")
i = FindWindowEx(H, 0, "_AOL_TabControl", vbNullString)
j = FindWindowEx(i, 0, "_AOL_TabPage", vbNullString)
j2 = GetWindow(j, 2)
j3 = GetWindow(j2, 2)
K = FindWindowEx(j3, 0, "_AOL_RadioBox", vbNullString)
DoEvents
Loop Until K > 0
For L = 1 To 4
K = GetWindow(K, 2)
Next L
IconClick K
For L2 = 1 To 3
K = GetWindow(K, 2)
Next L2
IconClick K
M = FindWindowEx(H, 0, "_AOL_Icon", vbNullString)
IconClick M
Pause 1
PostMessage f, WM_CLOSE, 0, 0
End Sub


Sub Ghost_Off()
'written by progee
a = FindWindow("aol frame25", vbNullString)
B = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(B, 0, "AOL Child", "Buddy List")
D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
For e = 1 To 6
D = GetWindow(D, 2)
Next e
IconClick D
Do
f = FindWindowEx(B, 0, "AOL Child", "Buddy List Setup")
g = FindWindowEx(f, 0, "_AOL_Edit", vbNullString)
DoEvents
Loop Until g > 0
g = GetWindow(g, 3)
g = GetWindow(g, 3)
IconClick g
Do
H = FindWindowEx(B, 0, "AOL Child", "Buddy List Preferences")
i = FindWindowEx(H, 0, "_AOL_TabControl", vbNullString)
j = FindWindowEx(i, 0, "_AOL_TabPage", vbNullString)
j2 = GetWindow(j, 2)
j3 = GetWindow(j2, 2)
K = FindWindowEx(j3, 0, "_AOL_RadioBox", vbNullString)
DoEvents
Loop Until K > 0
For L = 1 To 2
K = GetWindow(K, 2)
Next L
IconClick K
For L2 = 1 To 5
K = GetWindow(K, 2)
Next L2
IconClick K
M = FindWindowEx(H, 0, "_AOL_Icon", vbNullString)
IconClick M
Pause 1
PostMessage f, WM_CLOSE, 0, 0
End Sub
Public Sub AOL5_GhostOff()
'written by slove
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Window")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetUser & "'s Buddy Lists")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
aolcheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 5&
    aolcheckbox& = FindWindowEx(AOLChild&, aolcheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call IconClick(aolcheckbox&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
aolcheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call IconClick(aolcheckbox&)

AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 1
child& = FindWindow("#32770", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
Call IconClick(Button)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetUser & "'s Buddy Lists")
Killit = SendMessage(AOLChild&, WM_CLOSE, 0, 0&)

End Sub

Public Sub AOL5_GhostOn()
'written by slove
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Window")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetUser & "'s Buddy Lists")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
aolcheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 4&
    aolcheckbox& = FindWindowEx(AOLChild&, aolcheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call IconClick(aolcheckbox&)
Pause 0.5
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
aolcheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 6&
    aolcheckbox& = FindWindowEx(AOLChild&, aolcheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call IconClick(aolcheckbox&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 1
child& = FindWindow("#32770", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
Call IconClick(Button)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetUser & "'s Buddy Lists")
Killit = SendMessage(AOLChild&, WM_CLOSE, 0, 0&)

End Sub
Public Sub AOL6_FTP()
'written by slove
'signs on your aol ftp
RunMenu "3", "7177"
Pause 3
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", " FTP - File Transfer Protocol")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
Pause 6
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Anonymous FTP")
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
Call SendMessageByNum(AOLListbox&, LB_SETCURSEL, "4", 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Anonymous FTP")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
Pause 3
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
End Sub
Public Sub AOL6_SearchFor(What As String)
'written by slove
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLEdit& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit& = FindWindowEx(AOLToolbar2&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, What$)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 23&
    AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
IconClick AOLIcon&
End Sub

Public Function FindRoom() As Long
'written by slove
    Dim aol As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindRoom& = child&
End Function
Public Sub Pause(Duration As Long)
Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Sub drag(frm As Form)
Call ReleaseCapture
    Call SendMessage(frm.hwnd, WM_SYSCOMMAND, WM_MOVE, vbNullString)
End Sub
Public Sub CloseChat()
'written by slove
aol = FindWindow("aol frame25", vbNullString)
MDI = FindWindowEx(aol, 0&, "MDIClient", vbNullString)
Chil = FindWindowEx(MDI, 0&, "AOL Child", "progs")
Call SendMessage(Chil, WM_CLOSE, 0&, 0&)

End Sub
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
'taken from dos32
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
Public Function FileExists(FN As String) As Boolean
'written by slove
If Len(FN$) = 0 Then
FileExists = False
Exit Function
End If
If Len(Dir$(FN$)) Then
FileExists = True
Else
FileExists = False
End If
End Function


Public Sub AOL6_MailSend(psn As String, Subj As String, Msg As String)
'written by slove
KeyWord "mailto:" & psn
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Subj)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, Msg)

AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 17&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
End Sub

Public Function GetfromINI(Section As String, Key As String, Directory As String) As String
'written by slove
Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetfromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WritetoINI(Section As String, Key As String, KeyValue As String, Directory As String)
Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub













 
Public Sub AOL6_AddFav(descr As String, URL As String)
'written by slove
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 12&
    AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 1
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
Pause 1
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Add New Folder/Favorite Place")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, descr)

AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Add New Folder/Favorite Place")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, URL)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Add New Folder/Favorite Place")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub Chat_PressBold()
'written by slove
'simple, just presses bold key
Room = FindRoom
AOLIcon& = FindWindowEx(Room, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
End Sub
Public Sub AOL6_NewAwayMsg(Title As String, Msg As String)
'written by slove
'this like the other away subs toggle
'between on and off
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 1
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Away Message")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
Pause 1
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "New Away Message")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Title)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, Msg)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 8&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 1
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Away Message")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 1
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Away Message Off")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)


End Sub
Public Function OnListAsString(thing As String, List As Control) As String
'written by nut i believe
Dim i As Long, shit As String, fuck As Long
For i& = 0 To List.ListCount
shit$ = LCase(Trimnull(trimspaces(List.List(i))))
    If InStr(shit$, thing$) <> 0& Then OnListAsString$ = List.List(i&): Exit Function
Next i&
OnListAsString$ = ""
End Function
Function Trimnull(What As String) As String
'i dont know where i got this
Dim wstr As String, xx As Long, this_chr As Long, wordd As String
wstr$ = Trim(What$)
Do Until xx& = Len(What$)
Let xx& = xx& + 1
Let this_chr& = Asc(Mid$(What$, xx&, 1))
If this_chr& > 31 And this_chr& <> 256 Then Let wordd$ = wordd$ & Mid$(What$, xx&, 1)
Loop
Trimnull$ = wordd$
End Function
Function trimspaces(text As String) As String
'this one either
    Dim TheChar, TrimSpace
    Dim TheChars
    If InStr(text, " ") = 0 Then
        trimspaces = text
        Exit Function
    End If
    For TrimSpace = 1 To Len(text)
        TheChar = Mid(text, TrimSpace, 1)
        TheChars = TheChars & TheChar
        If TheChar = " " Then
            TheChars = Mid(TheChars, 1, Len(TheChars) - 1)
        End If
    Next TrimSpace
    trimspaces = TheChars
End Function
Function AddFontsToList(lst As listbox)
Dim Fontz As Variant
For Fontz = 0 To Screen.FontCount - 1
    lst.AddItem Screen.Fonts(Fontz)
Next Fontz
End Function
Sub AOL6_ChatPrefs()
'written by slove
'this changes the option of having
'onlinehost tell if a member left
'or entered the room
AOLChild& = FindRoom
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 12&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
IconClick (AOLIcon&)
Pause 1
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
aolcheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
IconClick (aolcheckbox&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
aolcheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
aolcheckbox& = FindWindowEx(AOLModal&, aolcheckbox&, "_AOL_Checkbox", vbNullString)
IconClick (aolcheckbox&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
IconClick (AOLIcon&)
End Sub
Public Function LineCount(MyString As String) As Long
'written by nut
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
Public Sub Scroll(ScrollString As String)
'written by nut
Dim CurLine As String, Count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindRoom& = 0 Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    Count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To Count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = Left(CurLine$, 92)
            End If
            Call ChatSend(CurLine$)
            Pause 0.7
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Sub
Public Function LineFromString(MyString As String, Line As Long) As String
'written by nut
Dim theline As String, Count As Long
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
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
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
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function
Sub AOL6_AddRoom(List As listbox, AddUser As Boolean)
'if you are wondering if i got addroom,
'yes i certainly have. yet i dont want
'to give it to you people cause it took
'a long hella time to get it and it's
'still being worked on. I don't think
'i will give it to anyone cept REALLY
'close friends.
End Sub



Public Sub List_KillDupes(listbox As listbox)
'written by turtle
Dim Search1 As Long
Dim Search2 As Long
Dim KillDupe As Long
KillDupe = 0
   For Search1& = 0 To listbox.ListCount - 1
       For Search2& = Search1& + 1 To listbox.ListCount - 1
           KillDupe = KillDupe + 1
           If listbox.List(Search1&) = listbox.List(Search2&) Then
               listbox.RemoveItem Search2&
               Search2& = Search2& - 1
           End If
       Next Search2&
   Next Search1&
End Sub
 
 
Public Function AOL6_GetText(WindowHandle As Long) As String
 'written by slove
 'works on win2k
   Dim Buffer As String, TextLength As Long, txtlen As Long, lineNum As Long
    TextLength& = SendMessage(WindowHandle, EM_GETLINECOUNT, 0&, 0&)
    lineNum = TextLength&
   txtlen = SendMessageByNum(WindowHandle, EM_LINELENGTH, lineNum, 0&)
   Buffer = String(txtlen, 0&)
   Call SendMessage(WindowHandle, EM_GETLINE, lineNum, ByVal Buffer$)
    AOL6_GetText$ = Buffer$
End Function
Public Sub LoadComboBox(ByVal Directory As String, Combo As ComboBox)
'taken from dos32
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
Public Sub AOL6_IM(psn As String, Msg As String)
'written by slove
Call KeyWord("aol://9293:" & psn)
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, Msg)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 9&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)

End Sub

Public Sub AOL6_Away_Default()
'written by slove
'this toggles being on and off
'this turns on the first message
'in the list, if none are there
'aol will let you know
'needs work
Dim AOLFrame&
Dim MDIClient&
Dim AOLChild&
Dim AOLIcon&
Dim AOLIcon2&
Dim AOLEdit&
Dim RICHCNTL&
AOLFrame& = FindWindow("aol frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "mdiclient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "aol child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
Call IconClick(AOLIcon&)
Call SendMessage(AOLIcon&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
AOLFrame& = FindWindow("aol frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "mdiclient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "aol child", "away message")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
Call IconClick(AOLIcon&)
Call SendMessage(AOLIcon&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
AOLFrame& = FindWindow("aol frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "mdiclient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "aol child", "away message off")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_aol_icon", vbNullString)
Call IconClick(AOLIcon&)
Call SendMessage(AOLIcon&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub IMsOff()
'written by slove
'on the new aol6 beta the ims off
'are back to normal
Call AOL6_IM("$IM_OFF", "slove")
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub IMsOn()
'written by slove
Call AOL6_IM("$IM_ON", "slove")
Pause 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Sub
Public Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
'written by slove
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Function List_Search(Wat As String, Lis As listbox)
'written by slove
Dim i As Integer
For i = 0 To Lis.ListCount
L$ = LCase(Lis.List(i))
If InStr(L$, LCase(Wat)) <> 0 Then
List_Search = True
Exit Function
End If
Next i
List_Search = False
End Function
Public Sub Loadlistbox(Directory As String, thelist As listbox)
'taken from dos32
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
Public Sub ProfileScroll(sn As String, noprofile As String)
'written by slove
'choose what to say if there is no
'profile
Dim text As String
RunMenu 2, 7178
Pause 1
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Get a Member's Profile")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, sn)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
IconClick AOLIcon&
Pause 1
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Member Profile")
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
text = GetText(RICHCNTL&)
If text = "" Then
child& = FindWindow("#32770", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
IconClick Button&
ChatSend noprofile
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Get a Member's Profile")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
Exit Sub
Else
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
Scroll text
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Get a Member's Profile")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
Exit Sub
End If

End Sub
Public Sub SaveListBox(Directory As String, thelist As listbox)
'taken from dos32
Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub SendChat(what1 As String, what2 As String)
'written by slove
'dont use this it's for annihilation 2.0
If Form1.txtfade = "1" Then GoTo fade
If Form1.txtfade = "0" Then GoTo nofade
fade:
Dim html As String
html = FadeByColor3(Form1.color1, Form1.color2, Form1.color3, Form1.ascii & " " & what1 & " " & Form1.iascii & " " & what2, False)
ChatSend html
Exit Sub
nofade:
ChatSend "<font color=#" & Form1.acolor.text & "><font face=" & Form1.afont.text & ">" & Form1.ascii.text & " <font color=#" & Form1.fcolor & "><font face=" & Form1.ffont.text & ">" & what1 & " " & Form1.iascii & " " & what2 & ""
Exit Sub
End Sub
Public Sub programtop(frm As Form)
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub


Public Function GetUser() As String
'written by slove
Dim aol As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    GetUser$ = ""
End Function
Public Sub OpenBuddy()
'written by slove
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
End Sub
Sub List_DeleteName(lst As Control, sn$)
'written by masta
Dim i As Long
For i = 0 To lst.ListCount - 1
If UCase(sn$) Like UCase(lst.List(i)) Then lst.RemoveItem i
Next i
End Sub
Public Sub Channels()
'written by slove
'hides or shows channels
Dim AOLFrame&
Dim AOLToolbar&
Dim AOLIcon&
AOLFrame& = FindWindow("aol frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "aol toolbar", vbNullString)
AOLToolbar& = FindWindowEx(AOLToolbar&, 0&, "_aol_toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, 0&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar&, AOLIcon&, "_aol_icon", vbNullString)
Call IconClick(AOLIcon&)
Call SendMessage(AOLIcon&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Sub
Public Sub SetText(Window As Long, text As String)
Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub
Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Sub SaveText(txtSave As TextBox, Path As String)
'written by slove
Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub ChatSend(strmessage As String)
'written by slove
On Error Resume Next
Dim lngchat As Long, lngrich As Long, strtext As String, WM_CLEAR As Data
    
    Let lngchat& = FindRoom&
    Let lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
    Let lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
    Let strtext$ = GetText(lngrich&)
    If strtext$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until GetText$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strtext$)
    End Sub

Sub ChatCommand()
'not done yet
End Sub
Sub IconClick(wndo)
'written by progee & slove
'You have to have the cursor over the icon
'for the   click to process for some weird
'reason. You can bearly see the cursor   move
'though so it works out nice.

Dim g As RECT
Dim H As POINTAPI
GetWindowRect wndo, g
GetCursorPos H
SetCursorPos g.Left + (g.Right - g.Left) / 2, g.Top + (g.Bottom - g.Top) / 2

SendMessageByNum wndo, &H201, 0, 0
SendMessageByNum wndo, &H202, 0, 0

SetCursorPos H.X, H.Y
End Sub
Public Sub MailNew()
Call RunMenu("1", "7170")
End Sub
Public Sub Mailold()
Call RunMenu("1", "7171")
End Sub
Public Sub MailSent()
Call RunMenu("1", "7172")
End Sub


Public Function GetText(WindowHandle As Long) As String
'written by slove
Dim Buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, Buffer$)
    GetText$ = Buffer$
End Function
Public Sub AOL6_BuddyAdd(text As String)
'written by slove
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Pause 0.5
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Setup")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, text$)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Setup")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call IconClick(AOLIcon&)
Call SendMessage(AOLIcon&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Sub
Function RunExe(PathName As String, WinStyle As Integer)
'written by slove
'runs an exe
'WinStyle constants are:
'0   Window is hidden and focus is passed to the hidden window.
'1   Window has focus and is restored to its original size and position.
'2   Window is displayed as an icon with focus.
'3   Window is maximized with focus.
'4   Window is restored to its most recent size and position.  The currently active window remains active.
'6   Window is displayed as an icon.  The currently active window remains active.
X = Shell(PathName, WinStyle)
End Function
Public Sub HWW()
'written by slove
'hide welcome window
aol = FindWindow("AOL Frame25", vbNullString)
MDI = FindWindowEx(aol, 0&, "MDIClient", vbNullString)
child = FindWindowEx(MDI, 0&, "AOL Child", " Welcome,")
If child <> 0 Then
X = ShowWindow(child, SW_HIDE)
End If
End Sub
Public Sub SWW()
'written by slove
'show welcome window
aol = FindWindow("AOL Frame25", vbNullString)
MDI = FindWindowEx(aol, 0&, "MDIClient", vbNullString)
child = FindWindowEx(MDI, 0&, "AOL Child", " Welcome,")
If child <> 0 Then
X = ShowWindow(child, SW_SHOW)
End If
End Sub
Public Sub BuddyAdd(sn As String)
'written by slove
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Setup")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, sn)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call IconClick(AOLIcon&)
End Sub
Public Sub KeyWord(KW As String)
    'written by slove
Dim aol As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Function GetCaption(WindowHandle As Long) As String
'written by slove
Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
