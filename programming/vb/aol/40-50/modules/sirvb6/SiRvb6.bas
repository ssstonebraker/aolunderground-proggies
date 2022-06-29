Attribute VB_Name = "SiRvb6"
Option Explicit
'SiRvb6.bas™
'Version 1.00
'Last update:  6/20/99
'SEND ALL BAS FILE ERRORS TO vbSiR@juno.com
'This bas module was made for Visual Basic 6, AOL4, and AIM 2.0+
'I tried to comment on everything and make it as simple as possible
'if you have any Questions or Comments you can email me at
' vbSiR@juno.com or contact me on AIM:  V IB 6
'I reccommend any beginner programmer to study the code in this bas
'if you can Find a window, Set text to the window's edit box, then push
'a button on the window, you can pretty much make any kind of aol program
'hell, thats all you need to know for 80% of aol programming.
'The only codes i did not write myself, are the addroom for aol (Not sure who made it first)
'i based mine on the original addroom and the run menu subs, By the way, i am the worlds worst
'code indenter LOL so if you don't like my style you can go through and indent the
'stuff the way you want it.
'L8r
'SiR
'NOTE: some of this bas can be used in vb4 & 5 but certain parts only work
'for visual basic 6, however it is nothing major that you couldn't make yourself
'and change around to work.
'(Scroll all the way down to see more comments for vb4 & vb5 users)
'=========================
'Shout Outs:
'Anubis - I read through every part of the vb3 core_api help file, and it taught me alot.   http://reapers.org
'Pat Or JK - always makes cool stuff, i use his api spy all the time.  Its pretty sweet.   http://www.patorjk.com
'NEON - he's only programed for a few months and is already putting out some kickass stuff, making everything himself.
'Gabo - he is the best programmer that nobody has heard of. I know he puts alot of the "ao-famous" programmers to shame.
'Others:  Bale, airwalk, Shocker, Rey, pawn, CoDa, chao, deexpimp, Ninj0r,
'             Pablo, Stevai, Acer and whoever else i forgot
'~~~~~~~~~~~~~~~~~~~~~~~~~
'finding windows
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'sendmessages
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
'manip windows
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
'used instead of doevents
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'mouse&cursor stuff
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
'adding rooms 2 list
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
'menu stuff
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
'mouse in stuff
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'play sound wav
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'move form without titlebar
Public Declare Sub ReleaseCapture Lib "user32" ()
'rebooting system
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
'flashing a window
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
'get a file's path shortname
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'CD rom & Sound control declare
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'move mouse pointer
Public Declare Function SetCursorPosition& Lib "user32" (ByVal X As Long, ByVal y As Long)
'needed for the api spy
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
'disable control alt delete
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
'ini stuff
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
'minimize window
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
'delete file
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFilename As String) As Long
'purge clipboard
Public Declare Function EmptyClipboard Lib "user32" () As Long
'get internal tick count for timeout's
Public Declare Function GetTickCount Lib "kernel32" () As Long
'========================
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
'=========================
Public Const SW_SHOWNORMAL = 1
Public Const SW_ShowMinimized = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_Hide = 0
Public Const SW_MAX = 10
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
'=========================
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_TRANSPARENT = &H20&
'=========================
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const SWP_SHOWME = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE
'=========================
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_UNDO = &HC7
Public Const EM_SETREADONLY = &HCF
Public Const EM_LIMITTEXT = &HC5
'=========================
Public Const SND_SYNC = 0
Public Const SND_ASYNC = 1
Public Const SND_NODEFAULT = 2
Public Const SND_LOOP = 8
Public Const SND_NOSTOP = 16
'=========================
Public Const HTCAPTION = 2
Public Const EWX_REBOOT = 2
'=========================
Public Const WM_CHAR = &H102
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SYSCOMMAND = &H112
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
'=========================
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
'=========================
Public Const LB_GETITEMDATA = &H199
Public Const LB_RESETCONTENT = &H184
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
'=========================
Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146
Public Const CB_GETCURSEL = &H147
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETLBTEXT = &H148
Public Const CB_GETITEMDATA = &H150
'=========================
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
'=========================
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
'=========================
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
'=========================
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
'=========================
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
'=========================
Public Const SPI_SCREENSAVERRUNNING = 97
Public Const SC_SCREENSAVE = &HF140
'=========================
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Public Const WM_MENUSELECT = &H11F
Public Const MIIM_TYPE = &H10
Public Const MIIM_ID = 2
Public Const MF_MOUSESELECT = &H8000&
'=========================
Public Enum VirtualKeys
    ADD_VKEY = &H6B
    BACK_VKEY = &H8
    CAPITAL_VKEY = &H14
    CLEAR_VKEY = &HC
    CONTROL_VKEY = &H11
    CRSEL_VKEY = &HF7
    DECIMAL_VKEY = &H6E
    DELETE_VKEY = &H2E
    DIVIDE_VKEY = &H6F
    DOWN_VKEY = &H28
    END_VKEY = &H23
    ESCAPE_VKEY = &H1B
    F1_VKEY = &H70
    F2_VKEY = &H71
    F3_VKEY = &H72
    F4_VKEY = &H73
    F5_VKEY = &H74
    F6_VKEY = &H75
    F7_VKEY = &H76
    F8_VKEY = &H77
    F9_VKEY = &H78
    F10_VKEY = &H79
    F11_VKEY = &H7A
    F12_VKEY = &H7B
    F13_VKEY = &H7C
    F14_VKEY = &H7D
    F15_VKEY = &H7E
    F16_VKEY = &H7F
    F17_VKEY = &H80
    F18_VKEY = &H81
    F19_VKEY = &H82
    F20_VKEY = &H83
    F21_VKEY = &H84
    F22_VKEY = &H85
    F23_VKEY = &H86
    F24_VKEY = &H87
    HOME_VKEY = &H24
    INSERT_VKEY = &H2D
    LBUTTON_VKEY = &H1
    LCONTROL_VKEY = &HA2
    LEFT_VKEY = &H25
    LSHIFT_VKEY = &HA0
    MULTIPLY_VKEY = &H6A
    NUMLOCK_VKEY = &H90
    NUMPAD0_VKEY = &H60
    NUMPAD1_VKEY = &H61
    NUMPAD2_VKEY = &H62
    NUMPAD3_VKEY = &H63
    NUMPAD4_VKEY = &H64
    NUMPAD5_VKEY = &H65
    NUMPAD6_VKEY = &H66
    NUMPAD7_VKEY = &H67
    NUMPAD8_VKEY = &H68
    NUMPAD9_VKEY = &H69
    PRINT_VKEY = &H2A
    RBUTTON_VKEY = &H2
    RCONTROL_VKEY = &HA3
    RETURN_VKEY = &HD
    RIGHT_VKEY = &H27
    RSHIFT_VKEY = &HA1
    SHIFT_VKEY = &H10
    SNAPSHOT_VKEY = &H2C
    SPACE_VKEY = &H20
    SUBTRACT_VKEY = &H6D
    TAB_VKEY = &H9
    UP_VKEY = &H26
    ZOOM_VKEY = &HFB
End Enum
'=========================
Type typRGB
 r As Long
 G As Long
 b As Long
End Type
'=========================
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'=========================
Type POINTAPI
   X As Long
   y As Long
End Type
'=========================
Public Type PAINTSTRUCT
        hdc As Long
        fErase As Long
        rcPaint As RECT
        fRestore As Long
        fIncUpdate As Long
        rgbReserved(32) As Byte
End Type
'=========================
Private aolframe As Long
Public HoldText As String
Public Listed As Boolean
Public Abort As Boolean
Public StopBust As Boolean
Public StopScroll As Boolean
'=========================

'_________________________________________
'                     Start Bas Code
'_________________________________________
'
Public Sub AddFonts2Combo(Combo As ComboBox)
'// add all the printer fonts on your PC into a combo box
Dim X As Long
 For X = 0 To Printer.FontCount - 1  'get number of printer fonts
  Combo.AddItem Printer.Fonts(X)   'add font(x) to combo
 Next X                                          'continue until all fonts are added
End Sub

Public Sub AIM_AddBuddyList(List As listbox)
'// add the aim buddylist to a listbox
On Error Resume Next
Dim oscarbuddylistwin As Long
Dim oscartabgroup As Long
Dim oscartree As Long
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)  'find buddylist
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString) 'find parent of list
oscartree& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tree", vbNullString)  'find list
 List.Clear  'clear your list so you can refresh it
  If oscartree& <> 0 Then
    LB2LB oscartree&, List   'uses my lb2lb to add the contents into your list
  End If
End Sub

Public Sub AIM_Addroom(List As listbox)
'// add the aim room to a listbox
On Error Resume Next
Dim aimchatwnd As Long
Dim oscartree As Long
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscartree& = FindWindowEx(aimchatwnd&, 0&, "_oscar_tree", vbNullString) 'find aim room's listbox
 List.Clear  'clear contents that are in listbox to refresh items
  If oscartree& <> 0 Then    'if oscartree is found then it continues
    LB2LB oscartree&, List  'add aim room to listbox
  End If
End Sub

Public Sub AIM_antiPunt()
'// This will scan the AIM chat room for a distort string, or the word punt
'// if either is found it will clear the chat screen and keep you from getting
'// any error message box that may come up as a result of the punters
On Error Resume Next
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim sHold As String
Dim txt As String
 aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
 wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
 ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
  txt$ = GetHalfText(ateclass&)
If InStr(1, LCase(txt$), LCase(">p<")) And InStr(1, LCase(txt$), LCase(">u<")) And InStr(1, LCase(txt$), LCase(">n<")) And InStr(1, LCase(txt$), LCase(">t<")) Then 'looks for the faded word punter
    SetText ateclass&, "Pünt Was Found In Chat"
End If
If InStr(1, LCase(txt$), LCase("punter")) Then 'looks for the word punter
    SetText ateclass&, "Pünt Was Found In Chat"
End If
If InStr(1, LCase(txt$), LCase(">e<")) And InStr(1, LCase(txt$), LCase(">r<")) And InStr(1, LCase(txt$), LCase(">r<")) And InStr(1, LCase(txt$), LCase(">o<")) Then 'looks for the word punter
    SetText ateclass&, "Error String Was Found In Chat"
End If
If InStr(1, LCase(txt$), Chr(9)) Then 'looks for distort
    SetText ateclass&, "Distort Was Found In Chat"
End If
If InStr(1, LCase(txt$), LCase(".clear")) Then 'looks for .clear in the room
    SetText ateclass&, "Manually Cleared Chat"
End If
End Sub

Public Sub AIM_ClearChat()
'// take a guess :P  this will clear the chat room text from an AIM chat room//
On Error Resume Next
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
 aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)  'chat window
 wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)  'parent area of text area
 ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString) 'finds chat view area
    SetText ateclass&, ""        'clears chat
End Sub

Public Function AIM_GetChat() As String
'// put AIM chat text into a textbox. it uses the function StripHTML
'// to strip out the html parts
On Error Resume Next
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim ChatText As String
 aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
 wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
 ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
    ChatText$ = GetAPIText(ateclass&)  'gets the chat text in raw format
    ChatText$ = StripHTML(ChatText$)   'strips all the html out of the ChatText$
 AIM_GetChat$ = ChatText$
End Function

Public Function AIM_GetHalfChat() As String
'// Gets half of the aim chat, makes AIM_LastLine work faster. Uses StripHtml
'// to strip out the html parts
On Error Resume Next
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim ChatText As String
 aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
 wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
 ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
    ChatText$ = GetHalfText(ateclass&)  'gets half the chat text in raw format
    ChatText$ = StripHTML(ChatText$)   'strips all the html out of the ChatText$
 AIM_GetHalfChat$ = ChatText$
End Function

Public Function AIM_GetIMsn() As String
'//  get the sender's SN from the IM
'//  text1 = aim_getimsn
Dim AIMim As Long
Dim Caption As String
Dim Dash As Long
AIMim& = FindWindow("aim_imessage", vbNullString)
    If AIMim& <> 0 Then
        Caption$ = GetAPIText(AIMim&)  'gets the im caption
        Dash& = InStr(1, Caption$, "-")    'finds the position of the hyphen -
        AIM_GetIMsn$ = Mid(Caption$, 1, Dash& - 2)  'gets to the left of the hyphen
    End If
End Function

Public Function AIM_GetIMtext() As String
'// get the message from an AIM instant message
'// used for a message machine
On Error Resume Next
Dim aimimessage As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim temptext As String
Dim StartAt As Long
Dim Text As String
Dim temptext2 As String
aimimessage& = FindWindow("aim_imessage", vbNullString)  'finds im
 wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString) 'parent to the text area
 ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)  'im text area
    temptext$ = GetAPIText(ateclass&)      'get the im text
    temptext2$ = StripHTML(temptext$) 'remove the html from the im
  StartAt& = InStr(1, temptext2$, ":")  'find colon
Text$ = Mid(temptext2, StartAt& + 1)  'mid out the SN and colon
AIM_GetIMtext = Trim(Text$)   'return text only
End Function

Public Function AIM_GetUser() As String
'// returns the AIM user's name
On Error Resume Next
Dim aimwin As Long
Dim TheCaption As String
Dim oscarbuddylistwin As Long
Dim thesn As String
Dim TheAppost As Long
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
 TheCaption$ = GetAPIText(oscarbuddylistwin&)  'get buddy list caption
  TheAppost& = InStr(1, TheCaption$, "'s Buddy List")  'find from the apostraphe
 thesn$ = Mid(TheCaption$, 1, TheAppost& - 1)  'mid out to the right of TheAppost&
AIM_GetUser$ = thesn$  'return SN
End Function

Public Sub AIM_Ignore(Person As String)
'// Ignore an AIM chat member by their SN, maybe you can use it for
'// AIM chat commands
On Error GoTo error_drat:
Dim C As Long
Dim numitems As Long
Dim sItemText As String * 255
Dim lstPlace As Long
Dim SN As Long
Dim aimchatwnd As Long
Dim oscariconbtn As Long
Dim oscartree As Long
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscartree& = FindWindowEx(aimchatwnd&, 0&, "_oscar_tree", vbNullString)
   numitems = SendMessageLong(oscartree&, LB_GETCOUNT, 0&, 0&)  'get the number of list items
   If numitems > 0 Then
      For C = 0 To numitems - 1
         lstPlace& = SendMessageByString(oscartree&, LB_SETCURSEL, C, 0)  'moves the highlighted list cursor
         SN& = SendMessageByString(oscartree&, LB_GETTEXT, C, ByVal sItemText)  'gets the text from lb item
          sItemText$ = Replace(sItemText$, Chr(32), "")  'removes spaces
          Person$ = Replace(Person$, Chr(32), "")  'remove spaces from sn to find
         sItemText$ = FixAPIString(sItemText$)  'fix any nulls
        If InStr(1, LCase(sItemText$), LCase(Person$)) <> 0 Then: ClickIt oscariconbtn& 'when found click the ignore button
      Next
   End If
error_drat:
End Sub

Public Sub AIM_KillAd()
'// kills the ad on the AIM buddy list
On Error Resume Next
Dim oscarbuddylistwin&
Dim wndateclass&
Dim ateclass&
 oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
 wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
    ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
    Win_CloseWin ateclass&
End Sub

Public Function AIM_LastLine() As String
'//  finds the last chat line in an aim chat room
Dim StringTemp As String
Dim Last As Long
Dim NewString As String
StringTemp$ = AIM_GetHalfChat()
Last& = InStrRev(StringTemp$, vbLf)
AIM_LastLine$ = Mid(StringTemp$, Last& + 1)
End Function

Public Function AIM_LastLineTxt() As String
'//  gets what was said on the last chat line
On Error Resume Next
Dim ChatString As String
Dim colon As Long
ChatString$ = AIM_LastLine()
colon& = InStr(1, ChatString$, ":")
AIM_LastLineTxt$ = Mid(ChatString$, colon& + 2)
End Function

Public Function AIM_LastLineSN() As String
'//  gets the SN from the last chat line in an aim chat room
On Error Resume Next
Dim ChatString As String
Dim colon As Long
ChatString$ = AIM_LastLine()
colon& = InStr(1, ChatString$, ":")
AIM_LastLineSN$ = Mid(ChatString$, 1, colon& - 1)
End Function

Public Sub AIM_RoomEnter(PersonSN As String, InviteMessage As String, RoomName As String)
'// invite someone to an AIM chat room or send it to the GetAIMsn to
'// use this as a room runner, to enter a room on AIM
'// AIMRoomEnter GetAIMsn, "enter room", "vb6"
'// that enters the AIM user into the chat room vb6
On Error Resume Next
Dim aimchatinvitesendwnd As Long
Dim edit As Long
Dim oscariconbtn As Long
Dim aimchatwnd As Long
Dim X As Long
'---------------
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
If aimchatwnd& = 0 Then
    RunMenu "_oscar_buddylistwin", "&People", "Send &Buddy Chat Invitation"
Else
    RunMenu "aim_chatwnd", "&People", "&Invite a Buddy..."
End If
'---------------
    For X = 0 To 100
        Sleep 15&
        aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
        If aimchatinvitesendwnd& <> 0 Then Exit For
    Next X
        aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
        edit& = FindWindowEx(aimchatinvitesendwnd&, 0&, "edit", vbNullString)
            SetText edit&, PersonSN$
        edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
            SetText edit&, InviteMessage$
        edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
            SetText edit&, RoomName$
        oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, 0&, "_oscar_iconbtn", vbNullString)
        oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
        oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
      Sleep 5&
    ClickIt oscariconbtn&
End Sub

Public Sub AIM_RoomLink(room As String)
'// link people to a room on AIM
AIM_SendRoom "aim:gochat?roomname=" & room$
End Sub

Public Sub AIM_SendIM(Person As String, SayWhat As String)
'//  send an instant message from Aol Instant Messenger
On Error Resume Next
RunMenu "_oscar_buddylistwin", "&People", "Send &Instant Message"
Dim aimimessage As Long
Dim oscarpersistantcombo As Long
Dim edit As Long
Dim ateclass As Long
Dim oscariconbtn As Long
Dim wndateclass As Long
  Pause 0.2
    aimimessage& = FindWindow("aim_imessage", vbNullString)
    oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
    edit& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
       SetText edit&, Person$
    Pause 0.1
        wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
        ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
        wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
           SetText wndateclass&, SayWhat$
    Pause 0.1
        oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
    Sleep 3&
        ClickIt oscariconbtn&
End Sub

Public Sub AIM_SendRoom(SayWhat As String)
'// Sends chat text to an AIM chat room
'// SendAIM "Your AIM ProgName"
On Error Resume Next
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim oscariconbtn As Long
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
    wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
        ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
            wndateclass& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
        SetText wndateclass&, SayWhat$
            oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
        oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
    oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
ClickIt oscariconbtn&
End Sub

Public Sub AOL_AddLB(wintitle As String, listbox As listbox)
'// add any aol child's listbox text, by the title of the child
'// example, addaol_lb "edit list ", list1
'// if you click on setup buddy list, Edit, and use that code
'// you can add everyone under that into list1
'// everyone and their mother uses this method of aol addlist, so whoever
'// made it gets the credit
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Dim room As Long
Dim Index As Long
Dim aolchild As Long
Dim aollistBox As Long
Dim aolthread As Long
Dim aolprocessthread As Long
' listbox.Clear
aolchild& = AOLChildByTitle(wintitle$)
aollistBox& = FindWindowEx(aolchild&, 0&, "_aol_listbox", vbNullString) 'find the listbox on the aolchild
    aolthread = GetWindowThreadProcessId(aollistBox&, AOLProcess)
     aolprocessthread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If aolprocessthread <> 0 Then
    For Index = 0 To SendMessage(aollistBox&, LB_GETCOUNT, 0, 0) - 1
     Person$ = String$(4, vbNullChar) 'create 4 nulls for a buffer
      ListItemHold = SendMessage(aollistBox&, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
       ListItemHold = ListItemHold + 24
        Call ReadProcessMemory(aolprocessthread, ListItemHold, Person$, 4, ReadBytes)  'gets the garbled list item
        Call CopyMemory(ListPersonHold, ByVal Person$, 4)  'copies to memory
       ListPersonHold = ListPersonHold + 6
     Person$ = String$(16, vbNullChar) 'creates new buffer to handle new item
    Call ReadProcessMemory(aolprocessthread, ListPersonHold, Person$, Len(Person$), ReadBytes) 'decypher garbled list item from memory and buffers
   Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1) 'strip the null characters from the name
  Call SendMessageByString(listbox.hwnd, LB_ADDSTRING, 0&, ByVal Person$) 'add the trimmed string to listbox
 Next Index
Call CloseHandle(aolprocessthread)  'close handle process
End If
End Sub

Public Sub AOL_AddMailBox(List As listbox)
'// Add your aol mailbox into a list, can be used for MMers
'// Servers or Anti-Spammers or whatever
On Error Resume Next
Dim MailBox As Long
Dim aoltabcontrol As Long
Dim aoltabpage As Long
Dim aoltree As Long
Do
 DoEvents
 MailBox& = AOLChildByTitle("'s Online Mailbox")
Loop Until MailBox& <> 0
 aoltabcontrol& = FindWindowEx(MailBox&, 0&, "_aol_tabcontrol", vbNullString)
   aoltabpage& = FindWindowEx(aoltabcontrol&, 0&, "_aol_tabpage", vbNullString)
 aoltree& = FindWindowEx(aoltabpage&, 0&, "_aol_tree", vbNullString)
LB2LB aoltree&, List
End Sub

Public Sub AOL_AddMemberDirectory(List As listbox)
'// adds aol4 member directory to listbox
'// i give whoever made the original 32 bit addroom partial credit with this
'// if you would like to search also then use something like aol4kw "aol://4950:0000010000|all:" & searchstring$
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Dim Index As Long
Dim aolchild As Long
Dim aollistBox As Long
Dim aolthread As Long
Dim theTab As Long
Dim aolprocessthread As Long
 List.Clear
Do
DoEvents
aolchild& = AOLChildByTitle("Member Directory Search Results")
Loop Until aolchild& <> 0
aollistBox& = FindWindowEx(aolchild&, 0&, "_aol_listbox", vbNullString)
    aolthread = GetWindowThreadProcessId(aollistBox&, AOLProcess)
     aolprocessthread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If aolprocessthread <> 0 Then
    For Index = 0 To SendMessage(aollistBox&, LB_GETCOUNT, 0, 0) - 1
     Person$ = String$(4, vbNullChar)
      ListItemHold = SendMessage(aollistBox&, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
       ListItemHold = ListItemHold + 24
        Call ReadProcessMemory(aolprocessthread, ListItemHold, Person$, 4, ReadBytes)
        Call CopyMemory(ListPersonHold, ByVal Person$, 4)
       ListPersonHold = ListPersonHold + 6
     Person$ = String$(16, vbNullChar)
    Call ReadProcessMemory(aolprocessthread, ListPersonHold, Person$, Len(Person$), ReadBytes)
   Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)  'trim all the buffer nulls
   theTab& = InStr(2, Person$, Chr(9))  'find the second tab
   Person$ = Mid(Person$, 2, theTab& - 1)  'get between the first tab and the second tab
  Call SendMessageByString(List.hwnd, LB_ADDSTRING, 0&, ByVal Person$)  'add only the screen name to the list box
 Next Index
Call CloseHandle(aolprocessthread)
End If
End Sub

Public Sub AOL_AddRoom(listbox As listbox)
'// adds aol4 chat room to a listbox
Dim room As Long
Dim Caption As String
   room& = AOL_FindRoom  'find room
    Caption$ = GetAPIText(room&)  'get room caption
   AOL_AddLB Caption$, listbox  'add room lb to list
End Sub

Public Sub AOL_Anti45nIdle()
'// kill the 45 minute idle, and the you have been idle screen
'// best used in a timer
On Error Resume Next
Dim aolframe As Long
Dim aolpalette As Long
Dim aolicon As Long
Dim Modal As Long
Dim aolbutton As Long
Dim aolstatic As Long
Dim Caption As String
aolframe& = FindWindow("aol frame25", vbNullString)
aolpalette& = FindWindow("_aol_palette", vbNullString)  'find the 45 minute window
If aolpalette& <> 0 Then
aolbutton& = FindWindowEx(aolpalette&, 0&, "_aol_button", "OK") 'aol3's button
aolicon& = FindWindowEx(aolpalette&, 0&, "_aol_icon", vbNullString) 'aol4's icon
    ClickIt aolbutton&  'click aol3's
    ClickIt aolicon&     'click aol4's
End If
aolframe& = FindWindow("aol frame25", vbNullString)
Modal& = FindWindow("_AOL_Modal", vbNullString) 'find any aol modal
aolstatic& = FindWindowEx(Modal&, 0&, "_aol_static", vbNullString)  'find the static on the window
Caption$ = GetAPIText(Modal&)  'get the window's caption
If aolstatic& <> 0 And Modal& <> 0 And Caption$ = "" Then  'if the static was found, and the modal was found, and the caption = "" or empty then continues
    aolicon& = FindWindowEx(Modal&, 0&, "_aol_icon", vbNullString) 'finds the aol icon on the modal
    ClickIt aolicon&  'clicks it to close the idle window
End If
End Sub

Public Sub AOL_BuddySetup()
'//  click the buddy list setup button so you can make/edit
'//  your buddy list preferences
On Error Resume Next
Dim aolchild As Long
Dim aolicon As Long
  aolchild& = AOLChildByTitle("Buddy List Window")
    aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)  'locate
    aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString) 'IM
    aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString) 'Setup
  ClickIt aolicon&  'click setup
End Sub

Public Function AOL_ChatView() As Long
'// you can use this with getapitext to get the chat text or window captions
On Error Resume Next
Dim Chat As Long
Dim Rich As Long
 Chat& = AOL_FindRoom()  'find chat window
    Rich& = FindWindowEx(Chat&, 0&, "RICHCNTL", vbNullString)  'chat text area
 AOL_ChatView = Rich&
End Function

Public Sub AOL_CheckIMs(Who2Check As String)
'//  Check to see if someone's instant messages are on or off
'//  This sends to the room, if you want to have it go into a label then
'//  You will have to add something like Txt as String, up in the syntax then
'//  Change the AOL_SendRoom to Txt = sReult$
On Error Resume Next
Dim theim As Long
Dim aolicon As Long
Dim msg As Long
Dim Button As Long
Dim msgStatic As Long
Dim sResult As String
    AOL4KW "aol://9293:" & Who2Check$
Do
    DoEvents
    theim& = AOLChildByTitle("Send Instant Message")
Loop Until theim& <> 0
        aolicon& = FindWindowEx(theim&, 0&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)  'Send
        aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)  'Available
ClickIt aolicon&  'click available
Win_CloseWin theim&  'close im window
Do      'begin loop looking for message box
  DoEvents
   msg& = FindWindow("#32770", vbNullString)  'message box
   msgStatic& = FindWindowEx(msg&, 0&, "static", vbNullString)  'the ! static picture
  msgStatic& = FindWindowEx(msg&, msgStatic&, "static", vbNullString)  'the information area
Loop Until msg& <> 0 And msgStatic& <> 0  'loop until it finds both the message box and the static area
 sResult$ = GetAPIText(msgStatic&)             'get the text from the information area
AOL_SendRoom "]yourasciihere[ " & sResult$
'//  if you don't want to have it send into a room
'//  then you can add it into a list, textbox or label
 AOL_Wait4OK   'close the message box
End Sub

Public Sub AOL_CreateProfile(Name As String, City As String, Birthday As String, Married As String, Hobbies As String, Computers As String, Occupation As String, Quote As String)
'//  Just insert the data you want your profile to say for each string then
'//  Click to start it, this is super fast :)
Dim EditProfile As Long
Dim YourName As Long
Dim CityState As Long
Dim Birthdate As Long
Dim MaritalStat As Long
Dim Hobby As Long
Dim CompUsed As Long
Dim Occup As Long
Dim PersQuote As Long
Dim MemDirectory As Long
Dim MyProfile As Long
Dim Update As Long
AOL4KW "aol://1722:member directory"
    Do
      DoEvents
      MemDirectory& = AOLChildByTitle("Member Directory")
    Loop Until MemDirectory& <> 0
MyProfile& = FindWindowEx(MemDirectory&, 0&, "_aol_icon", vbNullString)
ClickIt MyProfile&
Win_CloseWin MemDirectory&
    Do
      DoEvents
      EditProfile& = AOLChildByTitle("Edit Your Online Profile")
        YourName& = FindWindowEx(EditProfile&, 0&, "_aol_edit", vbNullString)
        CityState& = FindWindowEx(EditProfile&, YourName&, "_aol_edit", vbNullString)
        Birthdate& = FindWindowEx(EditProfile&, CityState&, "_aol_edit", vbNullString)
        MaritalStat& = FindWindowEx(EditProfile&, Birthdate&, "_aol_edit", vbNullString)
        Hobby& = FindWindowEx(EditProfile&, MaritalStat&, "_aol_edit", vbNullString)
        CompUsed& = FindWindowEx(EditProfile&, Hobby&, "_aol_edit", vbNullString)
        Occup& = FindWindowEx(EditProfile&, CompUsed&, "_aol_edit", vbNullString)
        PersQuote& = FindWindowEx(EditProfile&, Occup&, "_aol_edit", vbNullString)
    Loop Until EditProfile& <> 0 And PersQuote& <> 0
SetText YourName&, Name$
SetText CityState&, City$
SetText Birthdate&, Birthday$
SetText MaritalStat&, Married$
SetText Hobby&, Hobbies$
SetText CompUsed&, Computers$
SetText Occup&, Occupation$
SetText PersQuote&, Quote$
    Update& = FindWindowEx(EditProfile&, 0&, "_aol_icon", vbNullString)
  ClickIt Update&
AOL_Wait4OK
End Sub

Public Sub AOL_Find_a_Chat(ChatName As String)
'//  must be in a chat room to use this, put in a room name you
'//  want to look for, such as Find_a_Chat "Scrambler"  will search
'// aol member rooms for any rooms named Scrambler
On Error Resume Next
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aolicon As Long
Dim AOLIcon2 As Long
Dim aolchild2 As Long
Dim room As Long
Dim aollistBox As Long
Dim aoledit As Long
If AOL_FindRoom() <> 0 Then
room& = AOL_FindRoom()
    aolicon& = FindWindowEx(room&, 0&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(room&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(room&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(room&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(room&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(room&, aolicon&, "_aol_icon", vbNullString)  'Find A Chat icon on chatroom
  ClickIt aolicon&
 Pause 2
Do
  DoEvents
  aolchild2& = AOLChildByTitle("Find a Chat")   'loop until form appears
Loop Until aolchild2& <> 0
      Pause 2  'allow time for the member room list to load
            aollistBox& = FindWindowEx(aolchild2&, 0&, "_aol_listbox", vbNullString)                 'room area
            aollistBox& = FindWindowEx(aolchild2&, aollistBox&, "_aol_listbox", vbNullString)  'room names
            Win_CloseWin aollistBox&  'close the room list
     Pause 0.4
                AOLIcon2& = FindWindowEx(aolchild2&, 0&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString)
                AOLIcon2& = FindWindowEx(aolchild2&, AOLIcon2&, "_aol_icon", vbNullString) 'Search AOL chat icon
    ClickIt AOLIcon2&
Pause 0.5
  Do
    Sleep 0&
    aolchild& = AOLChildByTitle("Search Member Chats")  'loop until input form is found
  Loop Until aolchild& <> 0
    aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)  'finds the text area to type name
     SetText aoledit&, ChatName$  'set your room (string) to find
    EnterKey aoledit&  'presses enter to finish search
  End If
End Sub

Public Sub AOL_FindBuddyPR(Room2Check As listbox, addroom_list As listbox, BuddysSN As String)
'// Ok the Room2Check list is a list of possible rooms that your friend might hang out in
'// addroom_list is the listbox that the room will be added to then searched
'// buddysSN is the screen name to search for (text box most likely)
Dim Down As Long
Dim down2 As Long
Dim room As Long
For Down = 0 To Room2Check.ListCount - 1
   AOL4KW "aol://2719:2-2-" & Room2Check.List(Down)   'goes down the lb one by one to other rooms
    Pause 2
     room& = AOL_FindRoom  'find chat room
      If room& <> 0 Then
       Call AOL_AddRoom(addroom_list)  'add room to other list
        Pause 1
         For down2 = 0 To addroom_list.ListCount - 1   'start search for your friend's SN
           If addroom_list.List(down2) = BuddysSN$ Then
             If DoesFileExist("C:\WINDOWS\MEDIA\TADA.WAV") = True Then
                PlayWav "C:\WINDOWS\MEDIA\TADA.WAV"  'if you have this wav it plays when found
             Else
                Beep     'if not program will beep
             End If
               MsgBox BuddysSN$ & " has been found!"
             GoTo FOUND:
           End If
        Next down2  'continue searching for SN
    End If
Next Down         'go on to next room in room2check list
FOUND:
End Sub

Public Function AOL_FindRoom() As Long
'// Locate the AOL chat room, if its opened it will return a number
'// greater than 0, works great
On Error Resume Next
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aollistBox As Long
Dim aolicon As Long
Dim richcntl As Long
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = GetWindow(GetWindow(mdiclient&, GW_CHILD), GW_HWNDFIRST)  'find the first aol child
Do
    aollistBox& = FindWindowEx(aolchild&, 0&, "_aol_listbox", vbNullString)
    aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
    richcntl& = FindWindowEx(aolchild&, 0&, "richcntl", vbNullString)
    richcntl& = FindWindowEx(aolchild&, richcntl&, "richcntl", vbNullString)
    If (aollistBox& <> 0) And (richcntl& <> 0) And (richcntl& <> 0) And (aolicon& <> 0) Then Exit Do
    aolchild& = GetWindow(aolchild&, GW_HWNDNEXT) 'if only some of those items were found, it will go onto next aol child
Loop Until aolchild& = 0      'loops until you search through all aol child's
AOL_FindRoom = aolchild& '0 means room not found > 0 means its found
End Function

Public Sub AOL_GetProfile(Person As String)
'//  Thanks to kev for showing me in the right direction on this
Dim aolframe As Long
Dim aoltoolbar As Long
Dim aolicon As Long
Dim getprofile As Long
Dim aoledit As Long
aolframe& = FindWindow("aol frame25", vbNullString)
 aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
 aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
    aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString) 'read
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'write
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'mail center
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'print
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'my files
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'my aol
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'favorites
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'internet
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'channels
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'people
  SendVKey aolicon&, SPACE_VKEY  'push the space bar on the people icon
Call SendMessageByNum(aolicon&, WM_CHAR, 103, 0) 'send chr(103) "g" to the icon
  TimeOut 200  'slight timeout before starting the do/loop
    Do
      DoEvents
        getprofile& = AOLChildByTitle("Get a Member's Profile") 'loop until the child is found
    Loop Until getprofile& <> 0
  aoledit& = FindWindowEx(getprofile&, 0&, "_aol_edit", vbNullString) 'the text area on the getprofile child
SetText aoledit&, Person$  'sets the person's SN into the text area
  EnterKey aoledit&  'pushes enter
Win_CloseWin getprofile&  'close the getprofile window
End Sub

Public Function AOL_GetProfileText(SN2get As String) As String
'//  Use this if you want to get the member's profile and put it into a textbox
'//  Or to scroll in the room,   EXAMPLE:
'//  Dim txt As String
'//  txt$ = AOL_GetProfileText("SteveCase")
'//  SendMacro txt$
Dim msg As Long
Dim profile As Long
Dim profileview As Long
   Call AOL_GetProfile(SN2get$)     'open profile
Pause 1
msg& = FindWindow("#32770", "America Online")  'if no profile then skip to end
If msg& <> 0 Then GoTo No_Profile:
     Do
       DoEvents
        profile& = AOLChildByTitle("Member Profile")  'find profile window
    Loop Until profile& <> 0
Pause 1
    profileview& = FindWindowEx(profile&, 0&, "_aol_view", vbNullString)  'profiles text area
    AOL_GetProfileText = GetAPIText(profileview&)  'get the profile
Win_CloseWin profile&: Exit Function  'close profile and exit function
No_Profile: AOL_Wait4OK: AOL_GetProfileText = "no profile for " & SN2get$: Exit Function
End Function

Public Function AOL_GetUser() As String
'// returns the AOL user's name
On Error Resume Next
Dim welcome As Long
Dim TheCaption As String
Dim TheComma As Long
Dim NewCaption As String
Dim TheExlaim As Long
Dim thesn As String
welcome& = AOLChildByTitle("Welcome, ") 'find welcome window
 TheCaption$ = GetAPIText(welcome&)       'get its caption
  TheCaption$ = Trim(TheCaption$)             'trims spaces from caption if any
   TheComma& = InStr(1, TheCaption$, ", ") 'finds the comma in caption
   NewCaption$ = Mid(TheCaption$, TheComma& + 2) 'get to the right of the comma
  TheExlaim& = InStr(1, NewCaption$, "!") 'finds the exclaimation mark
 thesn$ = Mid(NewCaption$, 1, TheExlaim& - 1)  'get to the left of the exclaim
AOL_GetUser$ = Trim(thesn$)  'final return is your SN
End Function

Public Sub AOL_ignore(SNtoIgnore As String)
'// this is to be used in a room, can be made into an//
'// auto ignorer if you use dos' chat ocx or any other aol4 chat ocx//
'// i give whoever made the original 32 bit addroom partial credit with this//
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Dim Place As Long
Dim room As Long
Dim Index As Long
Dim aolframe As Long
Dim aolchild As Long
Dim aolcheckbox As Long
Dim aollistBox As Long
Dim aolthread As Long
Dim popoptions As Long
Dim who As String
Dim aolprocessthread As Long
Dim mdiclient As Long
Dim getit As Long
room& = AOL_FindRoom
aollistBox& = FindWindowEx(room&, 0&, "_aol_listbox", vbNullString)
    aolthread = GetWindowThreadProcessId(aollistBox&, AOLProcess)
     aolprocessthread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
  If aolprocessthread <> 0 Then
    For Index = 0 To SendMessage(aollistBox&, LB_GETCOUNT, 0, 0) - 1
     Place& = SendMessageByString(aollistBox&, LB_SETCURSEL, Index, 0)
      Person$ = String$(4, vbNullChar)
       ListItemHold = SendMessage(aollistBox&, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
       ListItemHold = ListItemHold + 24
        Call ReadProcessMemory(aolprocessthread, ListItemHold, Person$, 4, ReadBytes)
        Call CopyMemory(ListPersonHold, ByVal Person$, 4)
       ListPersonHold = ListPersonHold + 6
     Person$ = String$(16, vbNullChar)
    Call ReadProcessMemory(aolprocessthread, ListPersonHold, Person$, Len(Person$), ReadBytes)
   Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)  'remove nulls
   who$ = Trim(Person$) 'trim <- spaces ->
   Person$ = LCase(Replace(Person$, " ", "")) 'remove spaces list string
   SNtoIgnore$ = LCase(Replace(SNtoIgnore$, " ", "")) 'remove spaces from sn string
  If Person$ = SNtoIgnore$ Then
  Call SendMessageByNum(aollistBox&, WM_LBUTTONDBLCLK, Index, 0&) 'double clicks list index
    Pause 1
     popoptions& = AOLChildByTitle(who$)
        aolcheckbox& = FindWindowEx(popoptions&, 0&, "_aol_checkbox", vbNullString) 'find check
        getit& = SendMessageByNum(aolcheckbox&, BM_GETCHECK, 0&, 0&)   'get check value
     ClickIt aolcheckbox&  'click check box
   Win_CloseWin popoptions&: GoTo finished
  End If
  Next Index
finished:
Call CloseHandle(aolprocessthread)
End If
End Sub

Public Function AOL_im_sn() As String
'//  gets the aol screen name from the IM
On Error Resume Next
Dim theim As Long
Dim colon As Long
Dim Caption As String
Dim snText As String
 theim& = AOLChildByTitle("Instant Message From:")
If theim& <> 0 Then
   Caption$ = GetAPIText(theim&)    'get the caption of IM window
   colon& = InStr(1, Caption$, ":")    'find the colon
 snText$ = Mid(Caption$, colon& + 2)  'get to the right of colon
AOL_im_sn$ = Trim(snText$) 'trim any spaces
End If
End Function

Public Function AOL_im_txt() As String
'// gets the message left in an aol IM
On Error Resume Next
Dim theim As Long
Dim WholeText As String
Dim colon As Long
Dim richcntl As Long
Dim txtText As String
 theim& = AOLChildByTitle("Instant Message From:")
If theim& <> 0 Then
  richcntl& = FindWindowEx(theim&, 0&, "richcntl", vbNullString)
     WholeText$ = GetAPIText(richcntl&)
        colon& = InStr(1, WholeText$, ":")
        txtText$ = Mid(WholeText$, colon& + 3)  'gets everything after the colon
  AOL_im_txt = Trim(txtText$)
End If
End Function

Public Sub AOL_IMsOFF()
'//  Yeah
On Error Resume Next
Call AOL_SendIM("$IM_OFF", "Turning IM's Off")
End Sub

Public Sub AOL_IMsON()
'//  Bet you can't guess what this does
On Error Resume Next
Call AOL_SendIM("$IM_ON", "Turning IM's On")
End Sub

Public Sub AOL_KillWait()
'// kill the hourglass
On Error Resume Next
Dim aolframe As Long
Dim aolmodal As Long
Dim aolstatic As Long
Dim aolicon As Long
Call RunMenu("aol frame25", "&Help", "&About America Online")  'popup the about aol screen
 Pause 1.2
   aolframe& = FindWindow("aol frame25", vbNullString)
    aolmodal& = FindWindow("_aol_modal", vbNullString)
     aolstatic& = FindWindowEx(aolmodal&, 0&, "_aol_static", vbNullString)
     aolstatic& = FindWindowEx(aolmodal&, aolstatic&, "_aol_static", vbNullString) 'make sure its found
    If aolstatic& = 0 Then Exit Sub
  aolicon& = FindWindowEx(aolmodal&, 0&, "_aol_icon", vbNullString)  'find OK icon
ClickIt aolicon&  'click aolicon
End Sub

Public Function AOL_LastLine() As String
'//  I looked at another bas and they had like 15 lines of code, this takes 3 lines ;)
Dim StringTemp As String
Dim Last As Long
Dim NewString As String
StringTemp$ = GetHalfText(AOL_ChatView())
Last& = InStrRev(StringTemp$, vbCr)
AOL_LastLine$ = Mid(StringTemp$, Last& + 1)
End Function

Public Function AOL_LastLineTxt() As String
'//  gets what was said from the last chat line in an aol chat room
On Error Resume Next
Dim ChatString As String
Dim colon As Long
ChatString$ = AOL_LastLine()
colon& = InStr(1, ChatString$, ":")
AOL_LastLineTxt$ = Mid(ChatString$, colon& + 3)
End Function

Public Function AOL_LastLineSN() As String
'//  gets the SN from the last chat line in an aol chat room
On Error Resume Next
Dim ChatString As String
Dim colon As Long
ChatString$ = AOL_LastLine()
colon& = InStr(1, ChatString$, ":")
AOL_LastLineSN$ = Mid(ChatString$, 1, colon& - 1)
End Function

Public Sub AOL_LocateSN(List As listbox, OtherList As listbox)
'// locate members online (just like PowerTools does)
'// goes down each person's SN in a list and locates them
'// then puts their location into another listbox
On Error Resume Next
Dim X As Long
Dim aolframe As Long
Dim mdiclient As Long
Dim aolstatic As Long
Dim Location As Long
Dim msg As Long
Dim Button As Long
OtherList.Clear
For X = 0 To List.ListCount - 1
    Sleep 0&
    AOL4KW "aol://3548:" & List.List(X)
        Do
            On Error Resume Next
            Sleep 0&
                aolframe& = FindWindow("aol frame25", vbNullString)
                mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
                Location& = AOLChildByTitle("Locate " & List.List(X))
                msg& = FindWindow("#32770", "America Online")
            Sleep 0&
          If Location& <> 0 Then: Exit Do
          If msg& <> 0 Then: Exit Do
        Loop
      If msg& <> 0 Then
        Sleep 0&
        OtherList.AddItem List.List(X) + " is not signed on."
        Win_CloseWin msg&
      End If
If Location& <> 0 Then
    Sleep 0&
    aolstatic& = FindWindowEx(Location&, 0&, "_aol_static", vbNullString)
        OtherList.AddItem "@" & GetAPIText(aolstatic&)
        Sleep 0&
        Win_CloseWin Location&
End If
    OtherList.ListIndex = OtherList.ListCount - 1
Next X
Win_CloseWin Location&
End Sub

Public Function AOL_MailCount() As Long
'//  open up your aol mailbox and count the contents
Dim MailBox As Long
Dim aoltabcontrol&
Dim aoltabpage&
Dim aoltree&
AOL_MailByIcon 0&
Pause 2.5
    MailBox& = AOLChildByTitle("'s Online Mailbox")
    aoltabcontrol& = FindWindowEx(MailBox&, 0&, "_aol_tabcontrol", vbNullString)
    aoltabpage& = FindWindowEx(aoltabcontrol&, 0&, "_aol_tabpage", vbNullString)
    aoltree& = FindWindowEx(aoltabpage&, 0&, "_aol_tree", vbNullString)
      Pause 1.5
AOL_MailCount& = SendMessageLong(aoltree&, LB_GETCOUNT, 0&, 0&) 'how to count LB's contents
End Function

Public Sub AOL_MailByIcon(num As Long)
'//  select the case number you want to click on
'//  AOL_MailByIcon 0& opens up your mail box
'//  AOL_MailByIcon 1& opens up a blank email
On Error Resume Next
Dim aolframe As Long
Dim aoltoolbar As Long
Dim aolicon As Long
Dim MailBox As Long
Dim WriteMail As Long
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
MailBox& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
WriteMail& = FindWindowEx(aoltoolbar&, MailBox&, "_aol_icon", vbNullString)
    Select Case num
        Case 0
            '//  AOL_MailByIcon 0&  opens up your mailbox
            ClickIt MailBox&
        Case 1
            '//  AOL_MailByIcon 1& opens up a blank email to compose
            ClickIt WriteMail&
    End Select
End Sub

Public Sub AOL_MailOpenOld()
Dim aolframe As Long
Dim aoltoolbar As Long
Dim aolicon As Long
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
    aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)         'Read
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'Write
    aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString) 'Mail Center
  SendVKey aolicon&, SPACE_VKEY   'push space bar on the icon
Call SendMessageByNum(aolicon&, WM_CHAR, 111, 0)  'send "o" to open old email
End Sub

Public Sub AOL_RoomBust(RoomName As String)
'// fairly quick room buster for aol4, if you want to make it for
'// member rooms too, not just a private room, you'll have to
'// change the aol:// keywords
''// Private" = "aol://2719:2-2-"***"Arts & Entertainment" = "aol://2719:62-2-"***"Special Interests" = "aol://2719:67-2-"***"Hong Kong" = "aol://2719:77-2-"***"Town Square" = "aol://2719:61-2-"***"Friends" = "aol://2719:74-2-"***"Life" = "aol://2719:63-2-"***"News Sports & Finance" = "aol://2719:64-2-"***"Places" = "aol://2719:65-2-"***"Romance" = "aol://2719:66-2-"***"UK" = "aol://2719:69-2-"***"France" = "aol://2719:70-2-"***"Canada" = "aol://2719:71-2-"***"Japan" = "aol://2719:73-2-"
'// to stop the bust, make a button and put StopBust = True   and thats it
On Error Resume Next
Dim room As Long
Dim msg As Long
room& = AOL_FindRoom()
If room& <> 0 Then
    Win_CloseWin room&
End If
StopBust = False
Do
  DoEvents
    msg& = FindWindow("#32770", "America Online")
        If msg& <> 0 Then
            AOL_Wait4OK
        End If
    AOL4KW "aol://2719:2-2-" & RoomName$
    Pause 1 'for a faster room bust comment this pause out
    room& = AOL_FindRoom()
    If room& <> 0 Then StopBust = True: Exit Do
Loop Until StopBust = True
End Sub

Public Sub AOL_SendIM(Person As String, SayWhat As String)
'//  Send an instant message to someone on aol
'//  to turn off IM's you would put SendIM "$IM_OFF", "turning im's off"
'//  to turn them back on you put SendIM "$IM_ON", "im's are back on"
On Error Resume Next
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aolicon As Long
Dim aoledit As Long
Dim richcntl As Long
Dim theim As Long
Dim TheIM2 As Long
Dim msg As Long
Call AOL4KW("aol://9293:" & Person$) 'kw to open im
  Do
     Sleep 0&
      theim& = AOLChildByTitle("Send Instant Message")
  Loop Until theim& <> 0
   richcntl& = FindWindowEx(theim&, 0&, "richcntl", vbNullString)
    aolicon& = FindWindowEx(theim&, 0&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
    aolicon& = FindWindowEx(theim&, aolicon&, "_aol_icon", vbNullString)
  Sleep 1&
        SetText richcntl&, SayWhat$  'set text to im
  Sleep 1&
ClickIt aolicon&
Pause 1
    msg& = FindWindow("#32770", "America Online") 'check for msgbox
        If msg& <> 0 Then
            AOL_Wait4OK
            Win_CloseWin theim&
        End If
End Sub

Public Sub AOL_SendMail(Person As String, Subject As String, Body As String)
'// Send Email to someone
'// SendMail "vbSiR@juno.com", "hi", "i'm using your bas :D"
On Error Resume Next
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aoledit As Long
Dim richcntl As Long
Dim aolicon As Long
Dim aolicon1 As Long
Dim AOLIcon2 As Long
Call AOL_MailByIcon(1)
 Do
   Sleep 0&
    aolchild& = AOLChildByTitle("Write Mail")
 Loop Until aolchild& <> 0
    aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
       Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, Person$)
        Sleep 1&
         aoledit& = FindWindowEx(aolchild&, aoledit&, "_aol_edit", vbNullString)
          aoledit& = FindWindowEx(aolchild&, aoledit&, "_aol_edit", vbNullString)
          Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, Subject$)
         Sleep 1&
        richcntl& = FindWindowEx(aolchild&, 0&, "richcntl", vbNullString)
       Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, Body$)
   Pause 0.1
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString) 'i'm getting tired of commenting
Pause 0.5
ClickIt aolicon&
End Sub

Public Sub AOL_SendRoom(What As String)
'// Sends chat text to the room
'// AOL_SendRoom "Your ProgName"
On Error Resume Next
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim richcntl As Long
Dim room As Long
If AOL_FindRoom <> 0 And Len(What$) <> 0 Then
room& = AOL_FindRoom
    richcntl& = FindWindowEx(room&, 0&, "richcntl", vbNullString)
    richcntl& = FindWindowEx(room&, richcntl&, "richcntl", vbNullString)
        Sleep 0&
            Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, What$)
       EnterKey richcntl&
End If
End Sub

Public Sub AOL_SignOnAs(SNtoSignOn As String, Password As String)
'// i don't really know why i made this, since you can change names without//
'// signing off, oh well, what it does is searches the combo box for the SNtoSignOn//
'// and once found it places the PassWord and signs you on.//
'// i give whoever made the original 32 bit addroom partial credit with this//
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Dim room As Long
Dim Index As Long
Dim aolhandle As Long
Dim aolthread As Long
Dim aolprocessthread As Long
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim aolcombobox As Long
Dim aoledit As Long
Dim Place As Long
  aolframe& = FindWindow("aol frame25", vbNullString)
  mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
  aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
    aolhandle& = FindWindowEx(aolchild&, 0&, "_aol_combobox", vbNullString)
    aolthread& = GetWindowThreadProcessId(aolhandle, AOLProcess)
    aolprocessthread& = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
   If aolprocessthread& <> 0 Then
    For Index = 0 To SendMessage(aolhandle&, CB_GETCOUNT, 0, 0) - 1
    Place& = SendMessageByString(aolhandle&, CB_SETCURSEL, Index, 0)
     Person$ = String$(4, vbNullChar)
      ListItemHold& = SendMessage(aolhandle&, CB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
       ListItemHold& = ListItemHold& + 24
        Call ReadProcessMemory(aolprocessthread&, ListItemHold&, Person$, 4, ReadBytes&)
        Call CopyMemory(ListPersonHold&, ByVal Person$, 4)
       ListPersonHold& = ListPersonHold& + 6
     Person$ = String$(16, vbNullChar)
    Call ReadProcessMemory(aolprocessthread&, ListPersonHold&, Person$, Len(Person$), ReadBytes&)
   Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
   Person$ = LCase(Replace(Person$, " ", ""))
   SNtoSignOn$ = LCase(Replace(SNtoSignOn$, " ", ""))
  If Person$ = SNtoSignOn$ Then GoTo FOUND:
Next Index
GoTo Not_Found:
FOUND:
Call CloseHandle(aolprocessthread)
aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
SetText aoledit&, Password$
EnterKey aoledit&
End If
Not_Found:
End Sub

Public Sub AOL_UpChatOff()
'// Lets you turn off the upchat for any reason u may have
Dim Modal As Long
Dim Dummy As Long
Modal& = FindWindow("_AOL_MODAL", vbNullString)
Dummy& = showwindow(Modal&, SW_SHOWNORMAL)
End Sub

Public Sub AOL_UpChatOn()
'// lets you upload and use aol as you normally would
Dim Modal As Long
Dim Dummy As Long
Modal& = FindWindow("_AOL_MODAL", vbNullString)
Dummy& = showwindow(Modal&, SW_ShowMinimized)
Dummy& = showwindow(Modal&, SW_Hide)
End Sub

Public Function AOL_Version4() As Boolean
'// Check the users aol version, if they are on aol4 it returns True
'// otherwise it returns false
On Error Resume Next
Dim aolframe As Long
Dim aoltoolbar As Long
Dim aolglyph As Long
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolglyph& = FindWindowEx(aoltoolbar&, 0&, "_aol_glyph", vbNullString)
If aolglyph& <> 0 And aoltoolbar& <> 0 Then
    AOL_Version4 = True
Else
    AOL_Version4 = False
End If
End Function

Public Sub AOL_Wait4OK()
'// waits for a message box then kills it
On Error Resume Next
Dim msg As Long
  Do
   DoEvents
    msg& = FindWindow("#32770", "America Online")
  Loop Until msg& <> 0
    Win_CloseWin msg&
End Sub

Public Sub AOL4KW(TheKW As String)
'//  send an AOL4 keyword
On Error Resume Next
Dim aolframe As Long
Dim aoltoolbar As Long
Dim aolcombobox As Long
Dim edit As Long
Dim aolicon As Long
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolcombobox& = FindWindowEx(aoltoolbar&, 0&, "_aol_combobox", vbNullString)
edit& = FindWindowEx(aolcombobox&, 0&, "edit", vbNullString)
    SetText edit&, TheKW$
        Call SendMessageByNum(edit&, WM_CHAR, 32, 0&)
    EnterKey edit&
End Sub

Public Function AOLChildByTitle(title As String)
'// Finds any aolchild by its title, doesn't have to be the
'// windows full title, it can be partial title
'// this also works great EXAMPLE:
'// child& = AolChildByTitle("buddy list")     will search through all of the
'// aol child's until it finds that window..... if you are not looking for an aol
'// child then use findchildbytitle also found in this bas
On Error Resume Next
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim childtitle As String
Dim FoundIt As Long
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = GetWindow(GetWindow(mdiclient&, GW_CHILD), GW_HWNDFIRST)
Do
  DoEvents
    childtitle$ = GetAPIText(aolchild&)
    Sleep 2&
    FoundIt& = InStr(UCase(Replace(childtitle$, " ", "")), UCase(Replace(title$, " ", "")))
        If FoundIt& <> 0 Then Exit Do
    aolchild& = GetWindow(aolchild&, GW_HWNDNEXT)
Loop Until aolchild& = 0
AOLChildByTitle = aolchild&
End Function


Public Sub AOLwwwLink(Address As String, LinkText As String)
'// send a web link into the aol4 chat room
AOL_SendRoom "< a href=" & Chr(34) & Address$ & Chr(34) & ">" & LinkText$ & "</a>"
End Sub

Public Sub CD_Controls(returnstring As String)
'//  CD_Controls "open" will open the door
'//  CD_Controls "close" will close the door
On Error Resume Next
  Select Case LCase(returnstring$)
    Case "open"
        Call MciSendString("set CDAudio door open", 0, 127&, 0&)
    Case "close"
        Call MciSendString("set CDAudio door closed", 0, 127&, 0&)
  End Select
End Sub

Public Function ChrNumber(Text As String) As String
'// Quickly convert any string to its Chr(#) value
'// good for encrypting your programs name so it won't be hexed out
On Error Resume Next
Dim Letters As Long
Dim out As String
For Letters = 1 To Len(Text$)
    out$ = out$ + "Chr(" + CStr(Asc(Mid(Text$, Letters, 1))) + ") & "
Next Letters
 out$ = Trim(out$)
 out$ = Mid(out$, 1, Len(out$) - 2)
ChrNumber = out$
End Function

Public Sub ClickIt(THing As Long)
'// clicks a button or icon that you may need
DoEvents
Call SendMessage(THing&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(THing&, WM_LBUTTONUP, 0&, 0&)
DoEvents
End Sub

Public Sub Clip_Copy(Text As textbox)
'//  Copy text onto Clipboard.
Text.SelStart = 0
Text.SelLength = Len(Text)
'//  Copy all text onto Clipboard.
 Clipboard.SetText Text.SelText
End Sub

Public Sub Clip_Cut(Text As textbox)
'//  Copy text onto Clipboard.
Text.SelStart = 0
Text.SelLength = Len(Text)
Clipboard.SetText Text.SelText
'//  Delete selected text.
Text.SelText = ""
End Sub

Public Sub Clip_Paste(Text As textbox)
'//  Put Clipboard text in text box.
  Text.SelText = Clipboard.GetText()
End Sub

Public Sub Clip_Purge()
'// Purge or empty the contents on the clipboard
On Error Resume Next
Dim Dummy As Long
Dummy& = EmptyClipboard()
End Sub

Public Sub Clip_SelectAll(Text As textbox)
'//  selects all the text in a textbox
  Text.SelStart = 0
  Text.SelLength = Len(Text.Text)
End Sub

Public Sub CtrlAltDel(Number As Long)
'// enables or disables control alt delete, CtrlAltDel 0   disables it and CtrlAltDel 1 enables it again//
On Error Resume Next
Dim TheReturn As Long
Dim TorF As Boolean
 Select Case Number
    Case 0  'disables cntrl alt del
        TheReturn& = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, TorF, 0)
    Case 1  're-enables cntrl alt del
        TheReturn& = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, TorF, 0)
 End Select
End Sub

Public Sub DeltFile(PathnFile As String)
'// The api way of deleting a file, or you can use Kill
On Error Resume Next
Dim Dummy As Long
 If DoesFileExist(PathnFile$) = True Then
    Dummy& = DeleteFile(PathnFile$)
    MsgBox "The File [ " & PathnFile$ & " ] has been deleted"
 Else
    MsgBox "Sorry that file was not found."
 End If
End Sub

Public Function DoesFileExist(PathnFile As String) As Boolean
'// Check and see if a file exists on the computer, if it does it returns True, otherwise returns False
On Error Resume Next
If Len(Dir(PathnFile$)) >= 1 Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Sub EnterKey(Winder As Long)
'//  send an Enter key to a window or aoledit box
Call SendMessageByNum(Winder, WM_CHAR, 13, 0&)
End Sub

Public Sub ExitProg()
'// Lets the user choose if they really want to exit or if it was just an accident
On Error Resume Next
Dim mbResult As VbMsgBoxResult
mbResult = MsgBox("Are you sure you want to exit the program?", vbYesNo)
If mbResult = vbYes Then
    Form_UnloadAll
End If
End Sub

Public Function FindChildByTitle(Parent As Long, title As String)
'// Finds any child by its parent & title, doesn't have to be the
'// windows full title, it can be partial title EXAMPLE:
'// aolframe& = findwindow("aol frame25", vbNullString)
'// mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
'// aolchild& = FindChildByTitle(mdiclient&, "Buddy List")
On Error Resume Next
Dim child As Long
Dim childtitle As String
Dim FoundIt As Long
child& = GetWindow(Parent&, GW_CHILD)
Do
  Sleep 0&
    childtitle$ = GetAPIText(child&)
    Sleep 5&
    FoundIt& = InStr(UCase(childtitle$), UCase(title$))
        If FoundIt& <> 0 Then Exit Do
    child& = GetWindow(child&, GW_HWNDNEXT)
Loop Until child& = 0
FindChildByTitle = child&
End Function

Public Function FixAPIString(Text As String) As String
'// Removes null characters if found
On Error Resume Next
    If InStr(Text$, Chr(0)) <> 0 Then FixAPIString = Trim(Mid(Text$, 1, InStr(1, Text$, Chr(0)) - 1))
        If InStr(Text$, Chr(0)) = 0 Then FixAPIString = Text$
End Function

Public Sub Form_Bounce(Frm As Form)
'// Bounce the form all over the place, kind of like when you win
'// one of those solitare card games on the computer
Dim i  As Long
  For i = 1 To 35
    Frm.Left = Int((Rnd * Screen.Width) + 1)
    Frm.Top = Int((Rnd * Screen.Height) + 1)
  Next
End Sub

Public Sub Form_Center(Frm As Form)
'// centers form if you don't want to use the one already in the form's
'// preferences, or if you are using vb4
    Dim X  As Long
    Dim y  As Long
    On Error Resume Next
    X = (Screen.Width - Frm.Width) / 2
    y = (Screen.Height - Frm.Height) / 2
    Frm.Move X, y
End Sub

Public Sub Form_Cool(Frm As Form)
'// this is what i used in one of my programs to unload it
Dim z, X, E, d As Long
'// left
For z = 0 To (Screen.Width) / 2 Step 10
  Frm.Left = Frm.Left - z
If Frm.Left <= 0 Then Exit For
Next
'// up
For X = 0 To (Screen.Height) / 2 Step 10
  Frm.Top = Frm.Top - X
If Frm.Top <= 0 Then Exit For
Next
'// right
For E = 0 To (Screen.Width) / 2 Step 10
  Frm.Left = Frm.Left + E
  If Frm.Left >= (Screen.Width - Frm.Width) / 2 Then Exit For
  Next
'// down
For d = 0 To (Screen.Height) / 2 Step 10
  Frm.Top = Frm.Top + d
  If Frm.Top >= (Screen.Height - Frm.Height) / 2 Then Exit For
Next
End Sub

Public Sub Form_Cool2(Frm As Form)
'// This is what i used in one of my programs, for loading it//
Dim z, X, asdf, df, fe, de As Long
'// right
For z = 0 To (Screen.Width) / 1 Step 10
  Frm.Left = Frm.Left + z
If Frm.Left >= (Screen.Width - Frm.Width) / 1 Then Exit For
Next
'// down
For X = 0 To (Screen.Height) / 1 Step 10
  Frm.Top = Frm.Top + X
If Frm.Top >= (Screen.Height - Frm.Height) / 1 Then Exit For
Next
'// Left
For asdf = 0 To (Screen.Width) / 2 Step 10
  Frm.Left = Frm.Left - asdf
If Frm.Left <= 0 Then Exit For
Next
'// up
For df = 0 To (Screen.Height) / 2 Step 10
  Frm.Top = Frm.Top - df
If Frm.Top <= 0 Then Exit For
Next
'// right middle
For fe = 0 To (Screen.Width) / 2 Step 10
  Frm.Left = Frm.Left + fe
  If Frm.Left >= (Screen.Width - Frm.Width) / 2 Then Exit For
  Next
'// down middle
For de = 0 To (Screen.Height) / 2 Step 3
  Frm.Top = Frm.Top + de
    If Frm.Top >= (Screen.Height - Frm.Height) / 2 Then Exit For
Next de
End Sub

Public Sub Form_CustomFade(Frm As Form, SiR As Long, fade As Long)
'// example call CustomFade(Me, anynumber, anynumber)
'// NOTE: anynumber should be under 255
On Error Resume Next
Dim i As Long
  Frm.Cls
  Frm.ScaleHeight = 128
 For i = 0 To 255 Step 2
   Frm.Line (0, i)-(Frm.ScaleWidth, i + 2), RGB(i, SiR, fade), BF
 Next i
End Sub

Sub Form_FadeHorizon(theForm As Form)
'// form fade found in visual basic help file
Dim A As Long
Dim b
theForm.ScaleHeight = (256 * 2)
For A = 255 To 0 Step -1
theForm.Line (0, b)-(theForm.Width, b + 2), RGB(A + 3, A, A * 3), BF
b = b + 2
Next A
End Sub

Public Sub Form_Greets(sn1LableArray(), Person As String)
'//  This is a scrolling greets sub using an array of lables
'//  You will have to mess with the sn1LableArray(0).Top = 430 sn1LableArray(1).Top = 400 sn1LableArray(0).Left = 30 and the sn1LableArray(1).Left = 0 Numbers
'//  The labels are suppose to look like they have a drop shadow §
'//  To make an array just add 2 lables and name them the same
'sn1LableArray(0) = Person$
'sn1LableArray(1) = Person$
'TimeOut 1000
'    For X = 0 To 1820
'        sn1LableArray(0).Top = sn1LableArray(0).Top - 5
'        sn1LableArray(1).Top = sn1LableArray(1).Top + 5
'    Next X
'sn1LableArray(0) = ""
'sn1LableArray(1) = ""
'sn1LableArray(0).Top = 430
'sn1LableArray(1).Top = 400
'sn1LableArray(0).Left = 30
'sn1LableArray(1).Left = 0
'TimeOut 500
End Sub

Public Sub Form_Max(Frm As Form)
'// lets you maximize your form
On Error Resume Next
Frm.WindowState = 3
End Sub

Public Sub Form_Min(Frm As Form)
'// lets you minimize your own form
On Error Resume Next
Frm.WindowState = 1
End Sub

Public Sub Form_Move(Frm As Form)
'// put this in Form_mousedown to move a form without a titlebar
On Error Resume Next
Dim ReturnVal As Long
Call ReleaseCapture
ReturnVal = SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub

Public Sub Form_Password(Text As textbox, ThePassword As String, Frm2Show As Form)
'// Text is the text box that the user types the PW into, ThePassword is the
'// password you have chosen to allow them to access another form, Frm2Show
'// is the form name that will show Only if they get ThePassword Correct
If LCase(Text) = LCase(ThePassword) Then
    MsgBox "The password you have entered is CORRECT.", vbOKOnly, "Password Correct"
    Frm2Show.Show
Else
MsgBox "Sorry you do not have access to this area", vbOKOnly, "Password Denied"
Form_UnloadAll
End If
End Sub

Public Sub Form_TileImage(TileOn As Object, TileSource As Object)
'// tile any image onto your form or picture box
'// either one can be a form or a picture box
On Error Resume Next
Dim i As Long
Dim j As Long
For i = 0 To TileOn.ScaleWidth Step TileSource.Width
    For j = 0 To TileOn.ScaleHeight Step TileSource.Height
     TileOn.PaintPicture TileSource.Picture, i, j
    Next j
Next i
End Sub

Public Sub Form_Transparent(Frm As Form)
'// sets only your form as transparent, leaving other controls visible
On Error Resume Next
SetWindowLong Frm.hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
SetWindowPos Frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME
End Sub

Public Sub Form_UnloadAll()
'// a good way to unload all your forms from the computer's memory
'// put it in form_unload query or form_unload
Dim Frm As Form
For Each Frm In Forms
      Unload Frm
      Set Frm = Nothing   ' remove the object from memory
Next
End Sub

Public Sub Form_WipeRight(Frm As Form)
'// Make a form shrink then it will unload
On Error Resume Next
Dim i As Long
    For i = 1 To Frm.Width Step 20
          If Frm.Width <= 1700 Then: Unload Frm: Exit For
          Frm.Width = Frm.Width - i
     Next
End Sub

Public Function GetAPIText(hwnd As Long) As String
'// Gets text from a window, static, text, caption.....if it has text to get
'// this will get it
Dim X As Long
Dim Text As String
    On Error Resume Next
    X = SendMessageByNum(hwnd&, WM_GETTEXTLENGTH, 0, 0)
    Text$ = Space(X + 1)
    X = SendMessageByString(hwnd&, WM_GETTEXT, X + 1, Text$)
GetAPIText$ = FixAPIString(Text$)
End Function

Public Function GetHalfText(hwnd As Long) As String
'// Gets text from a window, static, text, caption.....if it has text to get
'// this will get 1/3 of it...... i know i named it halftext but thirdtext sounded dumb heh
Dim X As Long
Dim Text As String
    On Error Resume Next
    X = SendMessageByNum(hwnd&, WM_GETTEXTLENGTH, 0, 0)
    Text$ = Space(X + 1)
    X = SendMessageByString(hwnd&, WM_GETTEXT, X + 1, Text$)
    Text$ = Mid(Text$, X / 3)
GetHalfText$ = FixAPIString(Text$)
End Function

Public Function INI_Get(title As String, Heading As String) As String
'//  Read from your own configuration ini file
'//  EXAMPLE: text1 = INI_Get("Timeout", "Scroll Pause:")
Dim Amount As Long
Dim datastring As String * 1024
 datastring = Space$(1024)
    Amount& = GetPrivateProfileString(title$, Heading$, " ", datastring, 1024, App.Path + "\sirvb6.ini")
        If Amount& = 0 Then
            INI_Get = ""
        Else
            INI_Get = RTrim$(datastring)
        End If
End Function

Public Sub INI_Write(title As String, Nameunder As String, theData As String)
'//  Write to your own configuration ini file
'//  EXAMPLE Call INI_Write("Timeout", "Scroll Pause:", (text1))
Dim Dummy As Long
Dummy& = WritePrivateProfileString(title$, Nameunder$, theData$, App.Path + "\sirvb6.ini")
End Sub

Public Function isListed(List As listbox, SearchString As String) As Boolean
'// use this to see if a screen name is listed in a listbox, if the SN is listed it
'// will return True, if its not listed, then it is False, and you can add the SN
Dim Down As Long
 For Down = 0 To List.ListCount - 1
   If InStr(1, List.List(Down), SearchString$) <> 0 Then isListed = True: Exit Function
 Next Down
isListed = False
End Function

Public Sub LB2LB(Source As Long, Target As listbox)
'// a fast way of sending the contents of one listbox to another listbox
On Error Resume Next
Dim lstDown As Long
   Dim numitems As Long
   Dim sItemText As String * 255
  '// get the number of items in the source list
   numitems& = SendMessageLong(Source, LB_GETCOUNT, 0&, 0&)
  '// if it has contents, copy the items to the target list
   If numitems& > 0 Then
      For lstDown = 0 To numitems - 1
         Call SendMessageByString(Source, LB_GETTEXT, lstDown, sItemText)
         Call SendMessageByString(Target.hwnd, LB_ADDSTRING, 0&, sItemText)
      Next
   End If
End Sub

Public Sub LetterArray(sWord As String)
'//  Put each letter into an array, this is for personal use, but if you know your stuff
'//  You can turn this into a fader, scramble function, or just about anything
Dim Letter(99) As String
Dim l As Long
For l = 1 To Len(sWord)
   Letter(l) = Mid(sWord, l, 1)  'goes through each letter assigning a number to each one
Next l
End Sub

Public Sub List_Add(List As listbox, txt As String)
'// use this to avoid adding duplicate items into a textbox
'// its like an anti dupe, use it in your text_keypress when you
'// add the string to the listbox
On Error Resume Next
Dim result As Long
If txt$ = "" Then: Exit Sub
result& = SendMessageByString(List.hwnd, LB_FINDSTRINGEXACT, 0&, txt$)
  If result& = -1 Then
    Call SendMessageByString(List.hwnd, LB_ADDSTRING, 0&, txt$)
  End If
End Sub

Public Sub List_AllAscii(List As listbox)
'// adds all the ascii characters between 33 and 255 to a listbox
Dim X As Long
For X = 33 To 255
    List.AddItem Chr(X)
Next X
End Sub

Public Sub List_Clear(List As Long)
'//  If you want to use this instead of List1.Clear you will have to put List_Clear List1.hwnd
'//  However >=) this will clear list boxes outside of your form, just use the window hwnd handle ;)
Call SendMessage(List&, LB_RESETCONTENT, 0, ByVal 0&)
End Sub

Public Sub List_KillDupes(listbox As listbox)
'// search through a listbox and look for duplicate instances
'// if so it removes the dupes and leaves 1 instance of it
Dim SearchA As Long
Dim SearchB As Long
Dim KillDupes As Long
 KillDupes = 0
    For SearchA& = 0 To listbox.ListCount - 1
      For SearchB& = SearchA& + 1 To listbox.ListCount - 1
           KillDupes = KillDupes + 1
            If listbox.List(SearchA&) = listbox.List(SearchB&) Then
             listbox.RemoveItem SearchB&
             SearchB& = SearchB& - 1
           End If
      Next SearchB&
    Next SearchA&
End Sub

Public Sub List_Load(TheList As listbox, FileName As String)
'// Loads a file to a list box
On Error Resume Next
Dim TheContents As String
Dim fFile As Integer
fFile = FreeFile
 Open FileName For Input As fFile
   Do
     Line Input #fFile, TheContents$
        Call List_Add(TheList, TheContents$)
   Loop Until EOF(fFile)
 Close fFile
End Sub

Public Sub List_Save(TheList As listbox, FileName As String)
'// Save a listbox as FileName
On Error Resume Next
Dim Save As Long
Dim fFile As Integer
fFile = FreeFile
Open FileName For Output As fFile
   For Save = 0 To TheList.ListCount - 1
      Print #fFile, TheList.List(Save)
   Next Save
Close fFile
End Sub

Public Sub List_Remove(List As listbox)
On Error Resume Next
If List.ListCount < 0 Then Exit Sub
  List.RemoveItem List.ListIndex
End Sub

Public Sub List_RemString(List As listbox, RemoveString As String)
'// Removes a String from a listbox, use this if you want to remove
'// a screen name or certain word from a listbox
Dim result As Long
result& = SendMessageByString(List.hwnd, LB_FINDSTRINGEXACT, 0&, RemoveString$)
If result& > -1 Then
    List.RemoveItem (result&)
End If
End Sub

Public Sub List_SendRoom(List As listbox)
'// scroll a listbox into the chat room
Dim downlst As Long
For downlst = 0 To List.ListCount - 1
    AOL_SendRoom List.List(downlst)
    Pause 0.6
Next downlst
End Sub

Public Sub List2TextBCC(List As listbox, Text As textbox)
'// after you use AddRoom, you can use this to blind carbon copy
'// everyone on that list, to send them email
Dim downlst As Long
For downlst = 0 To List.ListCount - 1
      Text = Text.Text & "(" & List.List(downlst) & "@aol.com), "
Next downlst
End Sub

Public Sub MacFilter(TextB As textbox, TxtString As String, Name As String)
'//  Use this if you are making a macro draw - this use the MacroFilter function above
'//  Example
'//  Dim txt As String
'//  txt$ = Text1
'//  Call MacFilter(Text1, txt$, "Solid")
On Error Resume Next
  Select Case UCase(Name$)
    Case "DARKEN"
        TextB = MacroFilter(TxtString$, ":", ";")
    Case "LIGHTEN"
        TextB = MacroFilter(TxtString$, ";", ":")
    Case "CURVES"
        TextB = MacroFilter(TxtString$, "|", ")")
        TextB = MacroFilter(TxtString$, "l", "(")
        TextB = MacroFilter(TxtString$, "I", ")")
    Case "DASHES"
        TextB = MacroFilter(TxtString$, "_", "...")
        TextB = MacroFilter(TxtString$, "|", ":")
        TextB = MacroFilter(TxtString$, "l", ";")
        TextB = MacroFilter(TxtString$, "I", "!")
        TextB = MacroFilter(TxtString$, "/", ",'")
        TextB = MacroFilter(TxtString$, "\", "',")
    Case "SOLID"
        TextB = MacroFilter(TxtString$, "...", "_")
        TextB = MacroFilter(TxtString$, ":", "|")
        TextB = MacroFilter(TxtString$, ";", "l")
        TextB = MacroFilter(TxtString$, "!", "I")
        TextB = MacroFilter(TxtString$, ",'", "/")
        TextB = MacroFilter(TxtString$, "',", "\")
    Case "FILL"
        TextB = MacroFilter(TxtString$, " ", "::")
    Case "SHADE"
        TextB = MacroFilter(TxtString$, ",'", ",;;;")
        TextB = MacroFilter(TxtString$, "',", ";;;,")
  End Select
End Sub

Public Function MacroFilter(MainString As String, String2Replace As String, ReplaceWith As String) As String
'//  Just something you can make custom macro filters with
MacroFilter$ = Replace(MainString$, String2Replace$, ReplaceWith$)
End Function

Public Sub mIRC_AddRoom(List As listbox)
'// used to add any mIRC chat room list names into a listbox
On Error Resume Next
Dim mirc As Long
Dim mdiclient As Long
Dim channel As Long
Dim listbox As Long
 mirc& = FindWindow("mirc32", vbNullString)
   mdiclient& = FindWindowEx(mirc&, 0&, "mdiclient", vbNullString)
     channel& = FindWindowEx(mdiclient&, 0&, "channel", vbNullString)
  listbox& = FindWindowEx(channel&, 0&, "listbox", vbNullString)
LB2LB listbox&, List
End Sub

Public Sub mIRC_SendRoom(SayWhat As String)
'// SendChat text into a mIRC chat room
On Error Resume Next
Dim mirc As Long
Dim mdiclient As Long
Dim channel As Long
Dim edit As Long
 mirc& = FindWindow("mirc32", vbNullString)
    mdiclient& = FindWindowEx(mirc&, 0&, "mdiclient", vbNullString)
        channel& = FindWindowEx(mdiclient&, 0&, "channel", vbNullString)
        edit& = FindWindowEx(channel&, 0&, "edit", vbNullString)
    SetText edit&, SayWhat$
 EnterKey edit&
End Sub

Public Function MouseIn(Btn As Control) As Boolean
'// checks to see if the mouse is inside of a control, if it is
'// it will return True and you can do something (the control must have a hwnd handle)
Dim MousePos As POINTAPI
Dim Dummy As Long 'dummy variable for the call
On Error GoTo FUCK 'Resume Next 'typical error controller.
 Dummy = GetCursorPos(MousePos) ' Get the position of cursor.
    If WindowFromPoint(MousePos.X, MousePos.y) = Btn.hwnd Then
    MouseIn = True 'if mouse if over then its true.
    End If
FUCK:
End Function

Public Sub PasteMacro(Text As textbox, lst As listbox)
'//  Put a listbox's text in text box. Put this in double click of a list box
'//  good for use in an ascii shop or macro shop
  Text.SelText = lst.Text
  Text.SetFocus
End Sub

Public Sub Pause(interval As Long)
'//  this is same type of thing as TimeOut, if you want to put a pause in your code
'//  to make it wait for a certain amount of time, just put Pause 1.5
'//  in your code and that will make it pause 1.5 seconds before continuing
Dim Current As Long
    On Error Resume Next
    Current = Timer
    Do While Timer - Current < Val(interval)
    DoEvents
    Loop
End Sub

Public Sub PlayWav(wavName As String)
'// Play a wav in your program without freezing it
'// if a wav is already playing it will not process your request
'// to play a new wav
On Error Resume Next
Call SndPlaySound(wavName$, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP)
End Sub

Public Sub RebootSys()
'// reboot your computer using this
On Error Resume Next
Call ExitWindowsEx(EWX_REBOOT, 0&)
End Sub

Public Sub RunMenu(Main_Prog As String, Top_Position As String, Menu_String As String)
'// didn't write the original, i just converted it to  vb6 from vb3
'// works good, you can even use it on notepad if you wanted :P
'// Call RunMenu("aol frame25", "&Help", "&About America Online")
On Error GoTo stp
Dim Top_Position_Num As Long
Dim buffer As String
Dim Look_For_Menu_String As Long
Dim Trim_Buffer As String
Dim Sub_Menu_Handle As Long
Dim BY_POSITION As Long
Dim Get_ID As Long
Dim Click_Menu_Item As Long
Dim Menu_Parent As Long
Dim aol As Long
Dim Menu_Handle As Long
Dim Parent As Long
Top_Position_Num = -1
Parent& = FindWindow(Main_Prog, vbNullString)
Menu_Handle = GetMenu(Parent&)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    buffer$ = String$(255, 0)
    Look_For_Menu_String& = GetMenuString(Menu_Handle, Top_Position_Num, buffer$, Len(Top_Position) + 1, WM_USER)
    Trim_Buffer = FixAPIString(buffer$)
    If Trim_Buffer = Top_Position Then Exit Do
    If GetMenuItemID(Menu_Handle, Top_Position_Num) = 0 Then Exit Do
Loop
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    buffer$ = String(255, 0)
    Look_For_Menu_String& = GetMenuString(Sub_Menu_Handle, BY_POSITION, buffer$, Len(Menu_String) + 1, WM_USER)
    Trim_Buffer = FixAPIString(buffer$)
    If Trim_Buffer = Menu_String Then Exit Do
    If GetMenuItemID(Menu_Handle, BY_POSITION) = 0 Then Exit Do
Loop
DoEvents
Get_ID& = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = SendMessageByNum(Parent&, WM_COMMAND, Get_ID&, 0&)
stp:
End Sub

Public Sub RunMenuByName(Application As Long, StringSearch As String)
'// This one i had more to write and correct, it didn't work really good
'// RunMenuByName aolframe&, "&Sign Off"
On Error Resume Next
Dim ToSearch As Long
Dim MenuCount As Long
Dim FindString As Long
Dim GetString As Long
Dim SubCount As Long
Dim MenuString As String
Dim ToSearchSub As Long
Dim MenuItemCount As Long
Dim GetStringMenu As Long
Dim MenuItem As Long
Dim RunTheMenu As Long
    ToSearch& = GetMenu(Application)
    MenuCount& = GetMenuItemCount(ToSearch&)
For FindString = 0 To MenuCount& - 1
     ToSearchSub& = GetSubMenu(ToSearch&, FindString)
        MenuItemCount& = GetMenuItemCount(ToSearchSub&)
  For GetString = 0 To MenuItemCount& - 1
   SubCount& = GetMenuItemID(ToSearchSub&, GetString)
     MenuString$ = String$(100, " ")
      GetStringMenu& = GetMenuString(ToSearchSub&, SubCount&, MenuString$, 100, 1)
       If InStr(UCase(MenuString$), UCase(StringSearch)) Then
        MenuItem& = SubCount&
        GoTo MatchString
      End If
  Next GetString
Next FindString
MatchString:
RunTheMenu& = SendMessage(Application, WM_COMMAND, MenuItem&, 0)
End Sub

Public Function Scramble(txt As String) As String
'//  not mine scrambles text for a scrambler game
'//  txt$ = scramble(text1)
'//  aol_sendroom "-=[ word: " & txt$
On Error Resume Next
Dim Word As String
Dim buffer As String
Dim Random As Long
Dim i As Long
Dim A As Long
Separate:
Do: DoEvents
    A& = InStr(txt$, " ")
    If A& = 0 Then
        buffer$ = txt$
        txt$ = ""
        Exit Do
    End If
    If A& = 1 Then
        Scramble$ = Scramble$ & " "
        txt$ = Right$(txt$, Len(txt$) - 1)
    End If
    If A& > 1 Then
        buffer$ = Left$(txt$, A& - 1)
        txt$ = Right$(txt$, Len(txt$) - A& + 1)
        Exit Do
    End If
Loop Until A& = 0
Word$ = ""
For i& = 1 To Len(buffer$) - 1
    Random& = Int(Len(buffer$) * Rnd + 1)
    Word$ = Word$ & Mid$(buffer$, Random&, 1)
    buffer$ = Left$(buffer$, Random& - 1) & Right$(buffer$, Len(buffer$) - Random&)
Next i&
Word$ = Word$ & buffer$
Scramble$ = Scramble$ & Word$
If txt$ <> "" Then GoTo Separate
End Function

Public Sub ScreenSaver_On(Frm As Form)
'//  duh it starts up the default screen saver
Call SendMessage(Frm.hwnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub

Public Sub SendCharNum(win, chars)
'// send a key button to a window, see the EnterKey sub
'// in it Chars would = 13 which is the ascii value of the enter key
On Error Resume Next
Dim Dummy As Long
Dummy& = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub

Public Sub SendMacro(txt As String)
'// scroll a multiline text box into a chat room
'// to make a stop button for this, Simply put StopScroll = True
'// into a button/label/menu
StopScroll = False
txt = txt & vbCrLf
Do While (InStr(txt, vbCr) <> 0)
Pause 0.6
    AOL_SendRoom Mid("  " + txt, 1, InStr("  " + txt, vbCr) - 1)
    txt = Mid("  " + txt, InStr("  " + txt, vbCrLf) + 2)
If StopScroll = True Then: Exit Do
Loop
End Sub

Public Sub SendVKey(win As Long, Key As VirtualKeys)
'// kind of like the non-lame API version of SendKeys heh :P
On Error Resume Next
 Call SendMessageLong(win&, WM_KEYDOWN, Key, 0&)
 Call SendMessageLong(win&, WM_KEYUP, Key, 0&)
End Sub

Public Sub SetText(win As Long, txt As String)
'// you have to find the text box you want to set the text to
'// then once located you would put something like this in your code
'// SetText aoledit&, "hello room"
On Error Resume Next
Dim TheText As Long
TheText& = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Public Sub StayOnTop(theForm As Form)
'// makes your form stay on top of all other windows
'// in form_load put StayOnTop Me and also put it in
'// form_resize to make sure it will stay on top when you min/max it
On Error Resume Next
Dim SetWinOnTop As Long
    SetWinOnTop& = SetWindowPos(theForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Function StripHTML(HTMLString As String) As String
'// Strip HTML tags and replace the <BR> with vbCrLf (carriage return & line feed)aka chr(13) & chr(10)
'// this is super sweet, looks for anything between < & > and removes them, therefore removing
'// all html tags, only thing that may cause a problem is if someone <types like this>
'// but who does that anyways?  heh this will probably be the thing most copied in other bas files
On Error Resume Next
Dim i As Long
Dim HoldString As String
Dim sLook As String
Dim sHold As String
HTMLString$ = Replace(HTMLString$, "<BR>", vbCrLf)
  For i = 1 To Trim(Len(HTMLString$))
    sLook$ = Mid(HTMLString$, i, 6)
    sHold$ = Mid(HTMLString$, i, 1)
      If sHold$ = "<" Then
        Do Until sHold$ = ">"
         DoEvents
         i = i + 1
         sHold = Mid(HTMLString$, i, 1)
        Loop
     sHold$ = ""
    End If
   HoldString = HoldString$ & sHold$
  Next i
StripHTML$ = HoldString$
End Function

Public Function Text_Count(txtObj As textbox) As Long
'//  counts lines in a text box
'//  example:    x = text_count(text1) : label1.Caption = x
On Error Resume Next
If txtObj.MultiLine = True Then
Text_Count = SendMessage(txtObj.hwnd, EM_GETLINECOUNT, 0, 0&)
Else
Text_Count = 1
End If
End Function

Public Function Text_CountOccur(textbox As textbox, What2Count As String) As Long
'// Count the number of occurances in any text box, you can use this as
'// a non-api way of counting lines if you search for chr(13) or even find the number
'// of spaces in a text box and therefore counting the amount of words in a textbox
On Error Resume Next
Dim CountNum As Long
Dim Source As String
 CountNum& = 0
  Source$ = textbox.Text
   Do While InStr(Source$, What2Count$)
     CountNum& = CountNum& + 1
     Source = Mid$(Source$, InStr(Source$, What2Count$) + 1)
   Loop
Text_CountOccur = CountNum&
End Function

Public Sub Text_Load(Text As textbox, FileName As String)
'// load a textbox
On Error Resume Next
Dim sFile As String
Dim nFile As Variant
    sFile = FileName$
    nFile = FreeFile
    Open sFile For Input As nFile
        Text = Input(LOF(nFile), nFile)
    Close nFile
End Sub

Public Sub Text_Readonly(Text As textbox)
'//  makes it so nobody can type in a text box
'//  also removes the paste menu from the popup menu
Dim iReturn As Long
iReturn = SendMessage(Text.hwnd, EM_SETREADONLY, True, 0&)
End Sub

Public Sub Text_ReadonlyRemove(Text As textbox)
'//  removes the readonly in the textbox so the user can type in it again
Dim iReturn As Long
iReturn = SendMessage(Text.hwnd, EM_SETREADONLY, False, 0&)
End Sub

Public Sub Text_Save(TextB As textbox, FileName As String)
'// save a textbox without overwriting
On Error Resume Next
Dim file As String
Dim X As VbMsgBoxResult
Dim FreeFile
Dim tFile
Dim wFile As String
   file$ = FileName
        If file$ = "" Then Exit Sub
        If Len(Dir(file$)) > 0 Then
        X = MsgBox("This file already exists: [ " & file$ & " ] do you wish replace it?", vbYesNo, "e=mscrambler²")
        If X = vbNo Then Exit Sub
        If X = vbYes Then GoTo SaidYes:
    End If
SaidYes:
    tFile = 1
    wFile = file$
    Open wFile For Output As tFile
        Print #tFile, TextB
    Close tFile
End Sub

Public Sub Text_Undo(Text As textbox)
'// undoes deleted and keypressed text in textboxes
On Error Resume Next
Dim UndoResult As Long
UndoResult = SendMessage(Text.hwnd, EM_UNDO, 0&, 0&)
End Sub

Public Sub TimeOut(interval As Long)
'//  this is different than pause, because you can have much shorter timeouts
'//  or pauses. to use this you would put timeout 500   that would be a half a second
'//  this works the same, its just something different that not alot of others use.
Dim Current As Long
    On Error Resume Next
    Current = GetTickCount()
    Do While Current < Val(interval)
    DoEvents
    Loop
End Sub

Public Sub Tnet_AddTeam(Team As String)
'//  Add or change your tetrinet team name
Dim tform As Long
Dim tpanel As Long
Dim tedit As Long
Dim tbutton As Long
 tform& = FindWindow("tform1", vbNullString)
   tpanel& = FindWindowEx(tform&, 0&, "tpanel", vbNullString)
   tpanel& = FindWindowEx(tform&, tpanel&, "tpanel", vbNullString)
   tpanel& = FindWindowEx(tform&, tpanel&, "tpanel", vbNullString)
   tpanel& = FindWindowEx(tform&, tpanel&, "tpanel", vbNullString)
   tpanel& = FindWindowEx(tform&, tpanel&, "tpanel", vbNullString)
   tpanel& = FindWindowEx(tform&, tpanel&, "tpanel", vbNullString)
   tpanel& = FindWindowEx(tform&, tpanel&, "tpanel", "Partyline ")
    ClickIt tpanel&
        Pause 0.4
   tform& = FindWindow("tform1", vbNullString)
   tpanel& = FindWindowEx(tform&, 0&, "tpanel", vbNullString)
   tedit& = FindWindowEx(tpanel&, 0&, "tedit", vbNullString)
   tbutton& = FindWindowEx(tpanel&, 0&, "tbutton", vbNullString)
Call SendMessageByString(tedit&, WM_SETTEXT, 0, Team$)
  Pause 0.1
SendVKey tbutton&, SPACE_VKEY
End Sub

Public Function Tnet_GetChat() As String   '// example text1 = tnet_getchat
'//  Get the tetrinet chat text, can be useful if you wanted to make a small
'//  m-chat, so you can monitor the chat room while the tnet form isn't the top
'//  window.  Then you can see when the other players are saying/when another game starts.
Dim tform As Long
Dim tpanel As Long
Dim tsrichedit As Long
tform& = FindWindow("tform1", vbNullString)
  tpanel& = FindWindowEx(tform&, 0&, "tpanel", vbNullString)
    tsrichedit& = FindWindowEx(tpanel&, 0&, "tsrichedit", vbNullString)
  Tnet_GetChat$ = GetAPIText(tsrichedit&)
End Function

Public Function Tnet_LastLineSN() As String
'//  I looked at another bas and they had like 15 lines of code, this took 3 lines ;)
On Error Resume Next
Dim StringTemp As String
Dim Last As Long
Dim NewString As String
Dim Arrow As Long
Dim LastLine As String
StringTemp$ = Tnet_GetChat()
Last& = InStrRev(StringTemp$, "<")
LastLine$ = Mid(StringTemp$, Last& + 1)
Arrow& = InStr(1, LastLine$, ">")
Tnet_LastLineSN$ = Mid(LastLine$, 1, Arrow& - 1)
End Function

Public Function Tnet_LastLineTxt() As String
'//  Gets what the last person said
Dim StringTemp As String
Dim Last As Long
Dim NewString As String
StringTemp$ = Tnet_GetChat()
Last& = InStrRev(StringTemp$, ">")
Tnet_LastLineTxt$ = Mid(StringTemp$, Last& + 2)
End Function

Public Sub Tnet_SendRoom(WhatToSay As String)
'// Tetrinet is getting to be a favorite game of alot of people so i just
'// added this for anyone who wants to make some kind of stuff for tnet
On Error Resume Next
Dim tform&
Dim tpanel&
Dim tedit&
tform& = FindWindow("tform1", vbNullString)
 tpanel& = FindWindowEx(tform&, 0&, "tpanel", vbNullString)
  tedit& = FindWindowEx(tpanel&, 0&, "tedit", vbNullString)
  tedit& = FindWindowEx(tpanel&, tedit&, "tedit", vbNullString)
 SetText tedit&, WhatToSay$
EnterKey tedit&
End Sub

Public Sub Win_CloseWin(Winder As Long)
'// close a window by its handle
On Error Resume Next
Dim Dummy As Long
    Dummy& = SendMessageByNum(Winder, WM_CLOSE, 0&, 0&)
End Sub

Public Sub Win_Flash(Wnd2Flash As Long, Times2Flash As Long)
'// Find the window you want to flash, or you can use Me.Hwnd
'// then you can use Flash Me.Hwnd, 10   and the window will flash 10 times
Dim i As Long
  For i = 0 To Times2Flash
    Call FlashWindow(Wnd2Flash, True)
    Pause 1
  Next i
Call FlashWindow(Wnd2Flash, False)
End Sub

Public Sub Win_FocusOn(Winder As Long)
'// Set focus on any window
Dim Dummy As Long
    Dummy& = SetFocus(Winder&)
End Sub

Public Sub Win_Min(win As Long)
'// this minimizes any window outside of your program, i don't know why they
'// chose to call it CloseWindow in api, they just do.
On Error Resume Next
Dim Dummy As Long
    Dummy& = CloseWindow(win)
End Sub

Public Sub WindowSPY(TextB As textbox)
 '// This was on my web page so i threw it in here, its not written by me
 '// its from the people who made the freespy, or maybe from Microsoft MSDN
 '// Call This In A Timer
 '// WindowSPY Text1
    Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
    Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
    Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
    Dim hInstance As Long, sParentWindowText As String * 100
    Dim sModuleFileName As String * 100, r As Long
    Dim WinHdl As String, wintxt As String, WinClass As String, WinStyle As String
    Dim WinIDNum As String, WinPHandle As String, WinPText As String
    Dim WinPClass As String, WinModule As String
    Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.y
    hWndOver = WindowFromPointXY(ptx, pty)
    If hWndOver <> hWndLast Then
        TextB = WinHdl & vbCrLf & WinClass & vbCrLf & wintxt & vbCrLf & WinStyle & vbCrLf & WinIDNum & vbCrLf & WinPHandle & vbCrLf & WinPText & vbCrLf & WinPClass & vbCrLf & WinModule
        hWndLast = hWndOver
        WinHdl = "Window Handle: " & hWndOver
        r = GetWindowText(hWndOver, sWindowText, 100)
        wintxt = "Window Text: " & Left(sWindowText, r)
        r = GetClassName(hWndOver, sClassName, 100)
        WinClass = "Window Class Name: " & Left(sClassName, r)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle = "Window Style: " & lWindowStyle
        hWndParent = GetParent(hWndOver)
        If hWndParent <> 0 Then
            wID = GetWindowWord(hWndOver, GWW_ID)
            WinIDNum = "Window ID Number: " & wID
            WinPHandle = "Parent Window Handle: " & hWndParent
            r = GetWindowText(hWndParent, sParentWindowText, 100)
            WinPText = "Parent Window Text: " & Left(sParentWindowText, r)
            r = GetClassName(hWndParent, sParentClassName, 100)
            WinPClass = "Parent Window Class Name: " & Left(sParentClassName, r)
        Else
            WinIDNum = "Window ID Number: N/A"
            WinPHandle = "Parent Window Handle: N/A"
            WinPText = "Parent Window Text : N/A"
            WinPClass = "Parent Window Class Name: N/A"
        End If
        hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
        r = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule = "Module: " & Left(sModuleFileName, r)
        TextB = WinHdl & vbCrLf & WinClass & vbCrLf & wintxt & vbCrLf & WinStyle & vbCrLf & WinIDNum & vbCrLf & WinPHandle & vbCrLf & WinPText & vbCrLf & WinPClass & vbCrLf & WinModule
    End If
End Sub


Public Sub WWWaddress(Address As String)
'// open up the default web browser and send it to a web page address
On Error Resume Next
Dim ReturnVal As Long
ReturnVal& = Shell("Start.exe " & Address$, vbHide)
End Sub

'ATTENTION
'========================================
'Public Function Replace(txtObject As String, sWhat As String, sWith As String) As String
'//  vb4 and vb5 Users Uncomment this function
'//  The "Replace" function is new to VB6 and all you need is Replace(textbox, stringtoreplace, newstring)
'Dim text As String
'Dim Where As Long
'Dim sRight As String
'text$ = txtObject$
'Do While (InStr(1, text$, sWhat$, 1) > 0)
'    Where = InStr(1, text$, sWhat$)
'    If (Where > 0) Then
'        LeftSide$ = Mid(text$, 1, Where - 1)
'        sRight$ = Mid(text$, Where + Len(sWhat$))
'        text$ = LeftSide$ + sWith$ + sRight$
'        Replace = text$
'    End If
'Loop
'Replace = text$
'End Function
'======================================
'___________________________________________.
'Ran out of ideas
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Carpal Tunnel is setting in i'm finished with this version
'======================================
Private Sub ZZZ__How_To__ZZZ()
'This will just be sort of a tutorial or a FAQ area.
'
'1.)   Using DoS' Chat Scan
'
'       first off, if you don't understand how to use DoS' chat scan i am
'       suprised you can breathe.  It is a very simple control to use.
'       The first thing you need to do is turn the scan on  Chat1.ScanOn
'       Now, in the control you have Screen_Name and What_Said.
'       As an example we will make a simple lister bot
'       If UCase(What_Said) = UCase("/LIST ME") Then
'           If isListed(List1, Screen_Name) = False Then
'               List1.AddItem Screen_Name
'               AOL_SendRoom "-=[ " & Screen_Name & " has been added."
'           End If
'       End If
'       Now to turn it off you just put Chat1.ScanOff in a button, Told you it was simple.
'
'2.)   Common Dialog:  Save And Load List
'
'       *LOAD LIST*
'       Dim sPath as String
'        On Error GoTo Err_Errored
'           CommonDialog1.Filter = "All Files(*.*)|*.*|"
'           CommonDialog1.FilterIndex = 1
'           CommonDialog1.Action = 1   'Load Action
'         sPath$ = CommonDialog1.FileName
'           Call List_Load(List1, sPath$)
'         Err_Errored: Exit Sub
'
'       *SAVE LIST*
'       Dim sPath as String
'       Dim mbResult As VbMsgBoxResult
'       On Error GoTo Hell:
'       CommonDialog1.Filter = "All Files (*.*)|*.*|"
'       CommonDialog1.FilterIndex = 1
'       CommonDialog1.Action = 2   'Save Action
'       sPath$ = CommonDialog1.FileName
'       If DoesFileExist(sPath$) = True Then
'         mbResult = MsgBox("The file " & sPath$ & " already exists, are you sure you want to continue?", vbYesNo, "Saving Problem")
'       If mbResult = vbYes Then
'           Call List_Save(List1, sPath$)
'       End If
'       End If
'       Hell: Exit Sub
'       *NOTE*
'       If you want to load a text box instead of list box then just change
'       it from List_Load to Text_Load and same for List_Save to Text_Save
'
'3.)   SpyWorks Subclassing For KeyPresses
'       Step1 - After adding the DWSBC32.OCX control you must figure out
'                  which keypress you want to keep control of.  Look in your help
'                  file for the ascii value of that key.  Example, Enter would be 13
'       Step2 - Locate the window area you wish to monitor for keypresses
'                   point your api spy program (i suggest PATorJK's api spy) and
'                   find the window/child area you wish to subclass, for this example
'                   i will find the AOL IM area that you type in.  Create a button
'                   this will be the Start Subclass Button... now in it put
'                     Dim theIM as Long
'                     Dim aolrich as Long
'                     theIM& = AolChildByTitle("Instant Message")
'                     aolrich& = FindWindowEx(theIM&, 0&, "richcntl", vbNullString)
'                     aolrich& = FindWindowEx(theIM&, aolrich&, "richcntl", vbNullString)
'                     SubClass1.HwndParam = aolrich&
'       Step3 - Code To Capture
'                   the rest is a very simple if/then statement
'                   in your SubClass1_WndMessageX put
'                     If wp = 13 Then   'if enter is pressed then
'                       theIM& = AolChildByTitle("Instant Message")
'                       aolicon& = FindWindowEx(theIM&, 0&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       aolicon& = FindWindowEx(theIM&, aolicon&, "_aol_icon", vbNullString)
'                       ClickIt aolicon&  'click send
'                       End If
'       Step4 - Stop Subclass Button
'                    make one more button and put SubClass1.HwndParam = 0& in it to
'                    stop subclassing the IM typing box. then save, run and test
End Sub

Public Sub ZZZ_Me_ZZZ()
'Name:                       Steve
'Age:                          24
'State:                        Ohio
'Years Programming:   2 - Started in vb3, moved to vb5 and now up to vb6
'Homepage:                8op.com screwed up and formatted their hard drive with all
'                                the web pages on it, so we all got screwed and my page was
'                                deleted.  But you can check http://www.escrambler.com to see
'                                if my is posted there in the future.
'Note:                        If you use this module, please read through the code and try to understand
'                                what i did and how i did it.  The only reason i made this was to teach other
'                                programmers how to do things, some in different ways than others normally
'                                and hopefully better ways to do them.
End Sub

