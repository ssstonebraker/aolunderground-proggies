Attribute VB_Name = "keeb232"
'keeb232.bas by keeb (154 subs/functs)
' shouts go to :
'  cassy   warp   idal   bust   igneus'
' ChiChis   naz    tik   RAW   bust'
'rca    busta   ghst   meeh   dos'
' hound   glow   cia   dolan   Nate'
'  eses   file   allah   kast    tina'
' ashley   lynch   rem   MaceX   jnco'
'  ecko   409   cyx   sniff   skeme'
'   argon   carg0   cargo  MBomber'
'  har0   tool   n   tab'
' that's all i can think of
'if i forgot you i am sorry
' shouts out to everyone in
'  special unit² production
' thanks for downloading
'gotta love it in double 0
' visit www.keebsoft.com
'  for more shit
' and updates
'keeb out
'====================================
'         keeb232.bas by keeb
'        dont be a code copier
'           contact me at
'       mustbejewishh@aol.com
'      educational purposes only
'====================================
Public RoomHandle&
Public Declare Function AddPort Lib "winspool.drv" Alias "AddPortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pMonitorName As String) As Long
Public Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hWndInsertAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags&)
Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SenditByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Declare Function SenditbyNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Declare Function GetWindow& Lib "user32" (ByVal hwnd&, ByVal wCmd&)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd&, ByVal lpClassName$, ByVal nMaxCount&)
Declare Function GetWindowTextLength& Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd&)
Declare Function GetWindowText& Lib "user32" Alias "GetWindowTextA" (ByVal hwnd&, ByVal lpString$, ByVal cch&)
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef Dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SelectObject Lib "GDI" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu&) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Sub Rectangle Lib "GDI" (ByVal hDC As Integer, ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
Declare Sub ReleaseCapture Lib "user32" ()
Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230
Public Const WM_SYSCOMMAND = &H112
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const HWND_TOPMOST = -1
Public Const HWND_TOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const VK_SPACE = &H20
Public Const VK_RETURN = &HD
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const SC_MOVE = &HF012
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
Public Const VK_TAB = &H9
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
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
Public Const NUM_SQUARES = 9
Public Const NUM_PLAYERS = 2
Public Const PLAYER_DRAW = -1
Public Const PLAYER_NONE = 0
Public Const PLAYER_HUMAN = 1
Public Const PLAYER_COMPUTER = 2
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const PROCESS_VM_READ = &H10
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

Function PeriodText(Text As String)
    For i = 1 To Len(Text)
        TakeText$ = Mid(Text, i, 1)
        hmm$ = Mid(Text, i + 1, 1)
        If l$ = " " Then
            Spacee$ = Spacee$ & TakeText$
        ElseIf hmm$ <> " " Then
            Spacee$ = Spacee$ & TakeText$ & "."
        Else
            Spacee$ = Spacee$ & TakeText$
        End If
    Next i
    PeriodText = Spacee$
End Function

Function OwnText(MakeIt As String, Text As String)
    For i = 1 To Len(Text)
        TakeText$ = Mid(Text, i, 1)
        hmm$ = Mid(Text, i + 1, 1)
        If l$ = " " Then
            Spacee$ = Spacee$ & TakeText$
        ElseIf hmm$ <> " " Then
            Spacee$ = Spacee$ & TakeText$ & MakeIt
        Else
            Spacee$ = Spacee$ & TakeText$
        End If
    Next i
    OwnText = Spacee$
End Function

Public Function EchoBot(WhoToEcho As String)
    'this requires any version of
    'dos's ChatScan
    'dont forget Chat#.ScanOn
    Dim Screen_Name As String, What_Said As String
    If Screen_Name = WhoToEhcho Then
        ChatSend What_Said
    End If
End Function

Function SlashText(Strin As String)
    For i = 1 To Len(Strin)
        l$ = Mid(Strin, i, 1)
        l2$ = Mid(Strin, i + 1, 1)
        If l$ = " " Then
            Pd$ = Pd$ & l$
        ElseIf l2$ <> " " Then
            Pd$ = Pd$ & l$ & "/"
        Else
            Pd$ = Pd$ & l$
        End If
    Next i
    SlashText = Pd$
End Function

Public Sub HideAolToolBar()
    Dim AOLFrame As Long, AOLToolbar As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_AOL_Toolbar", vbNullString)
    X = SendMessageLong(AOLToolbar, WM_HIDE, 0&, 0&)
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As listbox, ListB As listbox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub

Public Sub Load2listboxes(Directory As String, ListA As listbox, ListB As listbox)
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        aString$ = Left(MyString$, InStr(MyString$, "*") - 1)
        bString$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
    Wend
    Close #1
End Sub

Public Sub ShowAolToolBar()
    Dim TheCount As Long
    Dim AOLFrame As Long, AOLToolbar As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_AOL_Toolbar", vbNullString)
    X = SendMessageLong(AOLToolbar, WM_SHOW, 0&, 0&)
End Sub

Public Sub DisableStart()
    Dim ShellTrayWnd As Long, Button As Long
    ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
    Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
    EnableWindow Button, 0
End Sub

Public Sub EnableStart()
    Dim ShellTrayWnd As Long, Button As Long
    ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
    Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
    EnableWindow Button, 1
End Sub

Public Sub DisableTrayWnd()
    ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
    EnableWindow ShellTrayWnd, 0
End Sub

Public Sub EnableTrayWnd()
    ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
    EnableWindow ShellTrayWnd, 1
End Sub

Public Function Attention(ToSend As String)
    ChatSend "--------Attention--------"
    ChatSend ToSend
    ChatSend "--------Attention--------"
End Function

Public Sub StartRun()
    Dim ShellTrayWnd As Long, Edit As Long, Button As Long
    ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
    Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
    PostMessage Button, &H201, 0&, 0&
    PostMessage Button, &H202, 0&, 0&
    PostMessage Button, &H102, 82, 0&
End Sub

Public Sub AddressBook()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    PostMessage AOLIcon, &H201, 0&, 0&
    PostMessage AOLIcon, &H202, 0&, 0&
    PostMessage AOLIcon, &H102, 65, 0&
End Sub

Public Sub MyAOLPrefs()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 5
        AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    PostMessage AOLIcon, &H201, 0&, 0&
    PostMessage AOLIcon, &H202, 0&, 0&
    PostMessage AOLIcon, &H102, 80, 0&
End Sub

Public Function LoopAttention(ToSend As String, Timer As Timer)
    Do
        Times = Val(Times + 1)
        ChatSend "--------Attention--------"
        ChatSend ToSend
        ChatSend "attentioned " & Times
        Wait 60
    Loop Until Timer.Enabled = False
End Function

Public Function AddFontsToList(List As listbox)
    For i = 0 To Screen.FontCount
        List.AddItem Screen.Fonts(i)
    Next
End Function

Public Sub AddFormToAolMdi(Form As Form)
    Dim AOL, MDI As Long
    AOL = FindWindow("AOL Frame25", vbNullString)
    MDI = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
    Call SetParent(Form.hwnd, MDI)
End Sub

Public Sub MakeAolYourForm(Form As Form)
    Dim AOL As Long
    AOL = FindWindow("AOL Frame25", vbNullString)
    Call SetParent(AOL, Form.hwnd)
End Sub

Public Sub ShutDownComp()
    SendMessage 0&, EWX_SHUTDOWN, 0&, 0&
End Sub

Public Sub Loadlistbox(Directory As String, TheList As listbox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub

Public Sub SendMail(Person As String, Subject As String, message As String)
    'copied sub from dos
    Dim AOL As Long, MDI As Long, tool As Long
    Dim Toolbar As Long, ToolIcon As Long, OpenSend As Long
    Dim DoIt As Long, Rich As Long, EditTo As Long
    Dim EditCC As Long, EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        X = ShowWindow(OpenSend&, 0)
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
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
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Wait 0.2
End Sub

Public Sub OpenMailAddNoSend(Person As String, Subject As String, message As String)
    Dim AOL As Long, MDI As Long, tool As Long
    Dim Toolbar As Long, ToolIcon As Long, OpenSend As Long
    Dim DoIt As Long, Rich As Long, EditTo As Long
    Dim EditCC As Long, EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
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
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
End Sub

Public Sub SaveListBox(Directory As String, TheList As listbox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Sub FormDrag(Form As Form)
    DoEvents
    ReleaseCapture
    keeb& = SendMessage(Form.hwnd, &HA1, 2, 0)
End Sub

Sub ChangAolCap(Text As String)
    Dim AOL As Long, X As Long
    AOL = FindWindow("AOL Frame25", vbNullString)
    X = SendMessageByString(AOL, WM_SETTEXT, 0, Text)
End Sub

Public Function FindRecycling() As Long
    Dim CabinetWClass As Long, Worker As Long, msctlsstatusbar As Long
    Dim SHELLDLLDefView As Long
    CabinetWClass = FindWindow("CabinetWClass", vbNullString)
    Worker = FindWindowEx(CabinetWClass, 0&, "Worker", vbNullString)
    msctlsstatusbar = FindWindowEx(CabinetWClass, 0&, "msctls_statusbar32", vbNullString)
    SHELLDLLDefView = FindWindowEx(CabinetWClass, 0&, "SHELLDLL_DefView", vbNullString)
    If Worker <> 0& And msctlsstatusbar <> 0& And SHELLDLLDefView <> 0& Then
        FindRecycling = CabinetWClass
        Exit Function
    Else
        While CabinetWClass
            CabinetWClass = GetWindow(CabinetWClass, 2)
            Worker = FindWindowEx(CabinetWClass, 0&, "Worker", vbNullString)
            msctlsstatusbar = FindWindowEx(CabinetWClass, 0&, "msctls_statusbar32", vbNullString)
            SHELLDLLDefView = FindWindowEx(CabinetWClass, 0&, "SHELLDLL_DefView", vbNullString)
            If Worker <> 0& And msctlsstatusbar <> 0& And SHELLDLLDefView <> 0& Then
                FindRecycling = CabinetWClass
                Exit Function
            End If
        Wend
    End If
End Function

Sub ChatRoom_ChngCapt(ToWhat As String)
    Room = FindRoom
    SendMessage Room, 0&, WM_SETTEXT, ToWhat
End Sub

Public Sub AOL25_RoomCaptChng(ToWhat As String)
    Room = AOL25_FindRoom
    SendMessage Room, 0&, WM_SETTEXT, ToWhat
End Sub

Public Sub SetText(hwnd As Long, TheText As String)
    Call SendMessageByString(hwnd&, WM_SETTEXT, 0&, TheText$)
End Sub

Function GetText(child)
    GetTrim = SenditbyNum(child, 14, 0&, 0&)
    TrimSpace$ = Space$(GetTrim)
    GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
    GetText = TrimSpace$
End Function

Sub Favorite_Places()
    Dim MainParent As Long, Parent1 As Long, Parent2 As Long
    Dim child As Long
    MainParent& = FindWindow("AOL Frame25", vbNullString)
    Parent1& = FindWindowEx(MainParent&, 0&, "AOL Toolbar", vbNullString)
    Parent2& = FindWindowEx(Parent1&, 0&, "_AOL_Toolbar", vbNullString)
    child& = FindWindowEx(Parent2&, 0, "_AOL_Icon", vbNullString)
    For X = 1 To 6
        child& = FindWindowEx(Parent2&, child&, "_AOL_Icon", vbNullString)
    Next X
    TheWindow& = child&
    PostMessage TheWindow&, &H201, 0&, 0&
    PostMessage TheWindow&, &H202, 0&, 0&
    PostMessage TheWindow&, &H102, s, 0&
End Sub

Public Sub ChangeChatCapt(txt As String)
    Dim X As Long
    If FindRoom Then
        X = SendMessageByString(stat&, WM_SETTEXT, 0, txt)
    End If
End Sub

Public Function EliteTalker(Word$)
    Made$ = ""
    For q = 1 To Len(Word$)
        Letter$ = ""
        Letter$ = Mid$(Word$, q, 1)
        leet$ = ""
        X = Int(Rnd * 3 + 1)
        If Letter$ = "a" Then
            If X = 1 Then leet$ = "â"
            If X = 2 Then leet$ = "å"
            If X = 3 Then leet$ = "ä"
        End If
        If Letter$ = "b" Then leet$ = "b"
        If Letter$ = "c" Then leet$ = "ç"
        If Letter$ = "d" Then leet$ = "d"
        If Letter$ = "e" Then
            If X = 1 Then leet$ = "ë"
            If X = 2 Then leet$ = "ê"
            If X = 3 Then leet$ = "é"
        End If
        If Letter$ = "i" Then
            If X = 1 Then leet$ = "ì"
            If X = 2 Then leet$ = "ï"
            If X = 3 Then leet$ = "î"
        End If
        If Letter$ = "j" Then leet$ = ",j"
        If Letter$ = "n" Then leet$ = "ñ"
        If Letter$ = "o" Then
            If X = 1 Then leet$ = "ô"
            If X = 2 Then leet$ = "ð"
            If X = 3 Then leet$ = "õ"
        End If
        If Letter$ = "s" Then leet$ = "š"
        If Letter$ = "t" Then leet$ = "†"
        If Letter$ = "u" Then
            If X = 1 Then leet$ = "ù"
            If X = 2 Then leet$ = "û"
            If X = 3 Then leet$ = "ü"
        End If
        If Letter$ = "w" Then leet$ = "vv"
        If Letter$ = "y" Then leet$ = "ÿ"
        If Letter$ = "0" Then leet$ = "Ø"
        If Letter$ = "A" Then
            If X = 1 Then leet$ = "Å"
            If X = 2 Then leet$ = "Ä"
            If X = 3 Then leet$ = "Ã"
        End If
        If Letter$ = "B" Then leet$ = "ß"
        If Letter$ = "C" Then leet$ = "Ç"
        If Letter$ = "D" Then leet$ = "Ð"
        If Letter$ = "E" Then leet$ = "Ë"
        If Letter$ = "I" Then
            If X = 1 Then leet$ = "Ï"
            If X = 2 Then leet$ = "Î"
            If X = 3 Then leet$ = "Í"
        End If
        If Letter$ = "N" Then leet$ = "Ñ"
        If Letter$ = "O" Then leet$ = "Õ"
        If Letter$ = "S" Then leet$ = "Š"
        If Letter$ = "U" Then leet$ = "Û"
        If Letter$ = "W" Then leet$ = "VV"
        If Letter$ = "Y" Then leet$ = "Ý"
        If Letter$ = "`" Then leet$ = "´"
        If Letter$ = "!" Then leet$ = "¡"
        If Letter$ = "?" Then leet$ = "¿"
        If Len(leet$) = 0 Then leet$ = Letter$
        Made$ = Made$ & leet$
    Next q
    EliteTalker = Made$
End Function

Function ChatSendBox()
    Dim Room As Long
    Dim Rich As Long
    Room = FindRoom()
    Rich = FindWindowEx(Room, 0, "RICHCNTL", vbNullString)
    Rich = FindWindowEx(Room, Rich, "RICHCNTL", vbNullString)
    ChatSendBox = Rich
End Function
Public Sub ChatSend(Text As String)
Room = FindRoom
RICHCNTL = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
RICHCNTL = FindWindowEx(Room, RICHCNTL, "RICHCNTL", vbNullString)
SendMessage RICHCNTL, 0&, WM_SETTEXT, ""
SendMessage RICHCNTL, 0&, WM_SETTEXT, Text

End Sub
Public Sub ChatSend2(Text As String)
Room = FindRoom
RICHCNTL = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
RICHCNTL = FindWindowEx(Room, RICHCNTL, "RICHCNTL", vbNullString)
SendMessage RICHCNTL, 0&, WM_SETTEXT, Text
End Sub
Public Function Add2Nums(Number1 As String, Number2 As String)
    Dim X As Long
    X = Val(Number1 + Number2)
    Add2Nums = X
End Function

Public Sub Enable_Window(WhichOne As String)
    EnableWindow WhichOne, 0
End Sub

Public Sub DisableWindow(WhichOne As String)
    EnableWindow WhichOne, 1
End Sub

Public Sub ClickMsgBox(WhatMB As String)
    PostMessage WhatMB, 0&, WM_LBUTTONDOWN, 0&
    PostMessage WhatMB, 0&, WM_LBUTTONUP, 0&
End Sub

Public Sub Add_Port(Port As String, Monitor As String)
    'you need to know the Monitors name
    'to have this work
    AddPort Port, 0&, Monitor
End Sub
Public Sub LogManager()
Dim AOLIcon As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)

AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 4
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
MenuIcon (AOLIcon)
PostMessage AOLIcon, WM_CHAR, 76, 0&
End Sub
Public Sub MyAOL()
AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)

AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 5
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
MenuIcon (AOLIcon)
PostMessage AOLIcon, WM_CHAR, 77, 0&
End Sub
Public Function AOL25_Plaza()
Room = AOL25_FindRoom
AOLIcon = FindWindowEx(Room, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 3
    AOLIcon = FindWindowEx(Room, AOLIcon, "_AOL_Icon", vbNullString)
Next
ClickIcon (AOLIcon)
End Function
Public Function FindSignOn() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim RICHCNTL As Long, AOLStatic As Long, AOLCombobox As Long
    Dim AOLEdit As Long, AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
    AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If RICHCNTL <> 0& And AOLStatic <> 0& And AOLCombobox <> 0& And AOLEdit <> 0& And AOLIcon <> 0& Then
        FindSignOn = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
            AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If RICHCNTL <> 0& And AOLStatic <> 0& And AOLCombobox <> 0& And AOLEdit <> 0& And AOLIcon <> 0& Then
                FindSignOn = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Public Sub AccessNumbers()
'must be signed off with aol sign
'on screen shown
Acces = FindSignOn
AOLIcon = FindWindowEx(Acces, 0&, "_AOL_Icon", vbNullString)
ClickIcon (AOLIcon)
End Sub
Public Sub AccessNumbers2(Num As String)
'must be signed off with aol sign
'on screen shown
'same as before except it gives the
'users area code
Acces = FindSignOn
AccesNum = FindAccessNums
AOLIcon = FindWindowEx(Acces, 0&, "_AOL_Icon", vbNullString)
ClickIcon (AOLIcon)
Wait 1
AOLEdit = FindWindowEx(AccessNum, 0&, "_AOL_Edit", vbNullString)
SendMessage AOLEdit, 0&, WM_SETTEXT, Nums
SendMessage AOLEdit, WM_CHAR, 13, 0&
End Sub
Public Sub AOL25_SendMail(SN As String, Subject As String, Mess As String)
AOL25_ClickMail
Wait 3
Compo = AOL25_FindCompose
AOLEdit = FindWindowEx(Compo, 0&, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit, 0&, WM_SETTEXT, SN
AOLEdit2 = FindWindowEx(Compo, 0&, "_AOL_Edit", vbNullString)
AOLEdit2 = FindWindowEx(Compo, AOLEdit2, "_AOL_Edit", vbNullString)
AOLEdit2 = FindWindowEx(Compo, AOLEdit2, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit2, 0&, WM_SETTEXT, Subject
AOLEdit3 = FindWindowEx(Compo, 0&, "_AOL_Edit", vbNullString)
For i = 1 To 3
    AOLEdit3 = FindWindowEx(Compo, AOLEdit3, "_AOL_Edit", vbNullString)
Next
SendMessageByString AOLEdit3, 0&, WM_SETTEXT, Mess
End Sub
Public Function AOL25_FindCompose() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLStatic As Long, AOLIcon As Long, AOLEdit As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
    If AOLStatic <> 0& And AOLIcon <> 0& And AOLEdit <> 0& Then
        AOL25_FindCompose = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
            If AOLStatic <> 0& And AOLIcon <> 0& And AOLEdit <> 0& Then
                AOL25_FindCompose = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function

Public Sub AOL25_ClickMail()
AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
     AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
 ClickIcon (AOLIcon)
End Sub
Public Function FindAccessNums() As Long
    Dim AOLFrame As Long, AOLModal As Long, AOLGlyph As Long
    Dim AOLStatic As Long, RICHCNTL As Long, AOLEdit As Long
    Dim AOLCombobox As Long, AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLModal = FindWindowEx(AOLFrame, 0&, "_AOL_Modal", vbNullString)

    AOLGlyph = FindWindowEx(AOLModal, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic = FindWindowEx(AOLModal, 0&, "_AOL_Static", vbNullString)
    RICHCNTL = FindWindowEx(AOLModal, 0&, "RICHCNTL", vbNullString)
    AOLEdit = FindWindowEx(AOLModal, 0&, "_AOL_Edit", vbNullString)
    AOLCombobox = FindWindowEx(AOLModal, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon = FindWindowEx(AOLModal, 0&, "_AOL_Icon", vbNullString)
    If AOLGlyph <> 0& And AOLStatic <> 0& And RICHCNTL <> 0& And AOLEdit <> 0& And AOLCombobox <> 0& And AOLIcon <> 0& Then
        FindAccessNums = AOLModal
        Exit Function
    Else
        While AOLModal
            AOLModal = GetWindow(AOLModal, 2)
            AOLGlyph = FindWindowEx(AOLModal, 0&, "_AOL_Glyph", vbNullString)
            AOLStatic = FindWindowEx(AOLModal, 0&, "_AOL_Static", vbNullString)
            RICHCNTL = FindWindowEx(AOLModal, 0&, "RICHCNTL", vbNullString)
            AOLEdit = FindWindowEx(AOLModal, 0&, "_AOL_Edit", vbNullString)
            AOLCombobox = FindWindowEx(AOLModal, 0&, "_AOL_Combobox", vbNullString)
            AOLIcon = FindWindowEx(AOLModal, 0&, "_AOL_Icon", vbNullString)
            If AOLGlyph <> 0& And AOLStatic <> 0& And RICHCNTL <> 0& And AOLEdit <> 0& And AOLCombobox <> 0& And AOLIcon <> 0& Then
                FindAccessNums = AOLModal
                Exit Function
            End If
        Wend
    End If
End Function
Public Function Minus2Nums(Number1 As String, Number2 As String)
    Dim X As Long
    X = Val(Number1 - Number2)
    Minus2Nums = X
End Function

Public Function Mult2Nums(Number1 As String, Number2 As String)
    Dim X As Long
    X = Val(Number1 * Number2)
    Mult2Nums = X
End Function

Public Function Div2Nums(Number1 As String, Number2 As String)
    Dim X As Long
    If B > 0 Then
        X = Val(Number1 / Number2)
    Else
        MsgBox "you dumbass... you cant divide by zero!"
    End If
    Div2Nums = X
End Function

Public Function CenterForm(Form As Form)
    Form.Top = Screen.Width / 2
    Form.Left = Screen.Height / 2
End Function

Public Sub AIM_SignOn(SN As String, PASS As String)
    Aim = FindAimSignOn
    ComboBox = FindWindowEx(Aim, 0&, "ComboBox", vbNullString)
    Edit = FindWindowEx(ComboBox, 0&, "Edit", vbNullString)
    SendMessageByString Edit, WM_SETTEXT, 0&, SN
    Edit = FindWindowEx(Edit, 0&, "Edit", vbNullString)
    SendMessageByString Edit, WM_SETTEXT, 0&, PASS
    Staticc = FindWindowEx(TheWindow, 0&, "Static", vbNullString)
    For i = 1 To 4
        Staticc = FindWindowEx(TheWindow, Staticc, "Static", vbNullString)
    Next
    ClickIcon (Staticc)
End Sub

Public Function FindAimSignOn() As Long
    Dim TheWindow As Long, Staticc As Long, Button As Long
    Dim ComboBox As Long, Edit As Long, OscarSeparator As Long
    Dim WndAteClass As Long, OscarIconBtn As Long
    TheWindow = FindWindow("#32770", vbNullString)
    Staticc = FindWindowEx(TheWindow, 0&, "Static", vbNullString)
    Button = FindWindowEx(TheWindow, 0&, "Button", vbNullString)
    ComboBox = FindWindowEx(TheWindow, 0&, "ComboBox", vbNullString)
    Edit = FindWindowEx(TheWindow, 0&, "Edit", vbNullString)
    OscarSeparator = FindWindowEx(TheWindow, 0&, "_Oscar_Separator", vbNullString)
    WndAteClass = FindWindowEx(TheWindow, 0&, "WndAte32Class", vbNullString)
    OscarIconBtn = FindWindowEx(TheWindow, 0&, "_Oscar_IconBtn", vbNullString)
    If Staticc <> 0& And Button <> 0& And ComboBox <> 0& And Edit <> 0& And OscarSeparator <> 0& And WndAteClass <> 0& And OscarIconBtn <> 0& Then
        FindAimSignOn = TheWindow
        Exit Function
    Else
        While TheWindow
            TheWindow = GetWindow(TheWindow, 2)
            Staticc = FindWindowEx(TheWindow, 0&, "Static", vbNullString)
            Button = FindWindowEx(TheWindow, 0&, "Button", vbNullString)
            ComboBox = FindWindowEx(TheWindow, 0&, "ComboBox", vbNullString)
            Edit = FindWindowEx(TheWindow, 0&, "Edit", vbNullString)
            OscarSeparator = FindWindowEx(TheWindow, 0&, "_Oscar_Separator", vbNullString)
            WndAteClass = FindWindowEx(TheWindow, 0&, "WndAte32Class", vbNullString)
            OscarIconBtn = FindWindowEx(TheWindow, 0&, "_Oscar_IconBtn", vbNullString)
            If Staticc <> 0& And Button <> 0& And ComboBox <> 0& And Edit <> 0& And OscarSeparator <> 0& And WndAteClass <> 0& And OscarIconBtn <> 0& Then
                FindAimSignOn = TheWindow
                Exit Function
            End If
        Wend
    End If
End Function

Public Sub AddChatRichtnlToForm(Form As Form)
    Room = FindRoom
    RICHCNTL = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    Call SetParent(RICHCNTL, Form.hwnd)
End Sub

Public Sub ClickPrint()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 3
        AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    ClickIcon (AOLIcon)
End Sub
Public Function AOLMDI()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLMDI = MDIClient
End Function
Public Function AOL25_MDI()
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOL25_MDI = MDIClient
End Function
Public Sub PeopleHereChng(ToWhat As String)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    For i = 1 To 3
        AOLStatic = FindWindowEx(AOLChild, AOLStatic, "_AOL_Static", vbNullString)
    Next
    SendMessageByString AOLStatic, WM_SETTEXT, 0&, ToWhat
End Sub

Public Sub Ghost_On()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    ClickIcon (AOLIcon)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 4
        AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    ClickIcon (AOLIcon)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLCheckbox = FindWindowEx(AOLChild, 0&, "_AOL_Checkbox", vbNullString)
    For i = 1 To 4
        AOLCheckbox = FindWindowEx(AOLChild, AOLCheckbox, "_AOL_Checkbox", vbNullString)
    Next
    ClickIcon (AOLCheckbox)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 3
        AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    ClickIcon (AOLIcon)
    Dim AOLFrame As Long, TheWindow As Long, Button As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    TheWindow = FindWindowEx(AOLFrame, 0&, "#32770", vbNullString)
    Button = FindWindowEx(TheWindow, 0&, "Button", vbNullString)
    PostMessage Button, &H201, 0&, 0&
    PostMessage Button, &H202, 0&, 0&
End Sub

Public Sub Ghost_Off()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    ClickIcon (AOLIcon)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 4
        AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    ClickIcon (AOLIcon)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLCheckbox = FindWindowEx(AOLChild, 0&, "_AOL_Checkbox", vbNullString)
    For i = 1 To 1
        AOLCheckbox = FindWindowEx(AOLChild, AOLCheckbox, "_AOL_Checkbox", vbNullString)
    Next
    ClickIcon (AOLCheckbox)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 3
        AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    ClickIcon (AOLIcon)
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    TheWindow = FindWindowEx(AOLFrame, 0&, "#32770", vbNullString)
    Button = FindWindowEx(TheWindow, 0&, "Button", vbNullString)
    PostMessage Button, &H201, 0&, 0&
    PostMessage Button, &H202, 0&, 0&
End Sub

Public Sub AddAolToFrom(Form As Form)
    AOL = FindWindow("Aol Frame25", vbNullString)
End Sub

Public Function OnlineClockSend()
    Keyword "online clock"
    Wait 0.5
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLModal = FindWindowEx(AOLFrame, 0&, "_AOL_Modal", vbNullString)
    AOLStatic = FindWindowEx(AOLModal, 0&, "_AOL_Static", vbNullString)
    TheLength = SendMessageLong(AOLStatic, &HE, 0&, 0&)
    TheText = String(TheLength, 0)
    SendMessageByString AOLStatic, &HD, TheLength, TheText
    ChatSend TheText
End Function

Public Sub AOL25_ChangeRMNum(ToWhat As String)
    Room = AOL25_FindRoom
    AOLStatic = FindWindowEx(Room, 0&, "_AOL_Static", vbNullString)
    AOLStatic = FindWindowEx(Room, AOLStatic, "_AOL_Static", vbNullString)
    AOLStatic = FindWindowEx(Room, AOLStatic, "_AOL_Static", vbNullString)
    SendMessage AOLStatic, 0&, WM_SETTEXT, ToWhat
End Sub

Public Function FindIM() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLView As Long, AOLStatic As Long, RICHCNTL As Long
    Dim AOLIcon As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AOLView <> 0& And AOLStatic <> 0& And RICHCNTL <> 0& And AOLIcon <> 0& Then
        FindIM = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AOLView <> 0& And AOLStatic <> 0& And RICHCNTL <> 0& And AOLIcon <> 0& Then
                FindIM = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function

Public Sub AddImToPictureBox(Pic As PictureBox)
    IM = FindIM
    Call SetParent(IM, Pic.hwnd)
End Sub

Public Sub ClearChat()
    Room = FindRoom
    RICHCNTL = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    SendMessage RICHCNTL, WM_CLEAR, 0&, 0&
End Sub

Public Sub AOL25_Keyword(ToWhere As String)
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLEdit As Long
    AOL = FindAol25
    AppActivate "America  Online"
    SendKeys "^{k}"
    AOLChild = FindWindowEx(AOL, 0&, "AOL Child", vbNullString)
    AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", "Keyword ")
    PostMessage AOLEdit, &HC, 0&, ToWhere
    Wait 0.5
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (AOLIcon)
End Sub

Public Function AddFontsToCombo(Combo As ComboBox)
    For i = 0 To Screen.FontCount
        Combo.AddItem Screen.Fonts(i)
    Next
End Function

Public Sub MailPrefs()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    MenuIcon (AOLIcon)
    PostMessage AOLIcon, WM_CHAR, 80, 0&
End Sub

Public Sub MenuIcon(Ico As String)
    PostMessage Ico, WM_LBUTTONDOWN, 0&, 0&
    PostMessage Ico, WM_LBUTTONUP, 0&, 0&
End Sub
Public Sub ClickIcon(Ico As String)
    SendMessage Ico, WM_LBUTTONDOWN, 0&, 0&
    SendMessage Ico, WM_LBUTTONUP, 0&, 0&
End Sub
Public Sub AIM_Hide()
    Dim Aim As Long
    Aim = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Aim = ShowWindow(Aim, 0)
End Sub

Public Sub AIM_Show()
    Dim Aim As Long
    Aim = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Aim = ShowWindow(Aim, 1)
End Sub

Public Function ObjectMoveDown(object)
    object.Top = Val(Obj.Top) - 20
End Function

Function ObjectMoveUp(object)
    object.Top = Val(Obj.Top) + 20
End Function

Function SlideObject(Form As Form, object As Object)
    'copied from hound
    If object.Left <= 0 Then
        Do Until (object.Left + object.Width) >= Frm.Width
            object.Move Val(object.Left) + 55, object.Top
            Wait (0.01)
        Loop
    End If
End Function

Public Function ReplaceString(Str As String, StrToReplace As String, WithWhat As String)
    ReplaceString = Replace(Str, StrToReplace, ReplaceWith)
End Function

Function RandomBackColor(Pic As PictureBox)
    R = Int(Rnd * 255) + 1
    g = Int(Rnd * 255) + 1
    B = Int(Rnd * 255) + 1
    Pic.BackColor = RGB(R, g, B)
End Function

Public Function CountMail()
    Mail = FindMailBox
    AOLTabControl = FindWindowEx(Mail, 0&, "_AOL_TabControl", vbNullString)
    AOLTabPage = FindWindowEx(AOLTabControl, 0&, "_AOL_TabPage", vbNullString)
    AOLTree = FindWindowEx(AOLTabPage, 0&, "_AOL_Tree", vbNullString)
    TheCount = SendMessageLong(AOLTree, &H18B, 0&, 0&)
    SendMessage Mail, 0&, WM_CLOSE, 0&
    TheCount = CountMail
End Function

Public Sub NewMail()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 1
        AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    Do Until FindMailBox
        ClickIcon AOLIcon
        Wait 0.5
    Loop
End Sub

Public Function FindMailBox() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLGlyph As Long, AOLStatic As Long, AOLImage As Long
    Dim AOLTabControl As Long, AOLIcon As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
    AOLTabControl = FindWindowEx(AOLChild, 0&, "_AOL_TabControl", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AOLGlyph <> 0& And AOLStatic <> 0& And AOLImage <> 0& And AOLTabControl <> 0& And AOLIcon <> 0& Then
        FindMailBox = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
            AOLTabControl = FindWindowEx(AOLChild, 0&, "_AOL_TabControl", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AOLGlyph <> 0& And AOLStatic <> 0& And AOLImage <> 0& And AOLTabControl <> 0& And AOLIcon <> 0& Then
                FindMailBox = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function

Function RandomNumber(LastNumber)
    keeb = Int(Rnd * LastNumber) + 1
    RandomNumber = keeb
End Function

Function SquareRootNum(Num As String)
    Dim X As Long
    X = Sqr(Num)
    SquareRootNum = X
End Function

Function LowerCaseText(Text As String)
    Text = LCase(Text)
    Text = LowerCaseText
End Function

Function UpperCaseText(Text As String)
    Text = UCase(Text)
    Text = UpperCaseText
End Function
Public Sub Wait(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Function CopyAString(What As String)
    Clipboard.SetText What
End Function

Public Function PasteAString()
    Dim What As String
    What = Clipboard.GetText
    What = PasteAString
End Function

Public Function TimeNow()
    TimeNow = Time
End Function

Public Function DateNow()
    DateNow = Date
End Function

Public Sub DownloadManager()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 4
        AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    MenuIcon (AOLIcon)
    SendChar AOLIcon, 60
End Sub

Public Sub OfflineMail()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 4
        AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    MenuIcon (AOLIcon)
    SendChar AOLIcon, 110
End Sub

Public Sub PersonalFilingCabinet()
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 4
        AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
    Next
    MenuIcon (AOLIcon)
    SendChar AOLIcon, 80
End Sub

Public Sub SendChar(Icon As String, WhichChr As String)
    SendMessage Icon, WM_CHAR, WhichChr, 0&
End Sub

Public Function DateNow2()
    DateNow2 = Replace(Date, "/", "-")
End Function

Public Function DateNow3(Chnge As String)
    DateNow3 = Replcace(Date, "/", Chnge)
End Function

Public Function TimeNow2()
    If Len(Time) = "11" Then
        TimeNow2 = Left(Time, 5)
    Else
        TimeNow2 = Left(Time, 4)
    End If
End Function

Function ScrambleText(txt)
    'copied sub from glue
    findlastspace = Mid(txt, Len(txt), 1)
    If Not findlastspace = " " Then
        txt = txt & " "
    Else
        txt = txt
    End If
    For scrambling = 1 To Len(txt)
        thechar$ = Mid(txt, scrambling, 1)
        Char$ = Char$ & thechar$
        If thechar$ = " " Then
            chars$ = Mid(Char$, 1, Len(Char$) - 1)
            firstchar$ = Mid(chars$, 1, 1)
            On Error GoTo gods
            lastchar$ = Mid(chars$, Len(chars$), 1)
            midchar$ = Mid(chars$, 2, Len(chars$) - 2)
            For SpeedBack = Len(midchar$) To 1 Step -1
                backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
            Next SpeedBack
            GoTo meeh
gods:
            scrambled$ = scrambled$ & firstchar$ & " "
            GoTo Fuck
meeh:
            scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & ""
Fuck:
            Char$ = ""
            backchar$ = ""
        End If
    Next scrambling
    Text_Scramble = scrambled$
    Exit Function
End Function

Sub AddRoomList(listbox As listbox)
    Dim AOLProcess As Long, ListItemHold As Long, Person As String
    Dim ListPersonHold As Long, ReadBytes As Long
    On Error Resume Next
    TheList.Clear
    Room = FindRoom()
    aolhandle = FindChildByClass(Room, "_AOL_Listbox")
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
            If Person$ = GetUser Then GoTo Na
            listbox.AddItem TrimSpaces(LCase(Person$))
Na:
        Next Index
        Call CloseHandle(AOLProcessThread)
    End If
End Sub

Sub AddRoomCombo(Comb As ComboBox)
    Dim AOLProcess As Long, ListItemHold As Long, Person As String
    Dim ListPersonHold As Long, ReadBytes As Long
    On Error Resume Next
    TheList.Clear
    Room = FindRoom()
    aolhandle = FindChildByClass(Room, "_AOL_Listbox")
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
            If Person$ = GetUser Then GoTo Na
            Comb.AddItem TrimSpaces(LCase(Person$))
Na:
        Next Index
        Call CloseHandle(AOLProcessThread)
    End If
End Sub

Function TrimSpaces(Text)
    If InStr(Text, " ") = 0 Then
        TrimSpaces = Text
        Exit Function
    End If
    For TrimSpace = 1 To Len(Text)
        thechar$ = Mid(Text, TrimSpace, 1)
        TheChars$ = TheChars$ & thechar$
        If thechar$ = " " Then
            TheChars$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
        End If
    Next TrimSpace
    TrimSpaces = TheChars$
End Function

Public Function LinkSend(URL As String, Mess As String)
    ChatSend "< a href=" + URL + ">" + Mess + "</a>"
End Function

Public Function DeleteFile(Path As String)
    Kill (Path)
End Function

Public Function LetterToChr(Letter As String)
    Ans = Asc(Letter)
    LetterToChr = Ans
End Function

Public Function SetAppFocus(App As String)
    AOL% = FindWindow("AOL Frame25", vbNullString)
    Call FocusSet(AOL%)
End Function

Function CreateFile(FileName As String)
    Free = FreeFile
    Open FileName For Random As Free
    Close Free
End Function

Function LoadTimes(ProgramName As String)
    On Error Resume Next
    keeb = GetSetting(ProgramName, "Load", "Times")
    keeb = Val(Num) + 1
    Call SaveSetting(ProgramName, "Load", "Times", Num)
    LoadTimes = keeb
End Function

Function FindChildByClass(Parent, child As String) As Integer
    childfocus& = GetWindow(Parent, 5)
    While childfocus&
        buffer$ = String$(250, 0)
        classbuffer& = GetClassName(childfocus&, buffer$, 250)
        If InStr(UCase(buffer$), UCase(child)) Then FindChildByClass = childfocus&: Exit Function
        childfocus& = GetWindow(childfocus&, 2)
    Wend
End Function

Function ChangeRoomNum(Num)
    Dim Chil As Long, Rich As Long, stat As Long
    Dim X
    Chil& = FindRoom
    Rich& = FindChildByClass(Chil&, "RICHCNTL")
    stat& = FindChildByClass(Chil&, "_AOL_Static")
    X = SendMessageByString(stat&, WM_SETTEXT, 0, Num)
End Function

Public Function FindAol4() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLToolbar As Long
    Dim AOLMMI As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLMMI = FindWindowEx(AOLFrame, 0&, "_AOL_MMI", vbNullString)
    If MDIClient <> 0& And AOLToolbar <> 0& And AOLMMI <> 0& Then
        FindAol4 = AOLFrame
        Exit Function
    Else
        While AOLFrame
            AOLFrame = GetWindow(AOLFrame, 2)
            MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
            AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
            AOLMMI = FindWindowEx(AOLFrame, 0&, "_AOL_MMI", vbNullString)
            If MDIClient <> 0& And AOLToolbar <> 0& And AOLMMI <> 0& Then
                FindAol4 = AOLFrame
                Exit Function
            End If
        Wend
    End If
End Function

Public Function FindAol25() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLToolbar As Long
    Dim AOLMMI As Long
    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLMMI = FindWindowEx(AOLFrame, 0&, "_AOL_MMI", vbNullString)
    If MDIClient <> 0& And AOLToolbar <> 0& And AOLMMI <> 0& Then
        FindAol25 = AOLFrame
        Exit Function
    Else
        While AOLFrame
            AOLFrame = GetWindow(AOLFrame, 2)
            MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
            AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
            AOLMMI = FindWindowEx(AOLFrame, 0&, "_AOL_MMI", vbNullString)
            If MDIClient <> 0& And AOLToolbar <> 0& And AOLMMI <> 0& Then
                FindAol25 = AOLFrame
                Exit Function
            End If
        Wend
    End If
End Function

Sub CloseRoom()
    Room = FindRoom
    SendMessage Room, 0&, WM_CLOSE, 0&
End Sub

Function FindRoom()
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindIt(AOL&, "MDIClient")
    firs& = GetWindow(MDI&, 5)
    listers& = FindIt(firs&, "RICHCNTL")
    listere& = FindIt(firs&, "RICHCNTL")
    listerb& = FindIt(firs&, "_AOL_Listbox")
    Do While (listers& = 0 Or listere& = 0 Or listerb& = 0) And (l <> 100)
        DoEvents
        firs& = GetWindow(firs&, 2)
        listers& = FindIt(firs&, "RICHCNTL")
        listere& = FindIt(firs&, "RICHCNTL")
        listerb& = FindIt(firs&, "_AOL_Listbox")
        If listers& And listere& And listerb& Then Exit Do
        l = l + 1
    Loop
    If (l < 100) Then
        FindRoom = firs&
        Exit Function
    End If
End Function

Function GetCaption(hwnd)
    hwndLength& = GetWindowTextLength(hwnd)
    hwndTitle$ = String$(hwndLength&, 0)
    Qo0& = GetWindowText(hwnd, hwndTitle$, (hwndLength& + 1))
    GetCaption = hwndTitle$
End Function

Function FindItsTitle(parentw, childhand)
    Num1& = GetWindow(parentw, 5)
    If UCase(GetCaption(Num1&)) Like UCase(childhand) Then GoTo god
    Num1& = GetWindow(parentw, GW_CHILD)
    While Num1&
        Num2& = GetWindow(parentw, 5)
        If UCase(GetCaption(Num2&)) Like UCase(childhand) & "*" Then GoTo god
        Num1& = GetWindow(Num1&, 2)
        If UCase(GetCaption(Num1&)) Like UCase(childhand) & "*" Then GoTo god
    Wend
    FindItsTitle = 0
god:
    Qo0& = Num1&
    FindItsTitle = Qo0&
End Function

Function FindChildByTitle(Parent, child As String) As Integer
    childfocus& = GetWindow(Parent, 5)
    While childfocus&
        hwndLength& = GetWindowTextLength(childfocus&)
        buffer$ = String$(hwndLength&, 0)
        WindowText& = GetWindowText(childfocus&, buffer$, (hwndLength& + 1))
        If InStr(UCase(buffer$), UCase(child)) Then FindChildByTitle = childfocus&: Exit Function
        childfocus& = GetWindow(childfocus&, 2)
    Wend
End Function

Sub IMSend(Recipiant, message)
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindIt(AOL&, "MDIClient")
    Call Keyword("im")
    Do: DoEvents
        IMWin& = FindItsTitle(MDI&, "Send Instant Message")
        AOEdit& = FindIt(IMWin&, "_AOL_Edit")
        AORich& = FindIt(IMWin&, "RICHCNTL")
        AOIcon& = FindIt(IMWin&, "_AOL_Icon")
    Loop Until AOEdit& <> 0 And AORich& <> 0 And AOIcon& <> 0
    Call SenditByString(AOEdit&, WM_SETTEXT, 0, Recipiant)
    Call SenditByString(AORich&, WM_SETTEXT, 0, message)
    For X = 1 To 9
        AOIcon& = GetWindow(AOIcon&, 2)
    Next X
    Call Wait(0.01)
    ClickIcon (AOIcon&)
    Do: DoEvents
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindIt(AOL&, "MDIClient")
        IMWin& = FindItsTitle(MDI&, "Send Instant Message")
        OkWin& = FindWindow("#32770", "America Online")
        If OkWin& <> 0 Then Call SendMessage(OkWin&, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin&, WM_CLOSE, 0, 0): Exit Do
        If IMWin& = 0 Then Exit Do
    Loop
End Sub

Sub ChatBold()
    Dim Room As Long, digbold As Long
    Room = FindRoom()
    digbold = FindIt(Room, "_AOL_Icon")
    digbold = GetWindow(digbold, GW_HWNDNEXT)
    ClickIcon (digbold)
End Sub

Sub ChooseChatColor()
    Dim Room, digbold, Click As Long
    Room = FindRoom
    Colorr = FindIt(Room, "_AOL_Icon")
    ClickIcon (Colorr)
End Sub

Sub ChooseChatFont()
    Dim Room, digbold, clic As Long
    Room = FindRoom()
    ChatFont = FindIt(Room, "_AOL_ComboBox")
    ClickIcon ChatFont
End Sub

Sub ChatUnderline()
    Dim Room, digbold, clic As Long
    Room = FindRoom()
    digbold = FindIt(Room, "_AOL_Icon")
    digbold = GetWindow(digbold, GW_HWNDNEXT)
    digbold = GetWindow(digbold, GW_HWNDNEXT)
    digbold = GetWindow(digbold, GW_HWNDNEXT)
    ClickIcon digbold
End Sub

Sub ChatItalic()
    Dim Room, digbold, clic As Long
    Room = FindRoom()
    digbold = FindIt(Room, "_AOL_Icon")
    digbold = GetWindow(digbold, GW_HWNDNEXT)
    digbold = GetWindow(digbold, GW_HWNDNEXT)
    ClickIcon digbold
End Sub

Public Function AOL25_FindRoom() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLView As Long, AOLEdit As Long, AOLIcon As Long
    Dim AOLImage As Long, AOLStatic As Long, AOLListbox As Long
    Dim AOLGlyph As Long
    AOLFrame = FindWindow("AOL Frame25", "America  Online")
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
    AOLView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
    AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLListbox = FindWindowEx(AOLChild, 0&, "_AOL_Listbox", vbNullString)
    AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
    If AOLView <> 0& And AOLEdit <> 0& And AOLIcon <> 0& And AOLImage <> 0& And AOLStatic <> 0& And AOLListbox <> 0& And AOLGlyph <> 0& Then
        AOL25_FindRoom = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
            AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLListbox = FindWindowEx(AOLChild, 0&, "_AOL_Listbox", vbNullString)
            AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
            If AOLView <> 0& And AOLEdit <> 0& And AOLIcon <> 0& And AOLImage <> 0& And AOLStatic <> 0& And AOLListbox <> 0& And AOLGlyph <> 0& Then
                AOL25_FindRoom = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function

Sub AOL25_ChatSend(Text As String)
    Room& = AOL25_FindRoom
    AOLEdit = FindWindowEx(Room&, 0&, "_AOL_Edit", vbNullString)
    SendMessageByString AOLEdit, &HC, 0&, Text
    SendMessageByString AOLEdit, &H102, 13, 0&
End Sub

Function AOL25_GetchatText()
    Room& = AOL25_FindRoom
    AOLView = FindWindowEx(Room&, 0&, "_AOL_View", vbNullString)
    TheLength = SendMessageLong(AOLView, &HE, 0&, 0&)
    TheText = String(TheLength, 0)
    SendMessageByString AOLView, &HD, TheLength, TheText
    AOL25_GetchatText = TheText
End Function

Public Function AOL25_LastChatLineWithSN()
    ChatText& = AOL25_GetchatText
    For FindChar = 1 To Len(ChatText&)
        thechar$ = Mid(ChatText&, FindChar, 1)
        TheChars$ = TheChars$ & thechar$
        If thechar$ = Chr(13) Then
            TheChatText$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
            TheChars$ = ""
        End If
    Next FindChar
    lastlen = Val(FindChar) - Len(TheChars$)
    LastLine = Mid(ChatText&, lastlen, Len(TheChars$))
    AOL25_LastChatLineWithSN = LastLine
End Function

Function AOL25_ScreenNameFromLastChatLine()
    ChatText$ = AOL25_LastChatLineWithSN
    ChatTrim$ = Left$(ChatText$, 17)
    For z = 1 To 17
        If Mid$(ChatTrim$, z, 1) = ":" Then
            SN = Left$(ChatTrim$, z - 1)
        End If
    Next z
    SNFromLastChatLine = SN
End Function
Public Function AOL25_SingOnScrn() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLGlyph As Long, AOLStatic As Long, AOLCombobox As Long
    Dim AOLEdit As Long, AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
    AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AOLGlyph <> 0& And AOLStatic <> 0& And AOLCombobox <> 0& And AOLEdit <> 0& And AOLIcon <> 0& Then
        AOL25_SingOn = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
            AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AOLGlyph <> 0& And AOLStatic <> 0& And AOLCombobox <> 0& And AOLEdit <> 0& And AOLIcon <> 0& Then
                AOL25_SingOn = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function

Function AOL25_LastChatLine()
    ChatText = AOL25_LastChatLineWithSN
    ChatTrimNum = Len(AOL25_SNFromLastChatLine)
    ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
    LastChatLine = ChatTrim$
End Function

Function AOL25_ChatBox()
    Room = AOL25_FindRoom
    AOLEdit = FindWindowEx(Room, 0&, "_AOL_Edit", vbNullString)
End Function

Sub LinksInvisible()
    ChatSend "<body bgcolor=#fffffe link=#fffffe>"
End Sub

Sub LinksBlack()
    ChatSend "<body bgcolor=#fffffe link=#000000>"
End Sub

Sub LinksRed()
    ChatSend "<body bgcolor=#fffffe link=#0000FF>"
End Sub

Sub LinksBlue()
    ChatSend "<body bgcolor=#fffffe link=#00FF00>"
End Sub

Public Sub LinksCustom(ToWhat As String)
    ChatSend "<body link=" + ToWhat + ">"
End Sub

Sub MaxRoom()
    Room = FindRoom
    Call ShowWindow(Room, 3)
End Sub

Sub MinRoom()
    Room = FindRoom
    Call ShowWindow(Room&, 2)
End If

Room = AOL25_FindRoom
ShowWindow Room, 2

Room = AOL25_FindRoom
ShowWindow Room, 3

Room = AOL25_FindRoom
ShowWindow Room, 0

Room = AOL25_FindRoom
SendMessage Room, 0&, WM_CLOSE, 0&

Dim X As Long
X = Shell(Directory$ & "aol.exe", vbNormalFocus)

Dim X As Long
X = Shell("C:\Program Files\Winamp\Winamp.exe", vbNormalFocus)

Dim X As String
X = Shell(Directory$ & "aol.exe", vbNormalFocus)

'copied sub from dos
txt.Text = " " & txt.Text & Chr(13)
Dim i As Long
For i = 1 To Len(txt.Text)
    l$ = Mid(txt.Text, i, 1)
    If l$ = Chr(13) Then
        Call ChatSend(Mid(TLine$, 2, Len(TLine$))): TLine$ = "": l$ = ""
        Wait (0.5)
    End If
    TLine$ = TLine$ & l$
Next i
txt.Text = Mid(txt.Text, 2, Len(txt.Text) - 2)

txt.Text = " " & txt.Text & Chr(13)
Dim i As Long
For i = 1 To Len(txt.Text)
    l$ = Mid(txt.Text, i, 1)
    If l$ = Chr(13) Then
        Call AOL25_ChatSend(Mid(TLine$, 2, Len(TLine$))): TLine$ = "": l$ = ""
        Wait (0.5)
    End If
    TLine$ = TLine$ & l$
Next i
txt.Text = Mid(txt.Text, 2, Len(txt.Text) - 2)

Dim X As String
X = Shell(Directory$ & "aol.exe", vbNormalFocus)

Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long
Dim ATLFB As Long
Progman = FindWindow("Progman", vbNullString)
SHELLDLLDefView = FindWindowEx(Progman, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer = FindWindowEx(SHELLDLLDefView, 0&, "Internet Explorer_Server", vbNullString)
ATLFB = FindWindowEx(InternetExplorerServer, 0&, "ATL:79F6B760", vbNullString)
If ATLFB <> 0& Then
    FindDesktop = InternetExplorerServer
    Exit Function
Else
    While InternetExplorerServer
        InternetExplorerServer = GetWindow(InternetExplorerServer, 2)
        ATLFB = FindWindowEx(InternetExplorerServer, 0&, "ATL:79F6B760", vbNullString)
        If ATLFB <> 0& Then
            FindDesktop = InternetExplorerServer
            Exit Function
        End If
    Wend
End If

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "AOL Toolbar", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "_AOL_Toolbar", vbNullString)
For hand = 1 To 9
    ourhandle& = FindWindowEx(OurParent&, Hand9, "_AOL_Icon", vbNullString)
Next hand

buffer$ = String$(250, 0)
getclas& = GetClassName(child, buffer$, 250)
GetClass = buffer$

AOL& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindIt(AOL&, "AOL Toolbar")
ParHand3& = FindIt(ParHand2&, "_AOL_Toolbar")
OurParent& = FindIt(ParHand3&, "_AOL_Combobox")
ourhandle& = FindIt(OurParent&, "Edit")
Call SenditByString(ourhandle&, WM_SETTEXT, 0, txt)
Call SenditbyNum(ourhandle&, WM_CHAR, VK_SPACE, 0)
Call SenditbyNum(ourhandle&, WM_CHAR, VK_RETURN, 0)

Dim TheLength As Long, TheText As String
Dim OscarBuddyListWin As Long, SN As String
OscarBuddyListWin = FindWindow("_Oscar_BuddyListWin", vbNullString)
TheLength = SendMessageLong(OscarBuddyListWin, &HE, 0&, 0&)
TheText = String(TheLength, 0)
SendMessageByString OscarBuddyListWin, &HD, TheLength, TheText
SN = Left(TheText, InStr(TheText, "'") + 1)
TrimSpaces (SN)
AIM_GetUser = SN

Dim TheLength As Long, TheText As String
Dim OscarBuddyListWin As Long, SN As String
OscarBuddyListWin = FindWindow("_Oscar_BuddyListWin", vbNullString)
TheLength = SendMessageLong(OscarBuddyListWin, &HE, 0&, 0&)
TheText = String(TheLength, 0)
SendMessageByString OscarBuddyListWin, &HD, TheLength, TheText
SN = Left(TheText, InStr(TheText, "'") + 1)
SN = TrimSpaces(LCase(SN))
AIM_GetUser = SN

Call FlashWindow(Window, 0&)

Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)

Call SetWindowPos(Window, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)

Current = Timer
Do While Timer - Current < Val(Time)
    DoEvents
Loop

Dim AOL As Long, MDI As Long, welcome As Long
Dim child As Long, UserString As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
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

Num1& = GetWindow(parentw, 5)
If UCase(Mid(GetClass(Num1&), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Num1& = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(Num1&), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
While Num1&
    Num2& = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(Num2&), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
    Num1& = GetWindow(Num1&, 2)
    If UCase(Mid(GetClass(Num1&), 1, Len(childhand))) Like UCase(childhand) Then GoTo god
Wend
FindIt = 0
god:
meeh& = Num1&
FindIt = meeh&

Room& = FindRoom()
If Room& Then
    hChatEdit& = Findit2(Room&, "RICHCNTL")
    ret = SenditByString(hChatEdit&, WM_SETTEXT, 0, txt)
    ret = SenditbyNum(hChatEdit&, WM_CHAR, 13, 0)
    Wait 0.079
End If

Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLStatic As Long, AOLEdit As Long, AOLIcon As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
If AOLStatic <> 0& And AOLEdit <> 0& And AOLIcon <> 0& Then
    FindFavEdit = AOLChild
    Exit Function
Else
    While AOLChild
        AOLChild = GetWindow(AOLChild, 2)
        AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
        AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
        AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
        If AOLStatic <> 0& And AOLEdit <> 0& And AOLIcon <> 0& Then
            FindFavEdit = AOLChild
            Exit Function
        End If
    Wend
End If

firs& = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
firs& = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
While firs&
    firs& = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
    firs& = GetWindow(firs&, 2)
    If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo found
Wend
Findit2 = 0
found:
firs& = GetWindow(firs&, 2)
If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
firs& = GetWindow(firs&, 2)
If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
While firs&
    firs& = GetWindow(firs&, 2)
    If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs& = GetWindow(firs&, 2)
    If UCase(Mid(GetClass(firs&), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
Wend
Findit2 = 0
Found2:
Findit2 = firs&

SendNow& = SenditbyNum(Button&, WM_LBUTTONDOWN, &HD, 0)
SendNow& = SenditbyNum(Button&, WM_LBUTTONUP, &HD, 0)

If List.ListCount < 0 Then
    Exit Sub
Else
    For X = 0 To List.ListCount - 1
        If List1.Text Like "*" + Item + "*" Then
            List.RemoveItem List.ListIndex
        End If
    Next X
End If

