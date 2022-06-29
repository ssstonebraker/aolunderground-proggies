Attribute VB_Name = "JMR32_AOL5"
'-     This is my AOL5 module.
'-     If you use my module then please add me to your
'-     greets.
'-
'-     Add to this module, manipulate my code,
'-     and most of all, be original.  I want to see some new
'-     shit out there, thats why I encourage add-ons to my
'-     module.
'-
'-     On a final note: do enjoy =)
'-     -JMR
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
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
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
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
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Type POINTAPI
X As Long
Y As Long
End Type
Function AOL()
    AOL = FindWindow("AOL Frame25", vbNullString)
End Function
Function FindChat()
    MDIC& = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
    Childo& = FindWindowEx(MDIC&, 0&, "AOL Child", vbNullString)
    Lister& = FindWindowEx(Childo&, 0&, "_AOL_Listbox", vbNullString)
    Richy1& = FindWindowEx(Childo&, 0&, "RICHCNTL", vbNullString)
    Richy2& = FindWindowEx(Childo&, Richy1&, "RICHCNTL", vbNullString)
    If MDIC& <> 0 And Childo& <> 0 And Lister& <> 0 And Richy1& <> 0 And Richy2& <> 0 Then
        GoTo oggy
    Else
        Do
            Childo& = FindWindowEx(MDIC&, Childo&, "AOL Child", vbNullString)
            Lister& = FindWindowEx(Childo&, 0&, "_AOL_Listbox", vbNullString)
            Richy1& = FindWindowEx(Childo&, 0&, "RICHCNTL", vbNullString)
            Richy2& = FindWindowEx(Childo&, Richy1&, "RICHCNTL", vbNullString)
            If Childo& <> 0 And Lister& <> 0 And Richy1& <> 0 And Richy2& <> 0 Then
            GoTo oggy
            End If
        Loop Until Childo& = 0
    End If
oggy:
    FindChat = Childo&
End Function
Sub ChatSend(YourText As String)
    If FindChat = 0 Then
        Exit Sub
    Else
        Do
        ab1& = FindWindowEx(FindChat, 0&, "RICHCNTL", vbNullString)
        ab2& = FindWindowEx(FindChat, ab1&, "RICHCNTL", vbNullString)
        If ab2& <> 0 Then Exit Do
        Loop
        BoxLen& = SendMessage(ab2&, WM_GETTEXTLENGTH, 0&, 0&)
        EText$ = String(BoxLen&, 0&)
        Call SendMessageByString(ab2&, WM_GETTEXT, BoxLen& + 1, EText$)
        Call SendMessageByString(ab2&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(ab2&, WM_SETTEXT, 0&, YourText)
        Call SendMessage(ab2&, WM_CHAR, ENTER_KEY, 0&)
        Call SendMessageByString(ab2&, WM_SETTEXT, 0&, EText$)
    End If
End Sub
Sub Pause(HowLong As Integer)
    Counter = Timer
    Do Until Timer - Counter >= HowLong
        DoEvents
    Loop
End Sub
Sub FlashMail()
Dim Mousey As POINTAPI
    TBar1& = FindWindowEx(AOL, 0&, "AOL Toolbar", vbNullString)
    TBar2& = FindWindowEx(TBar1&, 0&, "_AOL_Toolbar", vbNullString)
    Ico1& = FindWindowEx(TBar2&, 0&, "_AOL_Icon", vbNullString)
    Ico2& = FindWindowEx(TBar2&, Ico1&, "_AOL_Icon", vbNullString)
    Ico3& = FindWindowEx(TBar2&, Ico2&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(Mousey)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(Ico3&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(Ico3&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        PopWin& = FindWindow("#32768", vbNullString)
        If IsWindowVisible(PopWin&) <> 0 Then Exit Do
    Loop
    For GoDown = 1 To 14
        Call PostMessage(PopWin&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(PopWin&, WM_KEYUP, VK_DOWN, 0&)
    Next GoDown
    Call PostMessage(PopWin&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call PostMessage(PopWin&, WM_KEYUP, VK_RIGHT, 0&)
    Call PostMessage(PopWin&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(PopWin&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(Mousey.X, Mousey.Y)
End Sub
Sub AutoAOL()
Dim Mousey As POINTAPI
    TBar1& = FindWindowEx(AOL, 0&, "AOL Toolbar", vbNullString)
    TBar2& = FindWindowEx(TBar1&, 0&, "_AOL_Toolbar", vbNullString)
    Ico1& = FindWindowEx(TBar2&, 0&, "_AOL_Icon", vbNullString)
    Ico2& = FindWindowEx(TBar2&, Ico1&, "_AOL_Icon", vbNullString)
    Ico3& = FindWindowEx(TBar2&, Ico2&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(Mousey)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(Ico3&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(Ico3&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        PopWin& = FindWindow("#32768", vbNullString)
        If IsWindowVisible(PopWin&) <> 0 Then Exit Do
    Loop
    For GoDown = 1 To 12
        Call PostMessage(PopWin&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(PopWin&, WM_KEYUP, VK_DOWN, 0&)
    Next GoDown
    Call PostMessage(PopWin&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(PopWin&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(Mousey.X, Mousey.Y)
    Do
        MDIC& = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
        AOC& = FindWindowEx(MDIC&, 0&, "AOL Child", "Automatic AOL")
        Check1& = FindWindowEx(AOC&, 0&, "_AOL_Checkbox", vbNullString)
        Check2& = FindWindowEx(AOC&, Check1&, "_AOL_Checkbox", vbNullString)
        Check3& = FindWindowEx(AOC&, Check2&, "_AOL_Checkbox", vbNullString)
        Check4& = FindWindowEx(AOC&, Check3&, "_AOL_Checkbox", vbNullString)
        Check5& = FindWindowEx(AOC&, Check4&, "_AOL_Checkbox", vbNullString)
        Check6& = FindWindowEx(AOC&, Check5&, "_AOL_Checkbox", vbNullString)
        If MDIC& <> 0 And AOC <> 0 And Check6& <> 0 Then Exit Do
    Loop
    Pause 0.5 'Have to give it a little time or it will fuck up.
    RunBtn1& = FindWindowEx(AOC&, 0&, "_AOL_Icon", vbNullString)
    RunBtn2& = FindWindowEx(AOC&, RunBtn1&, "_AOL_Icon", vbNullString)
    Call SendMessage(Check1&, BM_SETCHECK, False, 0&)
    Call SendMessage(Check2&, BM_SETCHECK, True, 0&)
    Call SendMessage(Check3&, BM_SETCHECK, False, 0&)
    Call SendMessage(Check4&, BM_SETCHECK, False, 0&)
    Call SendMessage(Check5&, BM_SETCHECK, False, 0&)
    Call SendMessage(Check6&, BM_SETCHECK, False, 0&)
    Call SendMessage(RunBtn2&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(RunBtn2&, WM_LBUTTONUP, 0&, 0&)
    Do
        Modal& = FindWindow("_AOL_Modal", "Run Automatic AOL Now")
        Check7& = FindWindowEx(Modal&, 0&, "_AOL_Checkbox", vbNullString)
        If Modal& <> 0 And Check7 <> 0 Then Exit Do
    Loop
    Call SendMessage(Check7&, BM_SETCHECK, False, 0&)
    BeginBtn& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
    Call SendMessage(BeginBtn&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(BeginBtn&, WM_LBUTTONUP, 0&, 0&)
End Sub
Function GetListText(YourList As Long, TheItem) As String
    On Error Resume Next 'Justin Case (pun)
    Lenny& = SendMessage(YourList, LB_GETTEXTLEN, 0&, 0&)
    ELStringo$ = String$(Lenny& + 12, vbNullChar)
    Call SendMessageByString(YourList, LB_GETTEXT, TheItem, ELStringo$)
    GetListText = ELStringo$
End Function
Sub MakeFlashList(YourTextbox As TextBox)
    Dim BLT1 As ListBox, BLT2 As ListBox
    MDICl& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
    AOLCh& = FindWindowEx(MDICl&, 0, "AOL Child", "Incoming/Saved Mail")
    Listy& = FindWindowEx(AOLCh&, 0, "_AOL_Tree", vbNullString)
    If MDICl& <> 0 And AOLCh& <> 0 And Listy& <> 0 Then GoTo obesity
    FlashMail
    Do
        MDICl& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
        AOLCh& = FindWindowEx(MDICl&, 0, "AOL Child", "Incoming/Saved Mail")
        Listy& = FindWindowEx(AOLCh&, 0, "_AOL_Tree", vbNullString)
        If MDICl& <> 0 And AOLCh& <> 0 And Listy& <> 0 Then Exit Do
    Loop
obesity:
    ListAm& = SendMessage(Listy&, LB_GETCOUNT, 0&, 0&)
    For abcd = 0 To (ListAm& - 1)
        TheString$ = GetListText(Listy&, abcd)
        TheString = Right(TheString$, Len(TheString$) - Len(Left(TheString$, InStr(TheString$, vbTab))))
        TheString = Right(TheString$, Len(TheString$) - Len(Left(TheString$, InStr(TheString$, vbTab))))
        If abcd = 500 Then YourTextbox = YourTextbox & vbCrLf & "-endofmail1-"
        If abcd = 1000 Then YourTextbox = YourTextbox & vbCrLf & "-endofmail2-"
        If abcd = 1500 Then YourTextbox = YourTextbox & vbCrLf & "-endofmail3-"
        If abcd = 2000 Then YourTextbox = YourTextbox & vbCrLf & "-endofmail4-"
        If abcd = 2500 Then YourTextbox = YourTextbox & vbCrLf & "-endofmail5-"
        YourTextbox = YourTextbox & vbCrLf & "<u>" & (abcd + 1) & ".)</u>  " & TheString$
    Next abcd
End Sub
Sub SendLists(PersonTo, YourSubject, YourMessage, TheMailList)
Dim thelist1, thelist2, thelist3, thelist4, thelist5, thelist6
Dim listnum1, listnum2, listnum3, listnum4, listnum5
listnum1 = InStr(1, TheMaiList, "-endofmail1-")
listnum2 = InStr(1, TheMaiList, "-endofmail2-")
listnum3 = InStr(1, TheMaiList, "-endofmail3-")
listnum4 = InStr(1, TheMaiList, "-endofmail4-")
listnum5 = InStr(1, TheMaiList, "-endofmail5-")
If listnum1 Then

End If
End Sub
Sub ForwardFlash(ListIndex, People, Message)
    MDICl& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
    AOLCh& = FindWindowEx(MDICl&, 0, "AOL Child", "Incoming/Saved Mail")
    Listy& = FindWindowEx(AOLCh&, 0, "_AOL_Tree", vbNullString)
    If MDICl& <> 0 And AOLCh& <> 0 And Listy& <> 0 Then GoTo yekkk
    FlashMail
    Do
        MDICl& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
        AOLCh& = FindWindowEx(MDICl&, 0, "AOL Child", "Incoming/Saved Mail")
        Listy& = FindWindowEx(AOLCh&, 0, "_AOL_Tree", vbNullString)
        If MDICl& <> 0 And AOLCh& <> 0 And Listy& <> 0 Then Exit Do
    Loop
yekkk:
    MDICli1& = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
    AOLChi1& = FindWindowEx(MDICli1&, 0&, "AOL Child", "Incoming/Saved Mail")
    ListA& = FindWindowEx(AOLChi1&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessage(ListA&, LB_SETCURSEL, ListIndex - 1, 0&)
    Call SendMessage(ListA&, WM_KEYDOWN, VK_RETURN, 0&)
    Call SendMessage(ListA&, WM_KEYUP, VK_RETURN, 0&)
    Do
        MDICli2& = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
        AOLChi2& = FindWindowEx(MDICli2&, 0&, "AOL Child", vbNullString)
        Richy& = FindWindowEx(AOLChi2&, 0&, "RICHCNTL", vbNullString)
        If MDICli2& <> 0 And AOLChi2& <> 0 And Richy& <> 0 Then Exit Do
    Loop
    Pause 0.6
    FIco1& = FindWindowEx(AOLChi2&, 0&, "_AOL_Icon", vbNullString)
    FIco2& = FindWindowEx(AOLChi2&, FIco1&, "_AOL_Icon", vbNullString)
    FIco3& = FindWindowEx(AOLChi2&, FIco2&, "_AOL_Icon", vbNullString)
    FIco4& = FindWindowEx(AOLChi2&, FIco3&, "_AOL_Icon", vbNullString)
    FIco5& = FindWindowEx(AOLChi2&, FIco4&, "_AOL_Icon", vbNullString)
    FIco6& = FindWindowEx(AOLChi2&, FIco5&, "_AOL_Icon", vbNullString)
    FIco7& = FindWindowEx(AOLChi2&, FIco6&, "_AOL_Icon", vbNullString)
    FIco8& = FindWindowEx(AOLChi2&, FIco7&, "_AOL_Icon", vbNullString)
    Call SendMessage(FIco8&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(FIco8&, WM_LBUTTONUP, 0&, 0&)
    Do
        MDICli3& = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
        AOLChi3& = FindWindowEx(MDICli3&, 0&, "AOL Child", vbNullString)
        Edit1& = FindWindowEx(AOLChi3&, 0&, "_AOL_Edit", vbNullString)
        Edit2& = FindWindowEx(AOLChi3&, Edit1&, "_AOL_Edit", vbNullString)
        If MDICli3& <> 0 And AOLChi3& <> 0 And Edit1& <> 0 And Edit2& <> 0 Then Exit Do
    Loop
    Pause 0.7
    SIco1& = FindWindowEx(AOLChi3&, 0, "_AOL_Icon", vbNullString)
    SIco2& = FindWindowEx(AOLChi3&, SIco1&, "_AOL_Icon", vbNullString)
    SIco3& = FindWindowEx(AOLChi3&, SIco2&, "_AOL_Icon", vbNullString)
    SIco4& = FindWindowEx(AOLChi3&, SIco3&, "_AOL_Icon", vbNullString)
    SIco5& = FindWindowEx(AOLChi3&, SIco4&, "_AOL_Icon", vbNullString)
    SIco6& = FindWindowEx(AOLChi3&, SIco5&, "_AOL_Icon", vbNullString)
    SIco7& = FindWindowEx(AOLChi3&, SIco6&, "_AOL_Icon", vbNullString)
    SIco8& = FindWindowEx(AOLChi3&, SIco7&, "_AOL_Icon", vbNullString)
    SIco9& = FindWindowEx(AOLChi3&, SIco8&, "_AOL_Icon", vbNullString)
    SIco10& = FindWindowEx(AOLChi3&, SIco9&, "_AOL_Icon", vbNullString)
    SIco11& = FindWindowEx(AOLChi3&, SIco10&, "_AOL_Icon", vbNullString)
    SIco12& = FindWindowEx(AOLChi3&, SIco11&, "_AOL_Icon", vbNullString)
    SIco13& = FindWindowEx(AOLChi3&, SIco12&, "_AOL_Icon", vbNullString)
    SIco14& = FindWindowEx(AOLChi3&, SIco13&, "_AOL_Icon", vbNullString)
    Richa& = FindWindowEx(AOLChi3&, 0&, "RICHCNTL", vbNullString)
    Call SendMessageByString(Edit1&, WM_SETTEXT, 0&, People)
    Call SendMessageByString(Richa&, WM_SETTEXT, 0&, Message)
    Call SendMessage(SIco14&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SIco14&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(AOLChi2&, WM_CLOSE, 0&, 0&)
    Pause 0.6
End Sub
Sub SendMail(ThePeople, TheSubject, TheMessage)
    toolbar1& = FindWindowEx(AOL, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    TIco1& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    TIco2& = FindWindowEx(toolbar2&, TIco1&, "_AOL_Icon", vbNullString)
    Call PostMessage(TIco2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(TIco2&, WM_LBUTTONUP, 0&, 0&)
    Do
        MDIu& = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
        AOLCu& = FindWindowEx(MDIu, 0&, "AOL Child", "Write Mail")
        MIco1& = FindWindowEx(AOLCu&, 0&, "_AOL_Icon", vbNullString)
        MIco2& = FindWindowEx(AOLCu&, MIco1&, "_AOL_Icon", vbNullString)
        MIco3& = FindWindowEx(AOLCu&, MIco2&, "_AOL_Icon", vbNullString)
        MIco4& = FindWindowEx(AOLCu&, MIco3&, "_AOL_Icon", vbNullString)
        MIco5& = FindWindowEx(AOLCu&, MIco4&, "_AOL_Icon", vbNullString)
        MIco6& = FindWindowEx(AOLCu&, MIco5&, "_AOL_Icon", vbNullString)
        MIco7& = FindWindowEx(AOLCu&, MIco6&, "_AOL_Icon", vbNullString)
        MIco8& = FindWindowEx(AOLCu&, MIco7&, "_AOL_Icon", vbNullString)
        MIco9& = FindWindowEx(AOLCu&, MIco8&, "_AOL_Icon", vbNullString)
        MIco10& = FindWindowEx(AOLCu&, MIco9&, "_AOL_Icon", vbNullString)
        MIco11& = FindWindowEx(AOLCu&, MIco10&, "_AOL_Icon", vbNullString)
        MIco12& = FindWindowEx(AOLCu&, MIco11&, "_AOL_Icon", vbNullString)
        MIco13& = FindWindowEx(AOLCu&, MIco12&, "_AOL_Icon", vbNullString)
        MIco14& = FindWindowEx(AOLCu&, MIco13&, "_AOL_Icon", vbNullString)
        MIco15& = FindWindowEx(AOLCu&, MIco14&, "_AOL_Icon", vbNullString)
        MIco16& = FindWindowEx(AOLCu&, MIco15&, "_AOL_Icon", vbNullString)
        Edit1& = FindWindowEx(AOLCu&, 0&, "_AOL_Edit", vbNullString)
        Edit2& = FindWindowEx(AOLCu&, Edit1&, "_AOL_Edit", vbNullString)
        Edit3& = FindWindowEx(AOLCu&, Edit2&, "_AOL_Edit", vbNullString)
        Rych& = FindWindowEx(AOLCu&, 0&, "RICHCNTL", vbNullString)
        If MDIu& <> 0 And AOLCu& <> 0 And MIco16& <> 0 And Edit1& <> 0 And Edit2& <> 0 And Edit3& <> 0 And Rych& <> 0 Then Exit Do
        Pause 1
    Loop
    Call SendMessageByString(Edit1&, WM_SETTEXT, 0&, ThePeople)
    Call SendMessageByString(Edit3&, WM_SETTEXT, 0&, TheSubject)
    Call SendMessageByString(Rych&, WM_SETTEXT, 0&, TheMessage)
    Call SendMessage(MIco16&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(MIco16&, WM_LBUTTONUP, 0&, 0&)
    Pause 0.6
End Sub
Function FindFlash(WhatToFind) As String
    FindFlash = ""
    Midi& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
    AOC& = FindWindowEx(Midi&, 0, "AOL Child", "Incoming/Saved Mail")
    Listed& = FindWindowEx(AOC&, 0, "_AOL_Tree", vbNullString)
    Listop& = SendMessage(Listed&, LB_GETCOUNT, 0&, 0&)
    For SO = 0 To (Listop& - 1)
    Blok$ = GetListText(Listed&, SO)
    OND = InStr(1, LCase(Blok$), LCase(WhatToFind))
        If OND > 0 Then
            FindFlash = FindFlash & (SO + 1) & ".)" & vbTab & Blok$ & vbCrLf
        End If
    Next SO
End Function
Function UserSN() As String
    On Error Resume Next
    AMDI& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
    Elol% = GetWindow(AMDI&, 5)
    While Elol%
        TexLen = GetWindowTextLength(Elol%)
        TheStreeng$ = String$(TexLen, 0)
        TexLen = GetWindowText(Elol%, TheStreeng$, TexLen + 1)
        If InStr(TheStreeng$, "Welcome,") Then GoTo Nexx
        Elol% = GetWindow(Elol%, 2)
    Wend
Nexx:
    ThaLen = GetWindowTextLength(Elol%)
    AStreeng$ = String(ThaLen, 0&)
    Call GetWindowText(Elol%, AStreeng$, ThaLen + 1)
    AStreeng$ = Left(AStreeng$, Len(AStreeng$) - 2)
    AStreeng$ = Right(AStreeng$, Len(AStreeng$) - 9)
    UserSN = AStreeng$
End Function
Sub StayOnTop(YourWindow)
    Call SetWindowPos(YourWindow, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub DontStayOnTop(YourWindow)
    Call SetWindowPos(YourWindow, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Sub MoveWindow(YourWindow)
    ReleaseCapture
    Call SendMessage(YourWindow, WM_SYSCOMMAND, WM_MOVE, 0&)
End Sub
Sub RmvConfirmSend()
Dim Mouzer As POINTAPI
    TheTool1& = FindWindowEx(AOL, 0&, "AOL Toolbar", vbNullString)
    TheTool2& = FindWindowEx(TheTool1&, 0&, "_AOL_Toolbar", vbNullString)
    TIco1& = FindWindowEx(TheTool2&, 0&, "_AOL_Icon", vbNullString)
    TIco2& = FindWindowEx(TheTool2&, TIco1&, "_AOL_Icon", vbNullString)
    TIco3& = FindWindowEx(TheTool2&, TIco2&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(Mouzer)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(TIco3&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(TIco3&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        PopWeezle& = FindWindow("#32768", vbNullString)
        If IsWindowVisible(PopWeezle&) <> 0 Then Exit Do
    Loop
    For GoDown = 1 To 8
        Call PostMessage(PopWeezle&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(PopWeezle&, WM_KEYUP, VK_DOWN, 0&)
    Next GoDown
    Call PostMessage(PopWeezle&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(PopWeezle&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(Mouzer.X, Mouzer.Y)
    Do
        MrsModal& = FindWindow("_AOL_Modal", "Mail Preferences")
        Son1& = FindWindowEx(MrsModal&, 0&, "_AOL_Checkbox", vbNullString)
        Son2& = FindWindowEx(MrsModal&, Son1&, "_AOL_Checkbox", vbNullString)
        Son3& = FindWindowEx(MrsModal&, Son2&, "_AOL_Checkbox", vbNullString)
        Butin& = FindWindowEx(MrsModal&, 0&, "_AOL_Icon", vbNullString)
        Call SendMessage(Son3&, BM_SETCHECK, False, 0&) 'popped it in the loop so it effectivly works
        If MrsModal& <> 0 And Son3& <> 0 And Butin& <> 0 Then Exit Do
    Loop
    Call SendMessage(Butin&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Butin&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub SendIM(ThePerson, TheMessage)
Dim Mousey As POINTAPI
    TBar1& = FindWindowEx(AOL, 0&, "AOL Toolbar", vbNullString)
    TBar2& = FindWindowEx(TBar1&, 0&, "_AOL_Toolbar", vbNullString)
    Ico1& = FindWindowEx(TBar2&, 0&, "_AOL_Icon", vbNullString)
    For af = 1 To 9
        Ico1& = FindWindowEx(TBar2&, Ico1&, "_AOL_Icon", vbNullString)
    Next af
    Call GetCursorPos(Mousey)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(Ico1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(Ico1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        PopWin& = FindWindow("#32768", vbNullString)
        If IsWindowVisible(PopWin&) <> 0 Then Exit Do
    Loop
    For GoDown = 1 To 6
        Call PostMessage(PopWin&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(PopWin&, WM_KEYUP, VK_DOWN, 0&)
    Next GoDown
    Call PostMessage(PopWin&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(PopWin&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(Mousey.X, Mousey.Y)
    Do
        Mido& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
        EOLC& = FindWindowEx(Mido&, 0, "AOL Child", "Send Instant Message")
        Edert& = FindWindowEx(EOLC&, 0, "_AOL_Edit", vbNullString)
        Rychy& = FindWindowEx(EOLC&, 0, "RICHCNTL", vbNullString)
        Beten& = FindWindowEx(EOLC&, 0, "_AOL_Icon", vbNullString)
        For agg = 1 To 8
            Beten& = FindWindowEx(EOLC&, Beten&, "_AOL_Icon", vbNullString)
        Next agg
        If Mido& <> 0 And EOLC& <> 0 And Edert& <> 0 And Rychy& And Beten& <> 0 Then Exit Do
    Loop
    Call SendMessageByString(Edert&, WM_SETTEXT, 0&, ThePerson)
    Call SendMessageByString(Rychy&, WM_SETTEXT, 0&, TheMessage)
    Call PostMessage(Beten&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(Beten&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub IMsOff()
    SendIM "$IM_Off", " "
    Do
        AOL2& = FindWindow("#32770", "America Online")
        Bottun& = FindWindowEx(AOL2&, 0, "Button", vbNullString)
        If AOL2& <> 0 And Bottun& <> 0 Then Exit Do
    Loop
    Call SendMessage(Bottun&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Bottun&, WM_KEYUP, VK_SPACE, 0&)
    Do
        Mido& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
        EOLC& = FindWindowEx(Mido&, 0, "AOL Child", "Send Instant Message")
        Edert& = FindWindowEx(EOLC&, 0, "_AOL_Edit", vbNullString)
        Rychy& = FindWindowEx(EOLC&, 0, "RICHCNTL", vbNullString)
        Beten& = FindWindowEx(EOLC&, 0, "_AOL_Icon", vbNullString)
        For agg = 1 To 8
            Beten& = FindWindowEx(EOLC&, Beten&, "_AOL_Icon", vbNullString)
        Next agg
        If Mido& <> 0 And EOLC& <> 0 And Edert& <> 0 And Rychy& And Beten& <> 0 Then Exit Do
    Loop
    Call SendMessage(EOLC&, WM_CLOSE, 0&, 0&)
End Sub
Sub IMsOn()
    SendIM "$IM_On", " "
    Do
        AOL2& = FindWindow("#32770", "America Online")
        Bottun& = FindWindowEx(AOL2&, 0, "Button", vbNullString)
        If AOL2& <> 0 And Bottun& <> 0 Then Exit Do
    Loop
    Call SendMessage(Bottun&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Bottun&, WM_KEYUP, VK_SPACE, 0&)
    Do
        Mido& = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
        EOLC& = FindWindowEx(Mido&, 0, "AOL Child", "Send Instant Message")
        Edert& = FindWindowEx(EOLC&, 0, "_AOL_Edit", vbNullString)
        Rychy& = FindWindowEx(EOLC&, 0, "RICHCNTL", vbNullString)
        Beten& = FindWindowEx(EOLC&, 0, "_AOL_Icon", vbNullString)
        For agg = 1 To 8
            Beten& = FindWindowEx(EOLC&, Beten&, "_AOL_Icon", vbNullString)
        Next agg
        If Mido& <> 0 And EOLC& <> 0 And Edert& <> 0 And Rychy& And Beten& <> 0 Then Exit Do
    Loop
    Call SendMessage(EOLC&, WM_CLOSE, 0&, 0&)
End Sub
Sub Keyword(YourKeyword)
    Toobar1& = FindWindowEx(AOL, 0, "AOL Toolbar", vbNullString)
    Toobar2& = FindWindowEx(Toobar1&, 0, "_AOL_Toolbar", vbNullString)
    AOLChildr& = FindWindowEx(Toobar2&, 0, "_AOL_Combobox", vbNullString)
    TheEdit& = FindWindowEx(AOLChildr&, 0, "Edit", vbNullString)
    Call SendMessageByString(TheEdit&, WM_SETTEXT, 0&, YourKeyword)
    Call SendMessage(TheEdit&, WM_CHAR, ENTER_KEY, 0&)
End Sub
