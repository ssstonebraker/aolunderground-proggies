Attribute VB_Name = "dos32"


Option Explicit
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
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
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Const GW_HWNDFIRST = 0
Const GW_HWNDNEXT = 2
Const GW_CHILD = 5


Const WM_GETTEXT = &HD
Const WM_GETTEXTLENGTH = &HE
Const WM_SETTEXT = &HC
Const WM_CHAR = &H102

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

Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111

Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012

Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        Y As Long
End Type

Public Function FindForwardWindow() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim Rich1 As Long, Rich2 As Long, Combo As Long
    Dim FontCombo As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
    FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
    If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
        FindForwardWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
            Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
            FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
            If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
                FindForwardWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindForwardWindow& = 0&
End Function

Public Function FindSendWindow() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim SendStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
    If SendStatic& <> 0& Then
        FindSendWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
            If SendStatic& <> 0& Then
                FindSendWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindSendWindow& = 0&
End Function

Public Sub MailOpenFlash()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim CurPos As POINTAPI, winvis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        smod& = FindWindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    Call PostMessage(smod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(smod&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RIGHT, 0&)
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
End Sub

Public Sub MailOpenNew()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim toolicon As Long, smod As Long, CurPos As POINTAPI
    Dim winvis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        smod& = FindWindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
End Sub

Public Sub MailOpenOld()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim CurPos As POINTAPI, winvis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        smod& = FindWindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    For dothis& = 1 To 4
        Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Next dothis&
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
End Sub

Public Sub MailOpenSent()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim CurPos As POINTAPI, winvis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        smod& = FindWindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    For dothis& = 1 To 5
        Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Next dothis&
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
End Sub

Public Sub MailOpenEmailFlash(index As Long)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Sub
    Call SendMessage(fList&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(fList&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(fList&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenEmailNew(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenEmailOld(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenEmailSent(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Function MailCountFlash() As Long
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MailCountFlash& = Count&
End Function

Public Sub MailToListFlash(thelist As ListBox)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim Count As Long, MyString As String, AddMails As Long
    Dim sLength As Long, Spot As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    If fMail& = 0& Then Exit Sub
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MyString$ = String(255, 0)
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
    Next AddMails&
End Sub

Public Function FindMailBox() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim TabControl As Long, TabPage As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(Child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            TabControl& = FindWindowEx(Child&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            If TabControl& <> 0& And TabPage& <> 0& Then
                FindMailBox& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindMailBox& = 0&
End Function

Public Function MailCountNew() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountNew& = Count&
End Function

Public Function MailCountSent() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountSent& = Count&
End Function

Public Function MailCountOld() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountOld& = Count&
End Function

Public Sub MailDeleteNewByIndex(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If index& > Count& - 1 Or index& < 0& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailDeleteNewDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String, cSubject As String
    Dim SearchFor As Long, sSender As String, sSubject As String
    Dim CurCaption As String
    MailBox& = FindMailBox&
    CurCaption$ = VBForm.Caption
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0& Then Exit Sub
    For SearchFor& = 0& To Count& - 2
        DoEvents
        sSender$ = MailSenderNew(SearchFor&)
        sSubject$ = MailSubjectNew(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To Count& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Now checking #" & SearchFor& & " for match with #" & SearchBox&
            End If
            cSender$ = MailSenderNew(SearchBox&)
            cSubject$ = MailSubjectNew(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub

Public Sub MailDeleteNewBySender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0& Then Exit Sub
    For SearchBox& = 0& To Count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If LCase(cSender$) = LCase(Sender$) Then
            Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
            DoEvents
            Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub

Public Sub MailDeleteNewNotSender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = FindWindowEx(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = FindWindowEx(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0& Then Exit Sub
    For SearchBox& = 0& To Count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If cSender$ = "" Then Exit Sub
        If LCase(cSender$) <> LCase(Sender$) Then
            Call SendMessage(mTree&, LB_SETCURSEL, SearchBox&, 0&)
            DoEvents
            Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub

Public Function MailSenderFlash(index As Long) As String
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, spot1 As Long, spot2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Function
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fCount& = 0 Or index& > fCount& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(fList&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(fList&, LB_GETTEXT, index&, MyString$)
    spot1& = InStr(MyString$, Chr(9))
    spot2& = InStr(spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, spot1& + 1, spot2& - spot1& - 1)
    MailSenderFlash$ = MyString$
End Function

Public Function MailSenderNew(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim spot1 As Long, spot2 As Long, MyString As String
    Dim Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Or index& > Count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    spot1& = InStr(MyString$, Chr(9))
    spot2& = InStr(spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, spot1& + 1, spot2& - spot1& - 1)
    MailSenderNew$ = MyString$
End Function

Public Function MailSubjectFlash(index As Long) As String
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, Spot As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Function
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fCount& = 0 Or index& > fCount& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(fList&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(fList&, LB_GETTEXT, index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectFlash$ = MyString$
End Function

Public Function MailSubjectNew(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Or index& > Count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectNew$ = MyString$
End Function

Public Sub MailToListNew(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub

Public Sub MailToListOld(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub

Public Sub MailToListSent(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub

Public Sub SendMail(Person As String, subject As String, message As String)
    Dim AOL As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim toolicon As Long, OpenSend As Long, DoIt As Long
    Dim rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
    DoEvents
    Call SendMessageByString(rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailForward(SendTo As String, message As String, DeleteFwd As Boolean)
    Dim AOL As Long, MDI As Long, Error As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long
    Dim TempSubject As String
    OpenForward& = FindForwardWindow
    If OpenForward& = 0 Then Exit Sub
    SendButton& = FindWindowEx(OpenForward&, 0&, "_AOL_Icon", vbNullString)
    For DoIt& = 1 To 6
        SendButton& = FindWindowEx(OpenForward&, SendButton&, "_AOL_Icon", vbNullString)
    Next DoIt&
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OpenSend& = FindSendWindow
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    If DeleteFwd = True Then
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
        DoEvents
    End If
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SendTo$)
    DoEvents
    Call SendMessageByString(rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Do Until OpenSend& = 0& Or Error& <> 0&
        DoEvents
        AOL& = FindWindow("AOL Frame25", vbNullString)
        MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
        Error& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        OpenSend& = FindSendWindow
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 1
    Loop
    If OpenSend& = 0& Then Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub CloseOpenMails()
    Dim OpenSend As Long, OpenForward As Long
    Do
        DoEvents
        OpenSend& = FindSendWindow
        OpenForward& = FindForwardWindow
        Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
        DoEvents
        Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
        DoEvents
    Loop Until OpenSend& = 0& And OpenForward& = 0&
End Sub

Public Sub MailDeleteFlashByIndex(index As Long)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < index& Then Exit Sub
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    Call SendMessage(fList&, LB_SETCURSEL, index&, 0&)
    Call SendMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailDeleteFlashDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, SearchFor As Long
    Dim SearchBox As Long, CurCaption As String
    Dim sSender As String, sSubject As String
    Dim cSender As String, cSubject As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < 2& Then Exit Sub
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    CurCaption$ = VBForm.Caption
    If fCount& = 0& Then Exit Sub
    For SearchFor& = 0& To fCount& - 2
        DoEvents
        sSender$ = MailSenderFlash(SearchFor&)
        sSubject$ = MailSubjectFlash(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To fCount& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Checking #" & SearchFor& & " with #" & SearchBox&
            End If
            cSender$ = MailSenderFlash(SearchBox&)
            cSubject$ = MailSubjectFlash(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call SendMessage(fList&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call SendMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub

Public Sub SetMailPrefs()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim MDI As Long, mprefs As Long, mbutton As Long
    Dim gstatic As Long, mstatic As Long, fstatic As Long
    Dim mastatic As Long, dmod As Long, confirmcheck As Long
    Dim closecheck As Long, spellcheck As Long, OkButton As Long
    Dim CurPos As POINTAPI, winvis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        smod& = FindWindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    For dothis& = 1 To 3
        Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Next dothis&
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
    Do
        DoEvents
        mprefs& = FindWindowEx(MDI&, 0&, "AOL Child", "Preferences")
        gstatic& = FindWindowEx(mprefs&, 0&, "_AOL_Static", "General")
        mstatic& = FindWindowEx(mprefs&, 0&, "_AOL_Static", "Mail")
        fstatic& = FindWindowEx(mprefs&, 0&, "_AOL_Static", "Font")
        mastatic& = FindWindowEx(mprefs&, 0&, "_AOL_Static", "Marketing")
    Loop Until mprefs& <> 0& And gstatic& <> 0& And mstatic& <> 0& And fstatic& <> 0& And mastatic& <> 0&
    mbutton& = FindWindowEx(mprefs&, 0&, "_AOL_Icon", vbNullString)
    mbutton& = FindWindowEx(mprefs&, mbutton&, "_AOL_Icon", vbNullString)
    mbutton& = FindWindowEx(mprefs&, mbutton&, "_AOL_Icon", vbNullString)
    Do
        DoEvents
        Call SendMessage(mbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(mbutton&, WM_LBUTTONUP, 0&, 0&)
        dmod& = FindWindow("_AOL_Modal", "Mail Preferences")
        Pause 0.6
    Loop Until dmod& <> 0&
    confirmcheck& = FindWindowEx(dmod&, 0&, "_AOL_Checkbox", vbNullString)
    closecheck& = FindWindowEx(dmod&, confirmcheck&, "_AOL_Checkbox", vbNullString)
    spellcheck& = FindWindowEx(dmod&, closecheck&, "_AOL_Checkbox", vbNullString)
    spellcheck& = FindWindowEx(dmod&, spellcheck&, "_AOL_Checkbox", vbNullString)
    spellcheck& = FindWindowEx(dmod&, spellcheck&, "_AOL_Checkbox", vbNullString)
    spellcheck& = FindWindowEx(dmod&, spellcheck&, "_AOL_Checkbox", vbNullString)
    OkButton& = FindWindowEx(dmod&, 0&, "_AOL_icon", vbNullString)
    Call SendMessage(confirmcheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(closecheck&, BM_SETCHECK, True, vbNullString)
    Call SendMessage(spellcheck&, BM_SETCHECK, False, vbNullString)
    Call SendMessage(OkButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OkButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Call PostMessage(mprefs&, WM_CLOSE, 0&, 0&)
End Sub

Public Function ErrorName(name As Long) As String
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long, TempString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = GetText(ErrorTextWindow&)
    NameCount& = LineCount(ErrorString$) - 2
    If NameCount& < name& Then Exit Function
    TempString$ = LineFromString(ErrorString$, name& + 2)
    TempString$ = Left(TempString$, InStr(TempString$, "-") - 2)
    ErrorName$ = TempString$
End Function

Public Function ErrorNameCount() As Long
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim NameCount As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
    If ErrorWindow& = 0& Then Exit Function
    ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorString$ = GetText(ErrorTextWindow&)
    NameCount& = LineCount(ErrorString$) - 2
    ErrorNameCount& = NameCount&
End Function

Public Function CheckAlive(ScreenName As String) As Boolean
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, nowindow As Long, nobutton As Long
    Call SendMail("*, " & ScreenName$, "You alive?", "=)")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = GetText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
        CheckAlive = False
    Else
        CheckAlive = True
    End If
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do
        DoEvents
        nowindow& = FindWindow("#32770", "America Online")
        nobutton& = FindWindowEx(nowindow&, 0&, "Button", "&No")
    Loop Until nowindow& <> 0& And nobutton& <> 0
    Call SendMessage(nobutton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(nobutton&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Sub ChatSendg(Chat As String)
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Function FindIM() As Long
    Dim AOL As Long, MDI As Long, Child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = getcaption(Child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        FindIM& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Caption$ = getcaption(Child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                FindIM& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindIM& = Child&
End Function

Public Function FindRoom() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim rich As Long, aollist As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    aollist& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    If rich& <> 0& And aollist& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            aollist& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            If rich& <> 0& And aollist& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindRoom& = Child&
End Function

Public Function FindInfoWindow() As Long
    Dim AOL As Long, MDI As Long, Child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
            AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindInfoWindow& = Child&
End Function

Public Function RoomCount() As Long
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = Count&
End Function

Public Sub AddRoomToListbox(thelist As ListBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                thelist.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub AddRoomToCombobox(TheCombo As ComboBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                TheCombo.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
    If TheCombo.ListCount > 0 Then
        TheCombo.text = TheCombo.List(0)
    End If
End Sub

Public Sub ChatIgnoreByIndex(index As Long)
    Dim Room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, a As Long, Count As Long
    Count& = RoomCount&
    If index& > Count& - 1 Then Exit Sub
    Room& = FindRoom&
    sList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call SendMessage(sList&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(sList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = FindWindowEx(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
    Do
        DoEvents
        a& = SendMessage(iCheck&, BM_GETCHECK, 0&, 0&)
        Call PostMessage(iCheck&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call PostMessage(iCheck&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
    Loop Until a& <> 0&
    DoEvents
    Call PostMessage(iWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub ChatIgnoreByName(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Function ChatLineSN(TheChatLine As String) As String
    If InStr(TheChatLine, ":") = 0 Then
        ChatLineSN = ""
        Exit Function
    End If
    ChatLineSN = Left(TheChatLine, InStr(TheChatLine, ":") - 1)
End Function

Public Function ChatLineMsg(TheChatLine As String) As String
    If InStr(TheChatLine, Chr(9)) = 0 Then
        ChatLineMsg = ""
        Exit Function
    End If
    ChatLineMsg = Right(TheChatLine, Len(TheChatLine) - InStr(TheChatLine, Chr(9)))
End Function

Public Sub Scroll(ScrollString As String)
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

Public Sub WaitForOKOrRoom(Room As String)
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    Room$ = LCase(ReplaceString(Room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = getcaption(FindRoom&)
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
    DoEvents
End Sub
Public Sub MemberRoom(Room As String)
    Call Keyword("aol://2719:61-2-" & Room$)
End Sub

Public Sub PublicRoom(Room As String)
    Call Keyword("aol://2719:21-2-" & Room$)
End Sub

Public Sub PrivateRoom(Room As String)
    Call Keyword("aol://2719:2-2-" & Room$)
End Sub

Public Sub InstantMessage(Person As String, message As String)
    Dim AOL As Long, MDI As Long, IM As Long, rich As Long
    Dim SendButton As Long, ok As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until IM& <> 0& And rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(rich&, WM_SETTEXT, 0&, message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        ok& = FindWindow("#32770", "America Online")
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until ok& <> 0& Or IM& = 0&
    If ok& <> 0& Then
        Button& = FindWindowEx(ok&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Function CheckIMs(Person As String) As Boolean
    Dim AOL As Long, MDI As Long, IM As Long, rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(IM&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(IM&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(IM&, Available&, "_AOL_Icon", vbNullString)
    Loop Until IM& <> 0& And rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
    ChatSend "<font color=" & """" & "#000000" & """" & "><font face=" & """" & "arial narrow" & """" & ">" & Person$ & "'z |mz ÅrÈ øÑ"
    Else
        CheckIMs = False
    ChatSend "<font color=" & """" & "#000000" & """" & "><font face=" & """" & "arial narrow" & """" & ">" & Person$ & "'z |mz ÅrÈ øFf"
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End Function

Public Sub IMIgnore(Person As String)
    Call InstantMessage("$IM_OFF, " & Person$, "=)")
End Sub

Public Sub IMUnIgnore(Person As String)
    Call InstantMessage("$IM_ON, " & Person$, "=)")
End Sub

Public Sub IMsOff()
    Call InstantMessage("$IM_OFF", "=)")
End Sub

Public Sub IMsOn()
    Call InstantMessage("$IM_ON", "=)")
End Sub

Public Function IMSender() As String
    Dim IM As Long, Caption As String
    Caption$ = getcaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function

Public Function IMText() As String
    Dim rich As Long
    rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    IMText$ = GetText(rich&)
End Function

Public Function IMLastMsg() As String
    Dim rich As Long, MsgString As String, Spot As Long
    Dim NewSpot As Long
    rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    MsgString$ = GetText(rich&)
    NewSpot& = InStr(MsgString$, Chr(9))
    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
    MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
    IMLastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function

Public Sub IMRespond(Msg As String)
    Dim IM As Long, rich As Long, Icon As Long
    IM& = FindIM&
    If IM& = 0& Then Exit Sub
    rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    rich& = FindWindowEx(IM&, rich&, "RICHCNTL", vbNullString)
    Icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(rich&, WM_SETTEXT, 0&, Msg$)
    DoEvents
    Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub


Public Function DoubleText(MyString As String) As String
    Dim NewString As String, CurChar As String
    Dim DoIt As Long
    If MyString$ <> "" Then
        For DoIt& = 1 To Len(MyString$)
            CurChar$ = LineChar(MyString$, DoIt&)
            NewString$ = NewString$ & CurChar$ & CurChar$
        Next DoIt&
        DoubleText$ = NewString$
    End If
End Function

Public Function LineChar(TheText As String, CharNum As Long) As String
    Dim TextLength As Long, NewText As String
    TextLength& = Len(TheText$)
    If CharNum& > TextLength& Then
        Exit Function
    End If
    NewText$ = Left(TheText$, CharNum&)
    NewText$ = Right(NewText$, 1)
    LineChar$ = NewText$
End Function

Public Function LineCount(MyString As String) As Long
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

Public Function LineFromString(MyString As String, Line As Long) As String
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

Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
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

Public Function ReverseString(MyString As String) As String
    Dim TempString As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While Count& <= StringLength&
        Count& = Count& + 1
        NextChr$ = Mid$(TempString$, Count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
End Function

Public Function SwitchStrings(MyString As String, String1 As String, String2 As String) As String
    Dim TempString As String, spot1 As Long, spot2 As Long
    Dim Spot As Long, ToFind As String, ReplaceWith As String
    Dim NewSpot As Long, LeftString As String, RightString As String
    Dim NewString As String
    If Len(String2) > Len(String1) Then
        TempString$ = String1$
        String1$ = String2$
        String2$ = TempString$
    End If
    spot1& = InStr(MyString$, String1$)
    spot2& = InStr(MyString$, String2$)
    If spot1& = 0& And spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    If spot1& < spot2& Or spot2& = 0 Or Len(String1$) = Len(String2$) Then
        If spot1& > 0 Then
            Spot& = spot1&
            ToFind$ = String1$
            ReplaceWith$ = String2$
        End If
    End If
    If spot2& < spot1& Or spot1& = 0& Then
        If spot2& > 0& Then
            Spot& = spot2&
            ToFind$ = String2$
            ReplaceWith$ = String1$
        End If
    End If
    If spot1& = 0& And spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString$ = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot + Len(ReplaceWith$) - Len(ToFind$) + 1
        If Spot& <> 0& Then
            spot1& = InStr(Spot&, MyString$, String1$)
            spot2& = InStr(Spot&, MyString$, String2$)
        End If
        If spot1& = 0& And spot2& = 0& Then
            SwitchStrings$ = MyString$
            Exit Function
        End If
        If spot1& < spot2& Or spot2& = 0& Or Len(String1$) = Len(String2$) Then
            If spot1& > 0& Then
                Spot& = spot1&
                ToFind$ = String1$
                ReplaceWith$ = String2$
            End If
        End If
        If spot2& < spot1& Or spot1& = 0& Then
            If spot2& > 0& Then
                Spot& = spot2&
                ToFind$ = String2$
                ReplaceWith$ = String1$
            End If
        End If
        If spot1& = 0& And spot2& = 0& Then
            Spot& = 0&
        End If
        If Spot& > 0& Then
            NewSpot& = InStr(Spot&, MyString$, ToFind$)
        Else
            NewSpot& = Spot&
        End If
    Loop Until NewSpot& < 1&
    SwitchStrings$ = NewString$
End Function

Public Function MacroFilter_BCurve(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "\", "]")
    MyString$ = ReplaceString(MyString$, "/", "[")
    MacroFilter_BCurve$ = MyString$
End Function

Public Function MacroFilter_BubbleTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "°'°'°'")
    MacroFilter_BubbleTop$ = MyString$
End Function

Public Function MacroFilter_BubbleTop2(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "º°°°'")
    MacroFilter_BubbleTop2$ = MyString$
End Function

Public Function MacroFilter_ClawTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯¯¯¯", "¯\(¯)/" & Chr(34) & "¯")
    MacroFilter_ClawTop$ = MyString$
End Function

Public Function MacroFilter_Curve(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "\", ")")
    MyString$ = ReplaceString(MyString$, "/", "(")
    MacroFilter_Curve$ = MyString$
End Function

Public Function MacroFilter_CurveBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "î,î,î,")
    MacroFilter_CurveBottom$ = MyString$
End Function

Public Function MacroFilter_Darken(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¦", "|")
    MyString$ = ReplaceString(MyString$, ",´", "/ ")
    MyString$ = ReplaceString(MyString$, "`,", " \")
    MyString$ = ReplaceString(MyString$, ":", ";")
    MacroFilter_Darken$ = MyString$
End Function

Public Function MacroFilter_Destroy(MyString As String) As String
    MyString$ = ReplaceString(MyString$, " ", "")
    MacroFilter_Destroy$ = MyString$
End Function

Public Function MacroFilter_DrippingTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯¯¯", "¯\,/¯'v'")
    MacroFilter_DrippingTop$ = MyString$
End Function

Public Function MacroFilter_Electric(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "      |", "--^v^|")
    MyString$ = ReplaceString(MyString$, "|      ", "|^v^--")
    MacroFilter_Electric$ = MyString$
End Function

Public Function MacroFilter_FireyBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "_')\.")
    MacroFilter_FireyBottom$ = MyString$
End Function

Public Function MacroFilter_Ghost(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯", "¨¨")
    MyString$ = ReplaceString(MyString$, "/", ".·")
    MyString$ = ReplaceString(MyString$, "\", "·.")
    MyString$ = ReplaceString(MyString$, "|", ":")
    MyString$ = ReplaceString(MyString$, "_", "..")
    MyString$ = ReplaceString(MyString$, "¦", ":")
    MacroFilter_Ghost = MyString$
End Function

Public Function MacroFilter_Indent(MyString As String) As String
    Dim NewLine As String, OrgLen As Long, NumOfLines As Long
    Dim OrgCount As Long, SpaceIt As Long, CurLine As String
    Dim NewString As String
    NewLine$ = Chr(13) & Chr(10)
    OrgLen& = Len(MyString$)
    MyString$ = MyString$ & NewLine$
    NumOfLines& = LineCount(MyString$)
    OrgCount& = NumOfLines&
    For SpaceIt& = 1 To NumOfLines&
        DoEvents
        CurLine$ = LineFromString(MyString$, SpaceIt&)
        NewString$ = NewString$ & " " & CurLine$ & NewLine$
    Next SpaceIt&
    MyString$ = Left(NewString$, OrgLen& + OrgCount& - 1)
    MacroFilter_Indent$ = MyString$
End Function

Public Function MacroFilter_JaG(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯¯", "¯`v´¯")
    MacroFilter_JaG$ = MyString$
End Function

Public Function MacroFilter_Lighten(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|", "¦")
    MyString$ = ReplaceString(MyString$, "/ ", ",´")
    MyString$ = ReplaceString(MyString$, "\ ", "`,")
    MyString$ = ReplaceString(MyString$, " /", ",´")
    MyString$ = ReplaceString(MyString$, " \", "`,")
    MyString$ = ReplaceString(MyString$, ";", ":")
    MacroFilter_Lighten$ = MyString$
End Function

Public Function MacroFilter_PCurve(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "\", "}")
    MyString$ = ReplaceString(MyString$, "/", "{")
    MacroFilter_PCurve$ = MyString$
End Function

Public Function MacroFilter_PsYTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "`'¯`“")
    MacroFilter_PsYTop$ = MyString$
End Function

Public Function MacroFilter_RandomBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "-‚„‚¸‚")
    MacroFilter_RandomBottom$ = MyString$
End Function

Public Function MacroFilter_Rapid(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "   |", "-=|")
    MyString$ = ReplaceString(MyString$, "|   ", "|=-")
    MacroFilter_Rapid$ = MyString$
End Function

Public Function MacroFilter_ReplaceLines(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|", "¦")
    MacroFilter_ReplaceLines$ = MyString$
End Function

Public Function MacroFilter_ReplaceSlants(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "/ ", ",´")
    MyString$ = ReplaceString(MyString$, "\ ", "`,")
    MyString$ = ReplaceString(MyString$, " /", ",´")
    MyString$ = ReplaceString(MyString$, " \", "`,")
    MacroFilter_ReplaceSlants$ = MyString$
End Function

Public Function MacroFilter_Reverse(MyString As String) As String
    Dim CurChar As Long, NewLine As String, MyText As String
    Dim NumOfLines As Long, ReverseIt As Long, CheckLen As Long
    Dim CurLine As String, NewString As String
    If MyString$ <> "" Then
        NewLine$ = Chr(13) & Chr(10)
        MyText$ = MyString$ & NewLine$
        NumOfLines& = LineCount(MyText$)
        For ReverseIt& = 1 To NumOfLines
            CurLine$ = LineFromString(MyText$, ReverseIt&)
            CurLine$ = ReverseString(CurLine$)
            NewString$ = NewString$ & CurLine$ & NewLine$
        Next ReverseIt&
        NewString$ = SwitchStrings(NewString$, "/", "\")
        NewString$ = SwitchStrings(NewString$, "[", "]")
        NewString$ = SwitchStrings(NewString$, "{", "}")
        NewString$ = SwitchStrings(NewString$, "(", ")")
        NewString$ = SwitchStrings(NewString$, "«", "»")
        NewString$ = SwitchStrings(NewString$, "‹", "›")
        NewString$ = SwitchStrings(NewString$, "<", ">")
        CheckLen& = Len(NewString$)
        NewString$ = Left(NewString$, CheckLen& - 4)
        MacroFilter_Reverse$ = NewString$
    End If
End Function

Public Function MacroFilter_RoundedTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|¯¯", "|'˜¯")
    MyString$ = ReplaceString(MyString$, "¦¯¯", "¦'˜¯")
    MacroFilter_RoundedTop$ = MyString$
End Function

Public Function MacroFilter_Shadow(MyString As String) As String
    MyString$ = ReplaceString(MyString$, " |", ";|")
    MyString$ = ReplaceString(MyString$, "| ", "|;")
    MacroFilter_Shadow$ = MyString$
End Function

Public Function MacroFilter_Smear(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "|", "¦")
    MyString$ = ReplaceString(MyString$, "   ¦", ".:;¦")
    MyString$ = ReplaceString(MyString$, "  ¦", ":;¦")
    MyString$ = ReplaceString(MyString$, " ¦", ";¦")
    MyString$ = ReplaceString(MyString$, "   /", ".:;/")
    MyString$ = ReplaceString(MyString$, "  /", ":;/")
    MyString$ = ReplaceString(MyString$, " /", ";/")
    MyString$ = ReplaceString(MyString$, "   \", ".:;\")
    MyString$ = ReplaceString(MyString$, "  \", ":;\")
    MyString$ = ReplaceString(MyString$, " \", ";\")
    MyString$ = ReplaceString(MyString$, "   '", ".:;'")
    MyString$ = ReplaceString(MyString$, "  '", ":;'")
    MyString$ = ReplaceString(MyString$, " '", ";'")
    MacroFilter_Smear$ = MyString$
End Function

Public Function MacroFilter_SpikeBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "¸¡¸¡¸¡")
    MacroFilter_SpikeBottom$ = MyString$
End Function

Public Function MacroFilter_Straighten(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "}", "\")
    MyString$ = ReplaceString(MyString$, "{", "/")
    MyString$ = ReplaceString(MyString$, "]", "\")
    MyString$ = ReplaceString(MyString$, "[", "/")
    MyString$ = ReplaceString(MyString$, ")", "\")
    MyString$ = ReplaceString(MyString$, "(", "/")
    MacroFilter_Straighten$ = MyString$
End Function

Public Function MacroFilter_Stretch(MyString As String) As String
    Dim CurChar As Long, StretchIt As Long, MyText As String
    Dim NewLine As String, NumOfLines As Long, CheckLen As Long
    Dim CurLine As String, NewString As String
    If MyString$ <> "" Then
        NewLine$ = Chr(13) & Chr(10)
        MyText$ = MyString$ & NewLine$
        NumOfLines& = LineCount(MyText$)
        For StretchIt& = 1 To NumOfLines&
            CurLine$ = LineFromString(MyText, StretchIt&)
            CurLine$ = DoubleText(CurLine$)
            NewString$ = NewString$ & CurLine$ & NewLine$
        Next StretchIt&
        CheckLen& = Len(NewString$)
        NewString$ = Left(NewString$, CheckLen& - 4)
        MacroFilter_Stretch$ = NewString$
    End If
End Function

Public Function MacroFilter_StarTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "`**¯")
    MacroFilter_StarTop$ = MyString$
End Function

Public Function MacroFilter_ThickenBottom(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "___", "¸‚¸‚¸‚")
    MacroFilter_ThickenBottom$ = MyString$
End Function

Public Function MacroFilter_ThickenTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "”‘”‘”‘")
    MacroFilter_ThickenTop$ = MyString$
End Function

Public Function MacroFilter_TreadTop(MyString As String) As String
    MyString$ = ReplaceString(MyString$, "¯¯¯", "ª‘ª‘ª‘")
    MacroFilter_TreadTop$ = MyString$
End Function

Public Function MacroFilter_UnIndent(MyString As String) As String
    Dim OrgLen As Long, NewLine As String, NumOfLines As Long
    Dim OrgCount As Long, CurLine As String, NewString As String
    Dim SpaceIt As Long
    OrgLen& = Len(MyString$)
    NewLine$ = Chr(13) & Chr(10)
    MyString$ = MyString$ & NewLine$
    NumOfLines& = LineCount(MyString)
    OrgCount& = NumOfLines&
    For SpaceIt& = 1 To NumOfLines&
        CurLine$ = LineFromString(MyString$, SpaceIt&)
        If Len(CurLine$) < 1 Then
            NewString$ = NewString$ & CurLine$ & NewLine$
        Else
            NewString$ = NewString$ & Right(CurLine$, Len(CurLine$) - 1) & NewLine$
        End If
    Next SpaceIt&
    MyString$ = Left(NewString$, Len(NewString$) - 4)
    MacroFilter_UnIndent$ = MyString$
End Function

Public Function MacroFilter_UpsideDown(MyString As String) As String
    Dim CharCheck As Long, CurChar As Long, CurLine As String
    Dim FlipIt As Long, MyLine As Long, MyText As String
    Dim NewLine As String, NumOfLines As Long
    Dim CheckLen As Long, NewString As String
    If MyString$ <> "" Then
        NewLine$ = Chr(13) & Chr(10)
        MyText$ = MyString$ & NewLine$
        NumOfLines& = LineCount(MyText$)
        MyLine& = NumOfLines& - 1
        For FlipIt& = 1 To NumOfLines&
            DoEvents
            CurLine$ = LineFromString(MyText$, MyLine&)
            NewString$ = NewString$ & CurLine$ & NewLine$
            MyLine& = MyLine& - 1
        Next FlipIt&
        NewString$ = Left(NewString$, Len(NewString$) - 4)
        MyString$ = NewString$
        CheckLen& = Len(NewString$)
        NewString$ = SwitchStrings(MyString$, "/", "\")
        MyString$ = SwitchStrings(MyString$, "¯", "_")
        MyString$ = SwitchStrings(MyString$, ",", "'")
        MyString$ = ReplaceString(MyString$, ",,", ",")
        MyString$ = ReplaceString(MyString$, "`", ",")
        MyString$ = SwitchStrings(MyString$, "´", ".")
        MyString$ = ReplaceString(MyString$, "‘", ".")
        MyString$ = ReplaceString(MyString$, "’", ",")
        MyString$ = SwitchStrings(MyString$, "ˆ", "¸")
        MyString$ = SwitchStrings(MyString$, "„", Chr(34))
        MacroFilter_UpsideDown$ = MyString$
    End If
End Function

Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Sub LoadText(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub

Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub Loadlistbox(Directory As String, thelist As ListBox)
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

Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)
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

Public Sub SaveListBox(Directory As String, thelist As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
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

Public Function FileGetAttributes(theFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = dir(theFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(theFile$)
    End If
End Function

Public Sub FileSetNormal(theFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(theFile$)
    If SafeFile$ <> "" Then
        SetAttr theFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(theFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(theFile$)
    If SafeFile$ <> "" Then
        SetAttr theFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(theFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(theFile$)
    If SafeFile$ <> "" Then
        SetAttr theFile$, vbHidden
    End If
End Sub

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub

Public Function CheckIfMaster() As Boolean
    Dim AOL As Long, MDI As Long, pWindow As Long
    Dim pButton As Long, Modal As Long, mstatic As Long
    Dim mString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Call Keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Parental Controls")
        pButton& = FindWindowEx(pWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pWindow& <> 0& And pButton& <> 0&
    Pause 0.3
    Do
        DoEvents
        Call PostMessage(pButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(pButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 0.8
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        mstatic& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
        mString$ = GetText(mstatic&)
    Loop Until Modal& <> 0 And mstatic& <> 0& And mString$ <> ""
    mString$ = ReplaceString(mString$, Chr(10), "")
    mString$ = ReplaceString(mString$, Chr(13), "")
    If mString$ = "Set Parental Controls" Then
        CheckIfMaster = True
    Else
        CheckIfMaster = False
    End If
    Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
End Function

Public Function getcaption(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    getcaption$ = buffer$
End Function

Public Function GetListText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, buffer$)
    GetListText$ = buffer$
End Function

Public Function GetText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXT, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetText$ = buffer$
End Function

Public Sub Button(mbutton As Long)
    Call SendMessage(mbutton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mbutton&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub Icon(aIcon As Long)
    Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub CloseWindow(window As Long)
    Call PostMessage(window&, WM_CLOSE, 0&, 0&)
End Sub

Public Function ProfileGet(ScreenName As String) As String
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim nowindow As Long, OkButton As Long, CurPos As POINTAPI
    Dim winvis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(Toolbar&, toolicon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        smod& = FindWindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    Call PostMessage(smod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(smod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.x, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = GetText(pTextWindow&)
        nowindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or nowindow& <> 0&
    DoEvents
    If nowindow& <> 0& Then
        OkButton& = FindWindowEx(nowindow&, 0&, "Button", "OK")
        Call SendMessage(OkButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OkButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "< No Profile >"
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    End If
End Function

Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub Playwav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub

Public Sub SetText(window As Long, text As String)
    Call SendMessageByString(window&, WM_SETTEXT, 0&, text$)
End Sub

Public Function ListToMailString(thelist As ListBox) As String
    Dim DoList As Long, MailString As String
    If thelist.List(0) = "" Then Exit Function
    For DoList& = 0 To thelist.ListCount - 1
        MailString$ = MailString$ & "(" & thelist.List(DoList&) & "), "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToMailString$ = MailString$
End Function

Public Sub FormTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub FormDrag(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Public Sub FormExitDown(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) + 300))
    Loop Until TheForm.Top > 7200
End Sub

Public Sub FormExitLeft(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) - 300))
    Loop Until TheForm.Left < -TheForm.Width
End Sub

Public Sub FormExitRight(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) + 300))
    Loop Until TheForm.Left > Screen.Width
End Sub

Public Sub FormExitUp(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) - 300))
    Loop Until TheForm.Top < -TheForm.Width
End Sub

Public Sub HideWEL(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub

Public Sub ShowWEL(hwnd As Long)
    
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub

Public Sub RunMenu(TopMenu As Long, SubMenu As Long)
    Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub

Public Sub RunMenuByString(SearchString As String)
    Dim AOL As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(AOL&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Public Function AOLversion()
Dim AOL As Long, aMenu As Long, mCount As Long
    Dim SearchString As String
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            
            SearchString$ = "What's New in AOL 6.0"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            AOLversion = "Åøl V 6·º"
            Exit Function
            End If
            SearchString$ = "What's New in AOL 5.0"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            AOLversion = "Åøl V 5·º"
            Exit Function
            End If
            
            SearchString$ = "What's New in AOL 4.0"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            AOLversion = "Åøl V 4·º"
            Exit Function
            End If
        
            SearchString$ = "Send an Instant Message"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            AOLversion = "Åøl V 3·º"
            Exit Function
            End If
        
        Next LookSub&
    Next LookFor&
End Function
Sub Hide_Welcome()
Dim AOLFrame As String
Dim MDIClient As String
Dim AOLChild As String
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", "Welcome, " & (GetUser) & "!")
Call ShowWindow(AOLChild, SW_HIDE)
End Sub
Sub Show_Welcome()
Dim AOLFrame As String
Dim MDIClient As String
Dim AOLChild As String
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", "Welcome, " & (GetUser) & "!")
Call ShowWindow(AOLChild, SW_SHOW)
End Sub
Sub Change_AOL_Caption(text As String)
Dim AOLFrame As String
AOLFrame = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(AOLFrame, WM_SETTEXT, 0&, text$)
End Sub
Public Function Find_ChatRoom() As Long
Dim lngAOLframe As Long, lngMDIclient As Long, lngAOLchild As Long

lngAOLframe = FindWindow("aol frame25", vbNullString)
lngMDIclient = FindWindowEx(lngAOLframe, 0&, "mdiclient", vbNullString)
lngAOLchild = FindWindowEx(lngMDIclient, 0&, "aol child", vbNullString)

Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, Winkid7 As Long, Winkid8 As Long, Winkid9 As Long, FindOtherWin As Long

FindOtherWin = GetWindow(lngAOLchild, GW_HWNDFIRST)

Do While FindOtherWin <> 0
       DoEvents
       Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid2 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid3 = FindWindowEx(FindOtherWin, 0&, "_aol_combobox", vbNullString)
       Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       Winkid6 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
       Winkid7 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
       Winkid8 = FindWindowEx(FindOtherWin, 0&, "_aol_image", vbNullString)
       Winkid9 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
       If (Winkid1 <> 0) And _
          (Winkid2 <> 0) And _
          (Winkid3 <> 0) And _
          (Winkid4 <> 0) And _
          (Winkid5 <> 0) And _
          (Winkid6 <> 0) And _
          (Winkid7 <> 0) And _
          (Winkid8 <> 0) And _
          (Winkid9 <> 0) Then
              Find_ChatRoom = FindOtherWin
              Exit Function
       End If
       FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
Loop
Find_ChatRoom = 0
End Function
Public Function Text_ChatTitle() As String
    Dim lngTitle  As Long
    lngTitle = Find_ChatRoom
    Dim TheText As String, TL As Long
    TL = SendMessageLong(lngTitle, WM_GETTEXTLENGTH, 0&, 0&)
    TheText = String(TL + 1, " ")
    Call SendMessageByString(lngTitle, WM_GETTEXT, TL + 1, TheText)
    Text_ChatTitle = Left(TheText, TL)
End Function



Public Sub LoCkChAt()
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, "")
End Sub

Public Sub loCkKw()
 Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, "")
End Sub

Function TrimSpaces(text As String) As String
    Dim TheChar, TrimSpace
    Dim TheChars
    If InStr(text, " ") = 0 Then
        TrimSpaces = text
        Exit Function
    End If
    For TrimSpace = 1 To Len(text)
        TheChar = Mid(text, TrimSpace, 1)
        TheChars = TheChars & TheChar
        If TheChar = " " Then
            TheChars = Mid(TheChars, 1, Len(TheChars) - 1)
        End If
    Next TrimSpace
    TrimSpaces = TheChars
End Function
Public Function HideAOL()
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOLFrame&, SW_HIDE)
End Function

Public Function ShowAOL()
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOLFrame&, SW_SHOWNORMAL)
End Function
Public Sub aol4_killwait()
Dim AOL As Long, AOLModal As Long, AOLGlyph As Long
Dim AOLStatic As Long, AOLIcon As Long, AolInstance As Long
AOL& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
Call RunMenuByString("&About America Online")
Do: DoEvents
AOLModal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until AOLModal& <> 0& And AOLGlyph <> 0& And AOLStatic& <> 0& And AOLIcon& <> 0&
Do: DoEvents
AOLModal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
Call PostMessage(AOLModal&, WM_CLOSE, 0, 0&)
Loop Until AOLModal& = 0&
End Sub

 Public Sub killAOL()
 Dim AOL As String
 AOL = FindWindow("AOL Frame25", vbNullString)
    killwin AOL
End Sub
Sub killwin(Windo)
   Call SendMessageLong(Windo, WM_CLOSE, 0&, 0&)
End Sub
Function WavY(TheText As String)
Dim G As String
Dim a As String
Dim W As Long
Dim r As String
Dim U As String
Dim s As String
Dim t As String
Dim p As String
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    t$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & s$ & "</sub>" & t$
Next W
WavY = p$
End Function
Function EliteText(Word$)
Dim Made As String
Dim q As Long
Dim letter As String
Dim leet As String
Dim x As String
Made$ = ""
For q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, q, 1)
    leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If x = 1 Then leet$ = "â"
    If x = 2 Then leet$ = "å"
    If x = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If x = 1 Then leet$ = "ë"
    If x = 2 Then leet$ = "ê"
    If x = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If x = 1 Then leet$ = "ì"
    If x = 2 Then leet$ = "ï"
    If x = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = "J"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If x = 1 Then leet$ = "ô"
    If x = 2 Then leet$ = "ð"
    If x = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If x = 1 Then leet$ = "ù"
    If x = 2 Then leet$ = "û"
    If x = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "y"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If x = 1 Then leet$ = "Å"
    If x = 2 Then leet$ = "Ä"
    If x = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If x = 1 Then leet$ = "Ï"
    If x = 2 Then leet$ = "Î"
    If x = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If letter$ = "?" Then leet$ = "¿?"
    If letter$ = "1" Then leet$ = "¹"
    If letter$ = "2" Then leet$ = "²"
    If letter$ = "3" Then leet$ = "³"
    If letter$ = "0" Then leet$ = "º"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next q

EliteText = Made$

End Function
Public Function HideChatRoom()
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
Call ShowWindow(AOLChild&, SW_HIDE)
End Function
Public Function ShowChatRoom()
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
Call ShowWindow(AOLChild&, SW_SHOW)
End Function
Public Function ChangeChatCap(text As String)
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", (Text_ChatTitle))
Call SendMessageByString(AOLChild&, WM_SETTEXT, 0&, text$)
End Function
