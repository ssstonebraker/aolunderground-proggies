Attribute VB_Name = "Dayz32"
' ==============================================
' dayz32 was made by dayz
' it is combination of many subs for aol 5,6,7
' shout outs go to
' Bigal,PC,demo,neo,bofen,zone,zap,kelly,laser,
' xeek,light,sage,blaze,antoney,haze,fire,steez,gibs
' and anyone else i forgot sry
' ==============================================
' ;;;;;;;;;;  ;;;;;;; ;;;;; ;;;;; ;;;;;;;;;;
' ;;;;  ;;;;  ;;; ;;; ;;;;; ;;;;;       ;;;
' ;;;;  ;;;;  ;;   ;; ;;;;; ;;;;;      ;;;
' ;;;;  ;;;;  ;;;;;;;  ;;;; ;;;;      ;;;
' ;;;;  ;;;;  ;;; ;;;   ;;; ;;;      ;;;
' ;;;;  ;;;;  ;;; ;;;    ;; ;;      ;;;
' ;;;;  ;;;;  ;;; ;;;     ;;;       ;;;;;;;;;;;
' ;;;;;;;;;;              ;;;
' ==============================================

Option Explicit

Private Declare Function closehandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Public Declare Sub copymemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Public Declare Function findwindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function findwindowex Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function getcursorpos Lib "user32" Alias "GetCursorPos" (lpPoint As pointapi) As Long
Public Declare Function getmenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
Public Declare Function getmenuitemcount Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Public Declare Function getmenuitemid Lib "user32" Alias "GetMenuItemID" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function getmenustring Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function getsubmenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function getwindowtext Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function getwindowtextlength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function getwindowthreadprocessid Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function iswindowvisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
Public Declare Function openprocess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function postmessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function readprocessmemory Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function sendmessagelong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function setcursorpos Lib "user32" Alias "SetCursorPos" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function setwindowpos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uflags As Long) As Long
Public Declare Function releasecapture Lib "user32" Alias "ReleaseCapture" () As Long
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long

Declare Function enablewindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal cmd As Long) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const hwnd_notopmost = -2
Public Const hwnd_topmost = -1

Public Const lb_getcount = &H18B
Public Const lb_getitemdata = &H199
Public Const lb_gettext = &H189
Public Const lb_gettextlen = &H18A
Public Const lb_setcursel = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const sw_hide = 0
Public Const sw_show = 5

Public Const swp_nomove = &H2
Public Const swp_nosize = &H1

Public Const vk_down = &H28
Public Const vk_left = &H25
Public Const VK_MENU = &H12
Public Const vk_return = &HD
Public Const vk_right = &H27
Public Const VK_SHIFT = &H10
Public Const vk_space = &H20
Public Const vk_up = &H26

Public Const wm_char = &H102
Public Const wm_close = &H10
Public Const wm_command = &H111
Public Const wm_gettext = &HD
Public Const wm_gettextlength = &HE
Public Const wm_keydown = &H100
Public Const wm_keyup = &H101
Public Const wm_lbuttondblclk = &H203
Public Const wm_lbuttondown = &H201
Public Const wm_lbuttonup = &H202
Public Const wm_move = &HF012
Public Const wm_settext = &HC
Public Const wm_syscommand = &H112

Public Const process_read = &H10
Public Const rights_required = &HF0000

Public Const enter_key = 13
Public Const flags = swp_nomove Or swp_nosize

Public Type pointapi
        X As Long
        Y As Long
End Type

Public Function FindForwardWindow() As Long
    Dim aol As Long, mdi As Long, Child As Long
    Dim Rich1 As Long, Rich2 As Long, Combo As Long
    Dim FontCombo As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    Rich1& = findwindowex(Child&, 0&, "RICHCNTL", vbNullString)
    Rich2& = findwindowex(Child&, Rich1&, "RICHCNTL", vbNullString)
    Combo& = findwindowex(Child&, 0&, "_AOL_Combobox", vbNullString)
    FontCombo& = findwindowex(Child&, 0&, "_AOL_FontCombo", vbNullString)
    If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
        FindForwardWindow& = Child&
        Exit Function
    Else
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            Rich1& = findwindowex(Child&, 0&, "RICHCNTL", vbNullString)
            Rich2& = findwindowex(Child&, Rich1&, "RICHCNTL", vbNullString)
            Combo& = findwindowex(Child&, 0&, "_AOL_Combobox", vbNullString)
            FontCombo& = findwindowex(Child&, 0&, "_AOL_FontCombo", vbNullString)
            If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
                FindForwardWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindForwardWindow& = 0&
End Function

Public Function FindSendWindow() As Long
    Dim aol As Long, mdi As Long, Child As Long
    Dim SendStatic As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    SendStatic& = findwindowex(Child&, 0&, "_AOL_Static", "Send Now")
    If SendStatic& <> 0& Then
        FindSendWindow& = Child&
        Exit Function
    Else
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            SendStatic& = findwindowex(Child&, 0&, "_AOL_Static", "Send Now")
            If SendStatic& <> 0& Then
                FindSendWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindSendWindow& = 0&
End Function

Public Sub MailOpenFlash()
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As pointapi, WinVis As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call getcursorpos(CurPos)
    Call setcursorpos(Screen.Width, Screen.Height)
    Call postmessage(ToolIcon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ToolIcon&, wm_lbuttonup, 0&, 0&)
    Do
        sMod& = findwindow("#32768", vbNullString)
        WinVis& = iswindowvisible(sMod&)
    Loop Until WinVis& = 1
    Call postmessage(sMod&, wm_keydown, vk_up, 0&)
    Call postmessage(sMod&, wm_keyup, vk_up, 0&)
    Call postmessage(sMod&, wm_keydown, vk_right, 0&)
    Call postmessage(sMod&, wm_keyup, vk_right, 0&)
    Call postmessage(sMod&, wm_keydown, vk_return, 0&)
    Call postmessage(sMod&, wm_keyup, vk_return, 0&)
    Call setcursorpos(CurPos.X, CurPos.Y)
End Sub

Public Sub MailOpenNew()
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, sMod As Long, CurPos As pointapi
    Dim WinVis As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call getcursorpos(CurPos)
    Call setcursorpos(Screen.Width, Screen.Height)
    Call postmessage(ToolIcon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ToolIcon&, wm_lbuttonup, 0&, 0&)
    Do
        sMod& = findwindow("#32768", vbNullString)
        WinVis& = iswindowvisible(sMod&)
    Loop Until WinVis& = 1
    Call postmessage(sMod&, wm_keydown, vk_down, 0&)
    Call postmessage(sMod&, wm_keyup, vk_down, 0&)
    Call postmessage(sMod&, wm_keydown, vk_down, 0&)
    Call postmessage(sMod&, wm_keyup, vk_down, 0&)
    Call postmessage(sMod&, wm_keydown, vk_return, 0&)
    Call postmessage(sMod&, wm_keyup, vk_return, 0&)
    Call setcursorpos(CurPos.X, CurPos.Y)
End Sub

Public Sub MailOpenOld()
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As pointapi, WinVis As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call getcursorpos(CurPos)
    Call setcursorpos(Screen.Width, Screen.Height)
    Call postmessage(ToolIcon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ToolIcon&, wm_lbuttonup, 0&, 0&)
    Do
        sMod& = findwindow("#32768", vbNullString)
        WinVis& = iswindowvisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 4
        Call postmessage(sMod&, wm_keydown, vk_down, 0&)
        Call postmessage(sMod&, wm_keyup, vk_down, 0&)
    Next DoThis&
    Call postmessage(sMod&, wm_keydown, vk_return, 0&)
    Call postmessage(sMod&, wm_keyup, vk_return, 0&)
    Call setcursorpos(CurPos.X, CurPos.Y)
End Sub

Public Sub MailOpenSent()
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As pointapi, WinVis As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call getcursorpos(CurPos)
    Call setcursorpos(Screen.Width, Screen.Height)
    Call postmessage(ToolIcon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ToolIcon&, wm_lbuttonup, 0&, 0&)
    Do
        sMod& = findwindow("#32768", vbNullString)
        WinVis& = iswindowvisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 5
        Call postmessage(sMod&, wm_keydown, vk_down, 0&)
        Call postmessage(sMod&, wm_keyup, vk_down, 0&)
    Next DoThis&
    Call postmessage(sMod&, wm_keydown, vk_return, 0&)
    Call postmessage(sMod&, wm_keyup, vk_return, 0&)
    Call setcursorpos(CurPos.X, CurPos.Y)
End Sub

Public Sub MailOpenEmailFlash(Index As Long)
    Dim aol As Long, mdi As Long, fMail As Long, flist As Long
    Dim fcount As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = findwindowex(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    flist& = findwindowex(fMail&, 0&, "_AOL_Tree", vbNullString)
    fcount& = sendmessage(flist&, lb_getcount, 0&, 0&)
    If fcount& < Index& Then Exit Sub
    Call sendmessage(flist&, lb_setcursel, Index&, 0&)
    Call postmessage(flist&, wm_keydown, vk_return, 0&)
    Call postmessage(flist&, wm_keyup, vk_return, 0&)
End Sub

Public Sub MailOpenEmailNew(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& < Index& Then Exit Sub
    Call sendmessage(mTree&, lb_setcursel, Index&, 0&)
    Call postmessage(mTree&, wm_keydown, vk_return, 0&)
    Call postmessage(mTree&, wm_keyup, vk_return, 0&)
End Sub

Public Sub MailOpenEmailOld(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& < Index& Then Exit Sub
    Call sendmessage(mTree&, lb_setcursel, Index&, 0&)
    Call postmessage(mTree&, wm_keydown, vk_return, 0&)
    Call postmessage(mTree&, wm_keyup, vk_return, 0&)
End Sub

Public Sub MailOpenEmailSent(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& < Index& Then Exit Sub
    Call sendmessage(mTree&, lb_setcursel, Index&, 0&)
    Call postmessage(mTree&, wm_keydown, vk_return, 0&)
    Call postmessage(mTree&, wm_keyup, vk_return, 0&)
End Sub

Public Function MailCountFlash() As Long
    Dim aol As Long, mdi As Long, fMail As Long, flist As Long
    Dim count As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = findwindowex(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    flist& = findwindowex(fMail&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(flist&, lb_getcount, 0&, 0&)
    MailCountFlash& = count&
End Function

Public Sub MailToListFlash(TheList As ListBox)
    Dim aol As Long, mdi As Long, fMail As Long, flist As Long
    Dim count As Long, MyString As String, AddMails As Long
    Dim sLength As Long, Spot As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = findwindowex(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    If fMail& = 0& Then Exit Sub
    flist& = findwindowex(fMail&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(flist&, lb_getcount, 0&, 0&)
    MyString$ = String(255, 0)
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = sendmessage(flist&, lb_gettextlen, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call sendmessagebystring(flist&, lb_gettext, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = right(MyString$, Len(MyString$) - Spot&)
        MyString$ = ReplaceString(MyString$, Chr(0), "")
        TheList.AddItem MyString$
    Next AddMails&
End Sub

Public Function FindMailBox() As Long
    Dim aol As Long, mdi As Long, Child As Long
    Dim TabControl As Long, TabPage As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    TabControl& = findwindowex(Child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = Child&
        Exit Function
    Else
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            TabControl& = findwindowex(Child&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
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
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    MailCountNew& = count&
End Function

Public Function MailCountSent() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    MailCountSent& = count&
End Function

Public Function MailCountOld() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    MailCountOld& = count&
End Function

Public Sub MailDeleteNewByIndex(Index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If Index& > count& - 1 Or Index& < 0& Then Exit Sub
    Call sendmessage(mTree&, lb_setcursel, Index&, 0&)
    dButton& = findwindowex(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    Call sendmessage(dButton&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(dButton&, wm_lbuttonup, 0&, 0&)
End Sub

Public Sub MailDeleteNewDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String, cSubject As String
    Dim SearchFor As Long, sSender As String, sSubject As String
    Dim CurCaption As String
    MailBox& = FindMailBox&
    CurCaption$ = VBForm.Caption
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = findwindowex(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchFor& = 0& To count& - 2
        DoEvents
        sSender$ = MailSenderNew(SearchFor&)
        sSubject$ = MailSubjectNew(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To count& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Now checking #" & SearchFor& & " for match with #" & SearchBox&
            End If
            cSender$ = MailSenderNew(SearchBox&)
            cSubject$ = MailSubjectNew(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call sendmessage(mTree&, lb_setcursel, SearchBox&, 0&)
                DoEvents
                Call sendmessage(dButton&, wm_lbuttondown, 0&, 0&)
                Call sendmessage(dButton&, wm_lbuttonup, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub

Public Sub MailDeleteNewBySender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = findwindowex(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchBox& = 0& To count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If (cSender$) = (Sender$) Then
            Call sendmessage(mTree&, lb_setcursel, SearchBox&, 0&)
            DoEvents
            Call sendmessage(dButton&, wm_lbuttondown, 0&, 0&)
            Call sendmessage(dButton&, wm_lbuttonup, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub

Public Sub MailDeleteNewNotSender(Sender As String)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    Dim SearchBox As Long, cSender As String
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    dButton& = findwindowex(MailBox&, 0&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    dButton& = findwindowex(MailBox&, dButton&, "_AOL_Icon", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchBox& = 0& To count& - 1
        cSender$ = MailSenderNew(SearchBox&)
        If cSender$ = "" Then Exit Sub
        If (cSender$) <> (Sender$) Then
            Call sendmessage(mTree&, lb_setcursel, SearchBox&, 0&)
            DoEvents
            Call sendmessage(dButton&, wm_lbuttondown, 0&, 0&)
            Call sendmessage(dButton&, wm_lbuttonup, 0&, 0&)
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub

Public Function MailSenderFlash(Index As Long) As String
    Dim aol As Long, mdi As Long, fMail As Long, flist As Long
    Dim fcount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, Spot1 As Long, Spot2 As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = findwindowex(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    flist& = findwindowex(fMail&, 0&, "_AOL_Tree", vbNullString)
    fcount& = sendmessage(flist&, lb_getcount, 0&, 0&)
    If fcount& < Index& Then Exit Function
    DeleteButton& = findwindowex(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fcount& = 0 Or Index& > fcount& - 1 Or Index& < 0& Then Exit Function
    sLength& = sendmessage(flist&, lb_gettextlen, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call sendmessagebystring(flist&, lb_gettext, Index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderFlash$ = MyString$
End Function

Public Function MailSenderNew(Index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot1 As Long, Spot2 As Long, MyString As String
    Dim count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0 Or Index& > count& - 1 Or Index& < 0& Then Exit Function
    sLength& = sendmessage(mTree&, lb_gettextlen, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call sendmessagebystring(mTree&, lb_gettext, Index&, MyString$)
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderNew$ = MyString$
End Function

Public Function MailSubjectFlash(Index As Long) As String
    Dim aol As Long, mdi As Long, fMail As Long, flist As Long
    Dim fcount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, Spot As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = findwindowex(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    flist& = findwindowex(fMail&, 0&, "_AOL_Tree", vbNullString)
    fcount& = sendmessage(flist&, lb_getcount, 0&, 0&)
    If fcount& < Index& Then Exit Function
    DeleteButton& = findwindowex(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    If fcount& = 0 Or Index& > fcount& - 1 Or Index& < 0& Then Exit Function
    sLength& = sendmessage(flist&, lb_gettextlen, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call sendmessagebystring(flist&, lb_gettext, Index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectFlash$ = MyString$
End Function

Public Function MailSubjectNew(Index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0 Or Index& > count& - 1 Or Index& < 0& Then Exit Function
    sLength& = sendmessage(mTree&, lb_gettextlen, Index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call sendmessagebystring(mTree&, lb_gettext, Index&, MyString$)
    Spot& = InStr(MyString$, Chr(9))
    Spot& = InStr(Spot& + 1, MyString$, Chr(9))
    MyString$ = right(MyString$, Len(MyString$) - Spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectNew$ = MyString$
End Function

Public Sub MailToListNew(TheList As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = sendmessage(mTree&, lb_gettextlen, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call sendmessagebystring(mTree&, lb_gettext, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem MyString$
    Next AddMails&
End Sub

Public Sub MailToListOld(TheList As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = sendmessage(mTree&, lb_gettextlen, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call sendmessagebystring(mTree&, lb_gettext, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem MyString$
    Next AddMails&
End Sub

Public Sub MailToListSent(TheList As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = findwindowex(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = findwindowex(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = findwindowex(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = findwindowex(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = sendmessage(mTree&, lb_getcount, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = sendmessage(mTree&, lb_gettextlen, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call sendmessagebystring(mTree&, lb_gettext, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem MyString$
    Next AddMails&
End Sub

Public Sub sendmail(Person As String, subject As String, message As String)
    Dim aol As Long, mdi As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, errorwindow As Long
    Dim Button1 As Long, Button2 As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call postmessage(ToolIcon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ToolIcon&, wm_lbuttonup, 0&, 0&)
    DoEvents
    Do
        DoEvents
        OpenSend& = findwindowex(mdi&, 0&, "AOL Child", "Write Mail")
        EditTo& = findwindowex(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = findwindowex(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = findwindowex(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = findwindowex(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = findwindowex(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = findwindowex(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = findwindowex(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = findwindowex(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = findwindowex(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call sendmessagebystring(EditTo&, wm_settext, 0, Person$)
    DoEvents
    Call sendmessagebystring(EditSubject&, wm_settext, 0, subject$)
    DoEvents
    Call sendmessagebystring(Rich&, wm_settext, 0, message$)
    DoEvents
    pause 0.2
    Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = findwindow("AOL Frame25", vbNullString)
MDIClient& = findwindowex(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = findwindowex(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = findwindowex(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 15&
    AOLIcon& = findwindowex(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call sendmessage(AOLIcon&, wm_lbuttondown, 0&, 0&)
Call sendmessage(AOLIcon&, wm_lbuttonup, 0&, 0&)




 


End Sub

Public Sub MailForward(SendTo As String, message As String, DeleteFwd As Boolean)
    Dim aol As Long, mdi As Long, Error As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
    Dim Combo As Long, Button1 As Long, Button2 As Long
    Dim TempSubject As String
    OpenForward& = FindForwardWindow
    If OpenForward& = 0 Then Exit Sub
    SendButton& = findwindowex(OpenForward&, 0&, "_AOL_Icon", vbNullString)
    For DoIt& = 1 To 6
        SendButton& = findwindowex(OpenForward&, SendButton&, "_AOL_Icon", vbNullString)
    Next DoIt&
    Call sendmessage(SendButton&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(SendButton&, wm_lbuttonup, 0&, 0&)
    Do
        DoEvents
        OpenSend& = FindSendWindow
        EditTo& = findwindowex(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = findwindowex(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = findwindowex(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        Rich& = findwindowex(OpenSend&, 0&, "RICHCNTL", vbNullString)
        Combo& = findwindowex(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = findwindowex(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = findwindowex(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = findwindowex(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = findwindowex(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    If DeleteFwd = True Then
        TempSubject$ = gettext(EditSubject&)
        TempSubject$ = right(TempSubject$, Len(TempSubject$) - 5)
        Call sendmessagebystring(EditSubject&, wm_settext, 0, TempSubject$)
        DoEvents
    End If
    Call sendmessagebystring(EditTo&, wm_settext, 0, SendTo$)
    DoEvents
    Call sendmessagebystring(Rich&, wm_settext, 0, message$)
    DoEvents
    Do Until OpenSend& = 0& Or Error& <> 0&
        DoEvents
        aol& = findwindow("AOL Frame25", vbNullString)
        mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
        Error& = findwindowex(mdi&, 0&, "AOL Child", "Error")
        OpenSend& = FindSendWindow
        SendButton& = findwindowex(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = findwindowex(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        Call sendmessage(SendButton&, wm_lbuttondown, 0&, 0&)
        Call sendmessage(SendButton&, wm_lbuttonup, 0&, 0&)
        pause 1
    Loop
    If OpenSend& = 0& Then Call postmessage(OpenForward&, wm_close, 0&, 0&)
End Sub

Public Sub CloseOpenMails()
    Dim OpenSend As Long, OpenForward As Long
    Do
        DoEvents
        OpenSend& = FindSendWindow
        OpenForward& = FindForwardWindow
        Call postmessage(OpenSend&, wm_close, 0&, 0&)
        DoEvents
        Call postmessage(OpenForward&, wm_close, 0&, 0&)
        DoEvents
    Loop Until OpenSend& = 0& And OpenForward& = 0&
End Sub

Public Sub MailDeleteFlashByIndex(Index As Long)
    Dim aol As Long, mdi As Long, fMail As Long, flist As Long
    Dim fcount As Long, DeleteButton As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = findwindowex(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    flist& = findwindowex(fMail&, 0&, "_AOL_Tree", vbNullString)
    fcount& = sendmessage(flist&, lb_getcount, 0&, 0&)
    If fcount& < Index& Then Exit Sub
    DeleteButton& = findwindowex(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    Call sendmessage(flist&, lb_setcursel, Index&, 0&)
    Call sendmessage(DeleteButton&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(DeleteButton&, wm_lbuttonup, 0&, 0&)
End Sub

Public Sub MailDeleteFlashDuplicates(VBForm As Form, DisplayStatus As Boolean)
    Dim aol As Long, mdi As Long, fMail As Long, flist As Long
    Dim fcount As Long, DeleteButton As Long, SearchFor As Long
    Dim SearchBox As Long, CurCaption As String
    Dim sSender As String, sSubject As String
    Dim cSender As String, cSubject As String
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = findwindowex(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
    flist& = findwindowex(fMail&, 0&, "_AOL_Tree", vbNullString)
    fcount& = sendmessage(flist&, lb_getcount, 0&, 0&)
    If fcount& < 2& Then Exit Sub
    DeleteButton& = findwindowex(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = findwindowex(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    CurCaption$ = VBForm.Caption
    If fcount& = 0& Then Exit Sub
    For SearchFor& = 0& To fcount& - 2
        DoEvents
        sSender$ = MailSenderFlash(SearchFor&)
        sSubject$ = MailSubjectFlash(SearchFor&)
        If sSender$ = "" Then
            VBForm.Caption = CurCaption$
            Exit Sub
        End If
        For SearchBox& = SearchFor& + 1 To fcount& - 1
            If DisplayStatus = True Then
                VBForm.Caption = "Checking #" & SearchFor& & " with #" & SearchBox&
            End If
            cSender$ = MailSenderFlash(SearchBox&)
            cSubject$ = MailSubjectFlash(SearchBox&)
            If cSender$ = sSender$ And cSubject$ = sSubject$ Then
                Call sendmessage(flist&, lb_setcursel, SearchBox&, 0&)
                DoEvents
                Call sendmessage(DeleteButton&, wm_lbuttondown, 0&, 0&)
                Call sendmessage(DeleteButton&, wm_lbuttonup, 0&, 0&)
                DoEvents
                SearchBox& = SearchBox& - 1
            End If
        Next SearchBox&
    Next SearchFor&
    VBForm.Caption = CurCaption$
End Sub

Public Sub setmailprefs()
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim mdi As Long, mPrefs As Long, mButton As Long
    Dim gStatic As Long, mstatic As Long, fStatic As Long
    Dim maStatic As Long, dMod As Long, ConfirmCheck As Long
    Dim CloseCheck As Long, SpellCheck As Long, OKButton As Long
    Dim CurPos As pointapi, WinVis As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call getcursorpos(CurPos)
    Call setcursorpos(Screen.Width, Screen.Height)
    Call postmessage(ToolIcon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ToolIcon&, wm_lbuttonup, 0&, 0&)
    Do
        sMod& = findwindow("#32768", vbNullString)
        WinVis& = iswindowvisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 3
        Call postmessage(sMod&, wm_keydown, vk_down, 0&)
        Call postmessage(sMod&, wm_keyup, vk_down, 0&)
    Next DoThis&
    Call postmessage(sMod&, wm_keydown, vk_return, 0&)
    Call postmessage(sMod&, wm_keyup, vk_return, 0&)
    Call setcursorpos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        mPrefs& = findwindowex(mdi&, 0&, "AOL Child", "Preferences")
        gStatic& = findwindowex(mPrefs&, 0&, "_AOL_Static", "General")
        mstatic& = findwindowex(mPrefs&, 0&, "_AOL_Static", "Mail")
        fStatic& = findwindowex(mPrefs&, 0&, "_AOL_Static", "Font")
        maStatic& = findwindowex(mPrefs&, 0&, "_AOL_Static", "Marketing")
    Loop Until mPrefs& <> 0& And gStatic& <> 0& And mstatic& <> 0& And fStatic& <> 0& And maStatic& <> 0&
    mButton& = findwindowex(mPrefs&, 0&, "_AOL_Icon", vbNullString)
    mButton& = findwindowex(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    mButton& = findwindowex(mPrefs&, mButton&, "_AOL_Icon", vbNullString)
    Do
        DoEvents
        Call sendmessage(mButton&, wm_lbuttondown, 0&, 0&)
        Call sendmessage(mButton&, wm_lbuttonup, 0&, 0&)
        dMod& = findwindow("_AOL_Modal", "Mail Preferences")
        pause 0.6
    Loop Until dMod& <> 0&
    ConfirmCheck& = findwindowex(dMod&, 0&, "_AOL_Checkbox", vbNullString)
    CloseCheck& = findwindowex(dMod&, ConfirmCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = findwindowex(dMod&, CloseCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = findwindowex(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = findwindowex(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    SpellCheck& = findwindowex(dMod&, SpellCheck&, "_AOL_Checkbox", vbNullString)
    OKButton& = findwindowex(dMod&, 0&, "_AOL_icon", vbNullString)
    Call sendmessage(ConfirmCheck&, BM_SETCHECK, False, vbNullString)
    Call sendmessage(CloseCheck&, BM_SETCHECK, True, vbNullString)
    Call sendmessage(SpellCheck&, BM_SETCHECK, False, vbNullString)
    Call sendmessage(OKButton&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(OKButton&, wm_lbuttonup, 0&, 0&)
    DoEvents
    Call postmessage(mPrefs&, wm_close, 0&, 0&)
End Sub

Public Function ErrorName(Name As Long) As String
    Dim aol As Long, mdi As Long, errorwindow As Long
    Dim errortextwindow As Long, errorstring As String
    Dim NameCount As Long, TempString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    errorwindow& = findwindowex(mdi&, 0&, "AOL Child", "Error")
    If errorwindow& = 0& Then Exit Function
    errortextwindow& = findwindowex(errorwindow&, 0&, "_AOL_View", vbNullString)
    errorstring$ = gettext(errortextwindow&)
    NameCount& = linecount(errorstring$) - 2
    If NameCount& < Name& Then Exit Function
    TempString$ = linefromstring(errorstring$, Name& + 2)
    TempString$ = left(TempString$, InStr(TempString$, "-") - 2)
    ErrorName$ = TempString$
End Function

Public Function ErrorNameCount() As Long
    Dim aol As Long, mdi As Long, errorwindow As Long
    Dim errortextwindow As Long, errorstring As String
    Dim NameCount As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    errorwindow& = findwindowex(mdi&, 0&, "AOL Child", "Error")
    If errorwindow& = 0& Then Exit Function
    errortextwindow& = findwindowex(errorwindow&, 0&, "_AOL_View", vbNullString)
    errorstring$ = gettext(errortextwindow&)
    NameCount& = linecount(errorstring$) - 2
    ErrorNameCount& = NameCount&
End Function

Public Function checkalive(screenname As String) As Boolean
    Dim aol As Long, mdi As Long, errorwindow As Long
    Dim errortextwindow As Long, errorstring As String
    Dim mailwindow As Long, nowindow As Long, NoButton As Long
    Call sendmail("*, " & screenname$, "You alive?", "=)")
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        errorwindow& = findwindowex(mdi&, 0&, "AOL Child", "Error")
        errortextwindow& = findwindowex(errorwindow&, 0&, "_AOL_View", vbNullString)
        errorstring$ = gettext(errortextwindow&)
    Loop Until errorwindow& <> 0 And errortextwindow& <> 0 And errorstring$ <> ""
    If InStr((ReplaceString(errorstring$, " ", "")), (ReplaceString(screenname$, " ", ""))) > 0 Then
        checkalive = False
    Else
        checkalive = True
    End If
    mailwindow& = findwindowex(mdi&, 0&, "AOL Child", "Write Mail")
    Call postmessage(errorwindow&, wm_close, 0&, 0&)
    DoEvents
    Call postmessage(mailwindow&, wm_close, 0&, 0&)
    DoEvents
    Do
        DoEvents
        nowindow& = findwindow("#32770", "America Online")
        NoButton& = findwindowex(nowindow&, 0&, "Button", "&No")
    Loop Until nowindow& <> 0& And NoButton& <> 0
    Call sendmessage(NoButton&, wm_keydown, vk_space, 0&)
    Call sendmessage(NoButton&, wm_keyup, vk_space, 0&)
End Function

Public Sub chatsend(Chat As String)
    Dim room As Long, AORich As Long, AORich2 As Long
    room& = findroom&
    AORich& = findwindowex(room, 0&, "RICHCNTL", vbNullString)
    AORich2& = findwindowex(room, AORich, "RICHCNTL", vbNullString)
    Call sendmessagebystring(AORich2, wm_settext, 0&, Chat$)
    Call sendmessagelong(AORich2, wm_char, enter_key, 0&)
End Sub

Public Function FindIM() As Long
    Dim aol As Long, mdi As Long, Child As Long, Caption As String
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(Child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        FindIM& = Child&
        Exit Function
    Else
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(Child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                FindIM& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindIM& = Child&
End Function

Public Function findroom() As Long
    Dim aol As Long, mdi As Long, Child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    Rich& = findwindowex(Child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = findwindowex(Child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = findwindowex(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = findwindowex(Child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        findroom& = Child&
        Exit Function
    Else
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            Rich& = findwindowex(Child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = findwindowex(Child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = findwindowex(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = findwindowex(Child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                findroom& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    findroom& = Child&
End Function

Public Function FindInfoWindow() As Long
    Dim aol As Long, mdi As Long, Child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    AOLCheck& = findwindowex(Child&, 0&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = findwindowex(Child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = findwindowex(Child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = findwindowex(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = findwindowex(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = Child&
        Exit Function
    Else
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            AOLCheck& = findwindowex(Child&, 0&, "_AOL_Checkbox", vbNullString)
            AOLStatic& = findwindowex(Child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = findwindowex(Child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = findwindowex(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = findwindowex(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindInfoWindow& = Child&
End Function

Public Function RoomCount() As Long
    Dim aol As Long, mdi As Long, rMail As Long, rlist As Long
    Dim count As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    rMail& = findroom
    rlist& = findwindowex(rMail&, 0&, "_AOL_Listbox", vbNullString)
    count& = sendmessage(rlist&, lb_getcount, 0&, 0&)
    RoomCount& = count&
End Function

Public Function ChatSend2(text As String)
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim AOLIcon As Long
AOLFrame& = findwindow("AOL Frame25", vbNullString)
MDIClient& = findwindowex(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = findwindowex(MDIClient&, 0&, "AOL Child", "vb6")
AOLEdit& = findwindowex(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call sendmessagebystring(AOLEdit&, wm_settext, 0&, text$)
pause 0.2
AOLFrame& = findwindow("AOL Frame25", vbNullString)
MDIClient& = findwindowex(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = findwindowex(MDIClient&, 0&, "AOL Child", "vb6")
AOLIcon& = findwindowex(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call sendmessage(AOLIcon&, wm_lbuttondown, 0&, 0&)
Call sendmessage(AOLIcon&, wm_lbuttonup, 0&, 0&)

End Function
Public Sub AddRoomToListbox(TheList As ListBox, adduser As Boolean)
    On Error Resume Next
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnhold As Long, rbytes As Long, Index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    room& = findroom&
    If room& = 0& Then Exit Sub
    rlist& = findwindowex(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = getwindowthreadprocessid(rlist, cprocess&)
    mthread& = openprocess(process_read Or rights_required, False, cprocess&)
    If mthread& Then
        For Index& = 0 To sendmessage(rlist, lb_getcount, 0, 0) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = sendmessage(rlist, lb_getitemdata, ByVal CLng(Index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call readprocessmemory(mthread&, itmhold&, screenname$, 4, rbytes)
            Call copymemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call readprocessmemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ Or adduser = True Then
                TheList.AddItem screenname$
            End If
        Next Index&
        Call closehandle(mthread)
    End If
End Sub

Public Sub AddRoomToCombobox(TheCombo As ComboBox, adduser As Boolean)
    On Error Resume Next
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnhold As Long, rbytes As Long, Index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    room& = findroom&
    If room& = 0& Then Exit Sub
    rlist& = findwindowex(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = getwindowthreadprocessid(rlist, cprocess&)
    mthread& = openprocess(process_read Or rights_required, False, cprocess&)
    If mthread& Then
        For Index& = 0 To sendmessage(rlist, lb_getcount, 0, 0) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = sendmessage(rlist, lb_getitemdata, ByVal CLng(Index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call readprocessmemory(mthread&, itmhold&, screenname$, 4, rbytes)
            Call copymemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call readprocessmemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ Or adduser = True Then
                TheCombo.AddItem screenname$
            End If
        Next Index&
        Call closehandle(mthread)
    End If
    If TheCombo.ListCount > 0 Then
        TheCombo.text = TheCombo.list(0)
    End If
End Sub

Public Sub ChatIgnoreByIndex(Index As Long)
    Dim room As Long, slist As Long, iWindow As Long
    Dim iCheck As Long, a As Long, count As Long
    count& = RoomCount&
    If Index& > count& - 1 Then Exit Sub
    room& = findroom&
    slist& = findwindowex(room&, 0&, "_AOL_Listbox", vbNullString)
    Call sendmessage(slist&, lb_setcursel, Index&, 0&)
    Call postmessage(slist&, wm_lbuttondblclk, 0&, 0&)
    Do
        DoEvents
        iWindow& = FindInfoWindow
    Loop Until iWindow& <> 0&
    DoEvents
    iCheck& = findwindowex(iWindow&, 0&, "_AOL_Checkbox", vbNullString)
    DoEvents
    Do
        DoEvents
        a& = sendmessage(iCheck&, BM_GETCHECK, 0&, 0&)
        Call postmessage(iCheck&, wm_lbuttondown, 0&, 0&)
        DoEvents
        Call postmessage(iCheck&, wm_lbuttonup, 0&, 0&)
        DoEvents
    Loop Until a& <> 0&
    DoEvents
    Call postmessage(iWindow&, wm_close, 0&, 0&)
End Sub

Public Sub ChatIgnoreByName(Name As String)
    On Error Resume Next
    Dim cprocess As Long, itmhold As Long, screenname As String
    Dim psnhold As Long, rbytes As Long, Index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim lIndex As Long
    room& = findroom&
    If room& = 0& Then Exit Sub
    rlist& = findwindowex(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = getwindowthreadprocessid(rlist, cprocess&)
    mthread& = openprocess(process_read Or rights_required, False, cprocess&)
    If mthread& Then
        For Index& = 0 To sendmessage(rlist, lb_getcount, 0, 0) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = sendmessage(rlist, lb_getitemdata, ByVal CLng(Index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call readprocessmemory(mthread&, itmhold&, screenname$, 4, rbytes)
            Call copymemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call readprocessmemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ And (screenname$) = (Name$) Then
                lIndex& = Index&
                Call ChatIgnoreByIndex(lIndex&)
                 Call ChatIgnoreByIndex(Index&)
                DoEvents
                Exit Sub
            End If
        Next Index&
        Call closehandle(mthread)
    End If
End Sub

Public Function ChatLineSN(TheChatLine As String) As String
    If InStr(TheChatLine, ":") = 0 Then
        ChatLineSN = ""
        Exit Function
    End If
    ChatLineSN = left(TheChatLine, InStr(TheChatLine, ":") - 1)
End Function

Public Function ChatLineMsg(TheChatLine As String) As String
    If InStr(TheChatLine, Chr(9)) = 0 Then
        ChatLineMsg = ""
        Exit Function
    End If
    ChatLineMsg = right(TheChatLine, Len(TheChatLine) - InStr(TheChatLine, Chr(9)))
End Function

Public Sub Scroll(ScrollString As String)
    Dim CurLine As String, count As Long, ScrollIt As Long
    Dim sProgress As Long
    If findroom& = 0 Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    count& = linecount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To count&
        CurLine$ = linefromstring(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = left(CurLine$, 92)
            End If
            Call chatsend(CurLine$)
            pause 0.7
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            pause 0.55
        End If
    Next ScrollIt&
End Sub

Public Sub WaitForOKOrRoom(room As String)
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    room$ = (ReplaceString(room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(findroom&)
        RoomTitle$ = (ReplaceString(room$, " ", ""))
        FullWindow& = findwindow("#32770", "America Online")
        FullButton& = findwindowex(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or room$ = RoomTitle$
    DoEvents
    If FullWindow& <> 0& Then
        Do
            DoEvents
            Call sendmessage(FullButton&, wm_keydown, vk_space, 0&)
            Call sendmessage(FullButton&, wm_keyup, vk_space, 0&)
            Call sendmessage(FullButton&, wm_keydown, vk_space, 0&)
            Call sendmessage(FullButton&, wm_keyup, vk_space, 0&)
            FullWindow& = findwindow("#32770", "America Online")
            FullButton& = findwindowex(FullWindow&, 0&, "Button", "OK")
        Loop Until FullWindow& = 0& And FullButton& = 0&
    End If
    DoEvents
End Sub

Public Sub MemberRoom(room As String)
    Call keyword("aol://2719:61-2-" & room$)
End Sub

Public Sub PublicRoom(room As String)
    Call keyword("aol://2719:21-2-" & room$)
End Sub

Public Sub PrivateRoom(room As String)
    Call keyword("aol://2719:2-2-" & room$)
End Sub

Public Sub instantmessage(Person As String, message As String)
    Dim aol As Long, mdi As Long, im As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://9293:" & Person$)
    Do
        DoEvents
        im& = findwindowex(mdi&, 0&, "AOL Child", "Send Instant Message")
        Rich& = findwindowex(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = findwindowex(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = findwindowex(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call sendmessagebystring(Rich&, wm_settext, 0&, message$)
    Call sendmessage(SendButton&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(SendButton&, wm_lbuttonup, 0&, 0&)
    Do
        DoEvents
        OK& = findwindow("#32770", "America Online")
        im& = findwindowex(mdi&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or im& = 0&
    If OK& <> 0& Then
        Button& = findwindowex(OK&, 0&, "Button", vbNullString)
        Call postmessage(Button&, wm_keydown, vk_space, 0&)
        Call postmessage(Button&, wm_keyup, vk_space, 0&)
        Call postmessage(im&, wm_close, 0&, 0&)
    End If
End Sub

Public Function CheckIMs(Person As String) As Boolean
    Dim aol As Long, mdi As Long, im As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://9293:" & Person$)
    Do
        DoEvents
        im& = findwindowex(mdi&, 0&, "AOL Child", "Send Instant Message")
        Rich& = findwindowex(im&, 0&, "RICHCNTL", vbNullString)
        Available1& = findwindowex(im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = findwindowex(im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = findwindowex(im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = findwindowex(im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = findwindowex(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = findwindowex(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = findwindowex(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = findwindowex(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = findwindowex(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = findwindowex(im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call sendmessage(Available&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(Available&, wm_lbuttonup, 0&, 0&)
    Do
        DoEvents
        oWindow& = findwindow("#32770", "America Online")
        oButton& = findwindowex(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = findwindowex(oWindow&, 0&, "Static", vbNullString)
        oStatic& = findwindowex(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = gettext(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
    Else
        CheckIMs = False
    End If
    Call sendmessage(oButton&, wm_keydown, vk_space, 0&)
    Call sendmessage(oButton&, wm_keyup, vk_space, 0&)
    Call postmessage(im&, wm_close, 0&, 0&)
End Function

Public Sub imignore(Person As String)
    Call instantmessage("$IM_OFF, " & Person$, "=)")
End Sub

Public Sub IMUnIgnore(Person As String)
    Call instantmessage("$IM_ON, " & Person$, "=)")
End Sub

Public Sub imsoff()
    Call instantmessage("$IM_OFF", "=)")
End Sub

Public Sub imson()
    Call instantmessage("$IM_ON", "=)")
End Sub

Public Function IMSender() As String
    Dim im As Long, Caption As String
    Caption$ = GetCaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function

Public Function IMText() As String
    Dim Rich As Long
    Rich& = findwindowex(FindIM&, 0&, "RICHCNTL", vbNullString)
    IMText$ = gettext(Rich&)
End Function

Public Function IMLastMsg() As String
    Dim Rich As Long, MsgString As String, Spot As Long
    Dim NewSpot As Long
    Rich& = findwindowex(FindIM&, 0&, "RICHCNTL", vbNullString)
    MsgString$ = gettext(Rich&)
    NewSpot& = InStr(MsgString$, Chr(9))
    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
    MsgString$ = right(MsgString$, Len(MsgString$) - Spot& - 1)
    IMLastMsg$ = left(MsgString$, Len(MsgString$) - 1)
End Function

Public Sub IMRespond(Msg As String)
    Dim im As Long, Rich As Long, Icon As Long
    im& = FindIM&
    If im& = 0& Then Exit Sub
    Rich& = findwindowex(im&, 0&, "RICHCNTL", vbNullString)
    Rich& = findwindowex(im&, Rich&, "RICHCNTL", vbNullString)
    Icon& = findwindowex(im&, 0&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = findwindowex(im&, Icon&, "_AOL_Icon", vbNullString)
    Call sendmessagebystring(Rich&, wm_settext, 0&, Msg$)
    DoEvents
    Call sendmessage(Icon&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(Icon&, wm_lbuttonup, 0&, 0&)
End Sub

Public Sub keyword(kw As String)
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = findwindowex(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = findwindowex(Combo&, 0&, "Edit", vbNullString)
    Call sendmessagebystring(EditWin&, wm_settext, 0&, kw$)
    Call sendmessagelong(EditWin&, wm_char, vk_space, 0&)
    Call sendmessagelong(EditWin&, wm_char, vk_return, 0&)
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
    NewText$ = left(TheText$, CharNum&)
    NewText$ = right(NewText$, 1)
    LineChar$ = NewText$
End Function

Public Function linecount(MyString As String) As Long
    Dim Spot As Long, count As Long
    If Len(MyString$) < 1 Then
        linecount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        linecount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                linecount& = linecount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    linecount& = linecount& + 1
End Function

Public Function linefromstring(MyString As String, Line As Long) As String
    Dim theline As String, count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    count& = linecount(MyString$)
    If Line& > count& Then
        Exit Function
    End If
    If Line& = 1 And count& = 1 Then
        linefromstring$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        linefromstring$ = theline$
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
        linefromstring$ = theline$
    End If
End Function

Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr((MyString$), (ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
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
            NewSpot& = InStr(Spot&, (MyString$), (ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function

Public Function ReverseString(MyString As String) As String
    Dim TempString As String, StringLength As Long
    Dim count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While count& <= StringLength&
        count& = count& + 1
        NextChr$ = Mid$(TempString$, count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
End Function

Public Function SwitchStrings(MyString As String, String1 As String, String2 As String) As String
    Dim TempString As String, Spot1 As Long, Spot2 As Long
    Dim Spot As Long, ToFind As String, ReplaceWith As String
    Dim NewSpot As Long, LeftString As String, RightString As String
    Dim NewString As String
    If Len(String2) > Len(String1) Then
        TempString$ = String1$
        String1$ = String2$
        String2$ = TempString$
    End If
    Spot1& = InStr(MyString$, String1$)
    Spot2& = InStr(MyString$, String2$)
    If Spot1& = 0& And Spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    If Spot1& < Spot2& Or Spot2& = 0 Or Len(String1$) = Len(String2$) Then
        If Spot1& > 0 Then
            Spot& = Spot1&
            ToFind$ = String1$
            ReplaceWith$ = String2$
        End If
    End If
    If Spot2& < Spot1& Or Spot1& = 0& Then
        If Spot2& > 0& Then
            Spot& = Spot2&
            ToFind$ = String2$
            ReplaceWith$ = String1$
        End If
    End If
    If Spot1& = 0& And Spot2& = 0& Then
        SwitchStrings$ = MyString$
        Exit Function
    End If
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
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
            Spot1& = InStr(Spot&, MyString$, String1$)
            Spot2& = InStr(Spot&, MyString$, String2$)
        End If
        If Spot1& = 0& And Spot2& = 0& Then
            SwitchStrings$ = MyString$
            Exit Function
        End If
        If Spot1& < Spot2& Or Spot2& = 0& Or Len(String1$) = Len(String2$) Then
            If Spot1& > 0& Then
                Spot& = Spot1&
                ToFind$ = String1$
                ReplaceWith$ = String2$
            End If
        End If
        If Spot2& < Spot1& Or Spot1& = 0& Then
            If Spot2& > 0& Then
                Spot& = Spot2&
                ToFind$ = String2$
                ReplaceWith$ = String1$
            End If
        End If
        If Spot1& = 0& And Spot2& = 0& Then
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
    NumOfLines& = linecount(MyString$)
    OrgCount& = NumOfLines&
    For SpaceIt& = 1 To NumOfLines&
        DoEvents
        CurLine$ = linefromstring(MyString$, SpaceIt&)
        NewString$ = NewString$ & " " & CurLine$ & NewLine$
    Next SpaceIt&
    MyString$ = left(NewString$, OrgLen& + OrgCount& - 1)
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
        NumOfLines& = linecount(MyText$)
        For ReverseIt& = 1 To NumOfLines
            CurLine$ = linefromstring(MyText$, ReverseIt&)
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
        NewString$ = left(NewString$, CheckLen& - 4)
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
        NumOfLines& = linecount(MyText$)
        For StretchIt& = 1 To NumOfLines&
            CurLine$ = linefromstring(MyText, StretchIt&)
            CurLine$ = DoubleText(CurLine$)
            NewString$ = NewString$ & CurLine$ & NewLine$
        Next StretchIt&
        CheckLen& = Len(NewString$)
        NewString$ = left(NewString$, CheckLen& - 4)
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
    NumOfLines& = linecount(MyString)
    OrgCount& = NumOfLines&
    For SpaceIt& = 1 To NumOfLines&
        CurLine$ = linefromstring(MyString$, SpaceIt&)
        If Len(CurLine$) < 1 Then
            NewString$ = NewString$ & CurLine$ & NewLine$
        Else
            NewString$ = NewString$ & right(CurLine$, Len(CurLine$) - 1) & NewLine$
        End If
    Next SpaceIt&
    MyString$ = left(NewString$, Len(NewString$) - 4)
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
        NumOfLines& = linecount(MyText$)
        MyLine& = NumOfLines& - 1
        For FlipIt& = 1 To NumOfLines&
            DoEvents
            CurLine$ = linefromstring(MyText$, MyLine&)
            NewString$ = NewString$ & CurLine$ & NewLine$
            MyLine& = MyLine& - 1
        Next FlipIt&
        NewString$ = left(NewString$, Len(NewString$) - 4)
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

Public Function fileexists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        fileexists = False
        Exit Function
    End If
    If Len(dir$(sFileName$)) Then
        fileexists = True
    Else
        fileexists = False
    End If
End Function

Sub loadtext(txtLoad As TextBox, path As String)
    Dim TextString As String
    On Error Resume Next
    Open path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub

Sub savetext(txtSave As TextBox, path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub Loadlistbox(Directory As String, TheList As ListBox)
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

Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        aString$ = left(MyString$, InStr(MyString$, "*") - 1)
        bString$ = right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
    Wend
    Close #1
End Sub

Public Sub SaveListBox(Directory As String, TheList As ListBox)
    Dim savelist As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For savelist& = 0 To TheList.ListCount - 1
        Print #1, TheList.list(savelist&)
    Next savelist&
    Close #1
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.list(SaveLists&) & "*" & ListB.list(SaveLists)
    Next SaveLists&
    Close #1
End Sub

Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.list(SaveCombo&)
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

Public Function FileGetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function

Public Sub FileSetNormal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub

Public Function getfromini(section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = (Key$)
   getfromini$ = left(strBuffer, getprivateprofilestring(section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub writetoini(section As String, Key As String, KeyValue As String, Directory As String)
    Call writeprivateprofilestring(section$, UCase$(Key$), KeyValue$, Directory$)
End Sub

Public Function CheckIfMaster() As Boolean
    Dim aol As Long, mdi As Long, pWindow As Long
    Dim pbutton As Long, modal As Long, mstatic As Long
    Dim mstring As String
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDICLIENT", vbNullString)
    Call keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do
        DoEvents
        pWindow& = findwindowex(mdi&, 0&, "AOL Child", "Parental Controls")
        pbutton& = findwindowex(pWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pWindow& <> 0& And pbutton& <> 0&
    pause 0.3
    Do
        DoEvents
        Call postmessage(pbutton&, wm_lbuttondown, 0&, 0&)
        Call postmessage(pbutton&, wm_lbuttonup, 0&, 0&)
        pause 0.8
        modal& = findwindow("_AOL_Modal", vbNullString)
        mstatic& = findwindowex(modal&, 0&, "_AOL_Static", vbNullString)
        mstring$ = gettext(mstatic&)
    Loop Until modal& <> 0 And mstatic& <> 0& And mstring$ <> ""
    mstring$ = ReplaceString(mstring$, Chr(10), "")
    mstring$ = ReplaceString(mstring$, Chr(13), "")
    If mstring$ = "Set Parental Controls" Then
        CheckIfMaster = True
    Else
        CheckIfMaster = False
    End If
    Call postmessage(modal&, wm_close, 0&, 0&)
    DoEvents
    Call postmessage(pWindow&, wm_close, 0&, 0&)
End Function

Public Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = getwindowtextlength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call getwindowtext(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function

Public Function GetListText(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = sendmessage(WindowHandle&, lb_gettextlen, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call sendmessagebystring(WindowHandle&, lb_gettext, TextLength& + 1, Buffer$)
    GetListText$ = Buffer$
End Function

Public Function gettext(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = sendmessage(WindowHandle&, wm_gettextlength, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call sendmessagebystring(WindowHandle&, wm_gettext, TextLength& + 1, Buffer$)
    gettext$ = Buffer$
End Function

Public Sub Button(mButton As Long)
    Call sendmessage(mButton&, wm_keydown, vk_space, 0&)
    Call sendmessage(mButton&, wm_keyup, vk_space, 0&)
End Sub

Public Sub Icon(aIcon As Long)
    Call sendmessage(aIcon&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(aIcon&, wm_lbuttonup, 0&, 0&)
End Sub

Public Sub CloseWindow(Window As Long)
    Call postmessage(Window&, wm_close, 0&, 0&)
End Sub

Public Function ProfileGet(screenname As String) As String
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim mdi As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim nowindow As Long, OKButton As Long, CurPos As pointapi
    Dim WinVis As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    tool& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = findwindowex(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = findwindowex(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call getcursorpos(CurPos)
    Call setcursorpos(Screen.Width, Screen.Height)
    Call postmessage(ToolIcon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ToolIcon&, wm_lbuttonup, 0&, 0&)
    Do
        sMod& = findwindow("#32768", vbNullString)
        WinVis& = iswindowvisible(sMod&)
    Loop Until WinVis& = 1
    Call postmessage(sMod&, wm_keydown, vk_up, 0&)
    Call postmessage(sMod&, wm_keyup, vk_up, 0&)
    Call postmessage(sMod&, wm_keydown, vk_up, 0&)
    Call postmessage(sMod&, wm_keyup, vk_up, 0&)
    Call postmessage(sMod&, wm_keydown, vk_return, 0&)
    Call postmessage(sMod&, wm_keyup, vk_return, 0&)
    Call setcursorpos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = findwindowex(mdi&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = findwindowex(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = findwindowex(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call sendmessagebystring(pgEdit&, wm_settext, 0&, screenname$)
    Call sendmessage(pgButton&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(pgButton&, wm_lbuttonup, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = findwindowex(mdi&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = findwindowex(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = gettext(pTextWindow&)
        nowindow& = findwindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or nowindow& <> 0&
    DoEvents
    If nowindow& <> 0& Then
        OKButton& = findwindowex(nowindow&, 0&, "Button", "OK")
        Call sendmessage(OKButton&, wm_keydown, vk_space, 0&)
        Call sendmessage(OKButton&, wm_keyup, vk_space, 0&)
        Call postmessage(pgWindow&, wm_close, 0&, 0&)
        ProfileGet$ = "< No Profile >"
    Else
        Call postmessage(pWindow&, wm_close, 0&, 0&)
        Call postmessage(pgWindow&, wm_close, 0&, 0&)
        ProfileGet$ = pString$
    End If
End Function

Public Function getuser() As String
    Dim aol As Long, mdi As Long, welcome As Long
    Dim Child As Long, UserString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(Child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        getuser$ = UserString$
        Exit Function
    Else
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(Child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                getuser$ = UserString$
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    getuser$ = ""
End Function

Public Sub pause(Duration As Long)
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

Public Sub SetText(Window As Long, text As String)
    Call sendmessagebystring(Window&, wm_settext, 0&, text$)
End Sub

Public Function ListToMailString(TheList As ListBox) As String
    Dim DoList As Long, MailString As String
    If TheList.list(0) = "" Then Exit Function
    For DoList& = 0 To TheList.ListCount - 1
        MailString$ = MailString$ & "(" & TheList.list(DoList&) & "), "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToMailString$ = MailString$
End Function

Public Sub formontop(FormName As Form)
    Call setwindowpos(FormName.hwnd, hwnd_topmost, 0&, 0&, 0&, 0&, flags)
End Sub

Public Sub formnotontop(FormName As Form)
    Call setwindowpos(FormName.hwnd, hwnd_notopmost, 0&, 0&, 0&, 0&, flags)
End Sub

Public Sub formdrag(TheForm As Form)
    Call releasecapture
    Call sendmessage(TheForm.hwnd, wm_syscommand, wm_move, 0)
End Sub

Public Sub FormExitDown(TheForm As Form)
    Do
        DoEvents
        TheForm.top = Trim(Str(Int(TheForm.top) + 300))
    Loop Until TheForm.top > 10000
End Sub

Public Sub FormExitLeft(TheForm As Form)
    Do
        DoEvents
        TheForm.left = Trim(Str(Int(TheForm.left) - 300))
    Loop Until TheForm.left < -TheForm.Width
End Sub

Public Sub FormExitRight(TheForm As Form)
    Do
        DoEvents
        TheForm.left = Trim(Str(Int(TheForm.left) + 300))
    Loop Until TheForm.left > Screen.Width
End Sub

Public Sub FormExitUp(TheForm As Form)
    Do
        DoEvents
        TheForm.top = Trim(Str(Int(TheForm.top) - 300))
    Loop Until TheForm.top < -TheForm.Width
End Sub

Public Sub WindowHide(hwnd As Long)
    Call showwindow(hwnd&, sw_hide)
End Sub

Public Sub WindowShow(hwnd As Long)
    Call showwindow(hwnd&, sw_show)
End Sub

Public Sub runmenu(topmenu As Long, submenu As Long)
    Dim aol As Long, aMenu As Long, smenu As Long, mnID As Long
    Dim mVal As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    aMenu& = getmenu(aol&)
    smenu& = getsubmenu(aMenu&, topmenu&)
    mnID& = getmenuitemid(smenu&, submenu&)
    Call sendmessagelong(aol&, wm_command, mnID&, 0&)
End Sub

Public Sub runmenubystring(SearchString As String)
    Dim aol As Long, aMenu As Long, mcount As Long
    Dim LookFor As Long, smenu As Long, scount As Long
    Dim LookSub As Long, sID As Long, sstring As String
    aol& = findwindow("AOL Frame25", vbNullString)
    aMenu& = getmenu(aol&)
    mcount& = getmenuitemcount(aMenu&)
    For LookFor& = 0& To mcount& - 1
        smenu& = getsubmenu(aMenu&, LookFor&)
        scount& = getmenuitemcount(smenu&)
        For LookSub& = 0 To scount& - 1
            sID& = getmenuitemid(smenu&, LookSub&)
            sstring$ = String$(100, " ")
            Call getmenustring(smenu&, sID&, sstring$, 100&, 1&)
            If InStr((sstring$), (SearchString$)) Then
                Call sendmessagelong(aol&, wm_command, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Public Function clickyes()
Dim AOLIcon As Long
Dim aolmodal As Long
aolmodal& = findwindow("_AOL_Modal", vbNullString)
AOLIcon& = findwindowex(aolmodal&, 0&, "_AOL_Icon", vbNullString)
Call sendmessage(AOLIcon&, wm_lbuttondown, 0&, 0&)
Call sendmessage(AOLIcon&, wm_lbuttonup, 0&, 0&)
End Function

Public Function clickyes245mintimer()
Dim AOLIcon As Long
Dim aolmodal As Long
aolmodal& = findwindow("_AOL_Modal", vbNullString)
AOLIcon& = findwindowex(aolmodal&, 0&, "_AOL_Icon", vbNullString)
Call sendmessage(AOLIcon&, wm_lbuttondown, 0&, 0&)
Call sendmessage(AOLIcon&, wm_lbuttonup, 0&, 0&)
End Function
Function trimspaces(text As String) As String
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
Public Function AimFindAimWindow() As Long
    AimFindAimWindow& = findwindow("_Oscar_BuddyListWin", vbNullString)
End Function
Public Function AimUserSn() As String
    AimUserSn$ = Mid(GetCaption(AimFindAimWindow&), 1, InStr(GetCaption(AimFindAimWindow&), "'") - 1)
End Function
Public Function AimFindRoom() As Long
    AimFindRoom& = findwindow("AIM_ChatWnd", vbNullString)
End Function
Public Sub AimRoomSend(SendString As String, Optional ClearBefore As Boolean = True)
    Dim WndAte1 As Long, WndAte2 As Long, OscarButton1 As Long
    Dim OscarButton2 As Long, OscarButton3 As Long, OscarButton4 As Long
    WndAte1& = findwindowex(AimFindRoom&, 0&, "WndAte32Class", vbNullString)
    WndAte2& = findwindowex(AimFindRoom&, WndAte1&, "WndAte32Class", vbNullString)
    OscarButton1& = findwindowex(AimFindRoom&, 0&, "_Oscar_IconBtn", vbNullString)
    OscarButton2& = findwindowex(AimFindRoom&, OscarButton1&, "_Oscar_IconBtn", vbNullString)
    OscarButton3& = findwindowex(AimFindRoom&, OscarButton2&, "_Oscar_IconBtn", vbNullString)
    OscarButton4& = findwindowex(AimFindRoom&, OscarButton3&, "_Oscar_IconBtn", vbNullString)
    If ClearBefore = True Then Call sendmessagebystring(WndAte2&, wm_settext, 0&, "")
    Call sendmessagebystring(WndAte2&, wm_settext, 0&, SendString$)
    Call sendmessage(OscarButton4&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(OscarButton4&, wm_lbuttonup, 0&, 0&)
End Sub
Public Function AimOnline() As Boolean
    If AimFindAimWindow& <> 0& Then
        AimOnline = True
       ElseIf AimFindAimWindow& = 0& Then
        AimOnline = False
    End If
End Function

Public Sub AimRoomClear()
    Dim WndAte As Long
    WndAte& = findwindowex(AimFindRoom&, 0&, "WndAte32Class", vbNullString)
    Call sendmessagebystring(WndAte&, wm_settext, 0&, "")
End Sub
Public Function AimChatSend(txt As String)
    Dim ChatWin As Long, ChatBox As Long, WinAte As Long
    Dim chatview As Long, SendBut As Long
    ChatWin& = findwindow("AIM_ChatWnd", vbNullString)
    WinAte& = findwindowex(ChatWin&, 0&, "WndAte32Class", vbNullString)
    chatview& = findwindowex(WinAte&, 0&, "Ate32Class", vbNullString)
    ChatBox& = findwindowex(ChatWin&, WinAte&, "WndAte32Class", vbNullString)
    SendBut& = findwindowex(ChatWin&, 0, "_Oscar_IconBtn", vbNullString)
    SendBut& = findwindowex(ChatWin&, SendBut&, "_Oscar_IconBtn", vbNullString)
    SendBut& = findwindowex(ChatWin&, SendBut&, "_Oscar_IconBtn", vbNullString)
    SendBut& = findwindowex(ChatWin&, SendBut&, "_Oscar_IconBtn", vbNullString)
    Call sendmessagebystring(ChatBox&, wm_settext, 0&, txt)
    Do
        Call postmessage(SendBut&, wm_lbuttondown, 0&, 0&)
        Call postmessage(SendBut&, wm_lbuttonup, 0&, 0&)
        DoEvents
    Loop Until gettext(ChatBox&) = ""
End Function
Public Sub AnimateLabel(TheLabel As Label)
    Dim strlen As String, Ani As Long, Fsize As Integer
    Dim GreetTo As String
    GreetTo = TheLabel.Caption
    For Ani = 6 To 26
        TheLabel.FontSize = Ani
        Call pause(0.001)
    Next Ani
    strlen = Len(TheLabel.Caption)
    For Ani = 1 To strlen
        TheLabel.Caption = left(GreetTo, strlen - Ani)
        Call pause(0.001)
    Next Ani
End Sub
Public Sub FormExitDownSlow(TheForm As Form)
    Do
        DoEvents
        TheForm.top = Trim(Str(Int(TheForm.top) + 30))
    Loop Until TheForm.top > 12000
End Sub
Sub LoadText1(txtLoad As TextBox, path As String)
    Dim TextString As String
    On Error Resume Next
    Open path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub
Public Sub CenterForm(frmform As Form)
   With frmform
      .left = (Screen.Width - .Width) / 2
      .top = (Screen.Height - .Height) / 2
   End With
End Sub
Public Sub FadeColors(frm As Object, topcolor As Long, BottomColor As Long)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim i&, J&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim TopColorRed%, TopColorGreen%, TopColorBlue%
Dim BottomColorRed%, BottomColorGreen%, BottomColorBlue%
Dim ColorDifRed, ColorDifGreen, ColorDifBlue
SaveScale = frm.ScaleMode: SaveStyle = frm.DrawStyle
SaveRedraw = frm.AutoRedraw: frm.ScaleMode = 3
TopColorRed = topcolor And 255
  TopColorGreen = (topcolor And 65280) / 256
    TopColorBlue = (topcolor And 16711680) / 65536
BottomColorRed = BottomColor And 255
  BottomColorGreen = (BottomColor And 65280) / 256
    BottomColorBlue = (BottomColor And 16711680) / 65536
      aRed = TopColorRed
      aGreen = TopColorGreen
      aBlue = TopColorBlue
      pixels = frm.ScaleWidth
    If pixels <= 0 Then Exit Sub
        ColorDifRed = (BottomColorRed - TopColorRed)
        ColorDifGreen = (BottomColorGreen - TopColorGreen)
        ColorDifBlue = (BottomColorBlue - TopColorBlue)
          RedDelta = ColorDifRed / pixels
          GreenDelta = ColorDifGreen / pixels
          BlueDelta = ColorDifBlue / pixels
        frm.DrawStyle = 5
        frm.AutoRedraw = True
For Y = 0 To pixels + 1
        aRed = aRed + RedDelta
            If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
            If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
            If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
            If ThisColor > -1 Then
                frm.Line (Y - 2, -2)-(Y - 2, frm.Height + 2), ThisColor, BF
            End If
    Next Y
frm.ScaleMode = SaveScale
frm.DrawStyle = SaveStyle
frm.AutoRedraw = SaveRedraw
End Sub
Public Sub FormCircle(frm As Form, Size As Long)
    Dim E As Long
    
    'makes for do a circle.. [as seen in pH]
    'make size between 1 and 100 about..
    'example:
    '
    'Call FormCircle(Me, 20)
    
    For E& = Size& - 1 To 0 Step -1
        frm.left = frm.left - E&
        frm.top = frm.top + (Size& - E&)
    Next E&
    
    For E& = Size& - 1 To 0 Step -1
        frm.left = frm.left + (Size& - E&)
        frm.top = frm.top + E&
    Next E&
    
    For E& = Size& - 1 To 0 Step -1
        frm.left = frm.left + E&
        frm.top = frm.top - (Size& - E&)
    Next E&
    
    For E& = Size& - 1 To 0 Step -1
        frm.left = frm.left - (Size& - E&)
        frm.top = frm.top - E&
    Next E&
End Sub
Public Sub FadeBy2(frm As Object, topcolor As Long, BottomColor As Long)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim i&, J&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim TopColorRed%, TopColorGreen%, TopColorBlue%
Dim BottomColorRed%, BottomColorGreen%, BottomColorBlue%
Dim ColorDifRed, ColorDifGreen, ColorDifBlue
SaveScale = frm.ScaleMode: SaveStyle = frm.DrawStyle
SaveRedraw = frm.AutoRedraw: frm.ScaleMode = 3
TopColorRed = topcolor And 255
  TopColorGreen = (topcolor And 65280) / 256
    TopColorBlue = (topcolor And 16711680) / 65536
BottomColorRed = BottomColor And 255
  BottomColorGreen = (BottomColor And 65280) / 256
    BottomColorBlue = (BottomColor And 16711680) / 65536
      aRed = TopColorRed
      aGreen = TopColorGreen
      aBlue = TopColorBlue
      pixels = frm.ScaleWidth
    If pixels <= 0 Then Exit Sub
        ColorDifRed = (BottomColorRed - TopColorRed)
        ColorDifGreen = (BottomColorGreen - TopColorGreen)
        ColorDifBlue = (BottomColorBlue - TopColorBlue)
          RedDelta = ColorDifRed / pixels
          GreenDelta = ColorDifGreen / pixels
          BlueDelta = ColorDifBlue / pixels
        frm.DrawStyle = 5
        frm.AutoRedraw = True
For Y = 0 To pixels + 1
        aRed = aRed + RedDelta
            If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
            If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
            If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
            If ThisColor > -1 Then
                frm.Line (Y - 2, -2)-(Y - 2, frm.Height + 2), ThisColor, BF
            End If
    Next Y
frm.ScaleMode = SaveScale
frm.DrawStyle = SaveStyle
frm.AutoRedraw = SaveRedraw
End Sub
Public Function dupekill(list As Control) As Long
    Dim amount As Long, Y As Long, X As Long
    
    For Y = 0 To list.ListCount
        For X = Y + 1 To list.ListCount '- Y + 1
            If list.list(X) = list.list(Y) Then
                list.RemoveItem (X)
                amount = amount + 1
                X = X - 1
            End If
        Next X
    Next Y
    
    dupekill = amount
End Function
Public Sub generate3letters(howmanysns As Integer, list As ListBox)
    Dim alphanumericstring As String, alphastring As String
    Dim strletter As String, sn As String, randomtime As String
    Dim rndx As Integer, rndy As Integer, makinsns As Integer, i As Long
    
    alphanumericstring = "1234567890abcdefghijklmnopqrstuvwxyz"
    alphastring = "abcdefghijklmnopqrstuvwxyz"
    
    
    Do While makinsns <> howmanysns
        DoEvents
        
        
        sn = ""
        
        Call Randomize
        
        rndx = Int(Rnd * 26) + 1
        strletter = Mid(alphastring, rndx, 1)
        sn = sn + strletter
        
        rndy = Int(Rnd * 36) + 1
        strletter = Mid(alphanumericstring, rndy, 1)
        sn = sn + strletter
        
        rndy = Int(Rnd * 36) + 1
        strletter = Mid(alphanumericstring, rndy, 1)
        sn = sn + strletter
        
        list.AddItem sn
        makinsns = makinsns + 1
    Loop
End Sub

Public Sub findachat()
    Dim aol As Long, mdi As Long, facWin As Long
    Dim fwin As Long, flist As Long, fcount As Long
    Dim pcwin As Long, pcicon As Long
    
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    
    If aol& = 0& Or getuser = "" Then Exit Sub
    
    facWin& = findwindowex(mdi&, 0&, "AOL Child", "Find a Chat")
    If facWin& <> 0& Then Exit Sub
    
    If aolversion = "4" Or aolversion = "5" Then
        Call runtbmenu(10&, 3&)
    Else
        Call keyword25("pc")
        
        Do
            DoEvents
            pcwin& = findwindowex(mdi&, 0&, "AOL Child", " Welcome to People Connection")
            pcicon& = findwindowex(pcwin&, 0&, "_AOL_Icon", vbNullString)
            pcicon& = findwindowex(pcwin&, pcicon&, "_AOL_Icon", vbNullString)
            pcicon& = findwindowex(pcwin&, pcicon&, "_AOL_Icon", vbNullString)
            pcicon& = findwindowex(pcwin&, pcicon&, "_AOL_Icon", vbNullString)
            pcicon& = findwindowex(pcwin&, pcicon&, "_AOL_Icon", vbNullString)
            pcicon& = findwindowex(pcwin&, pcicon&, "_AOL_Icon", vbNullString) '
            pcicon& = findwindowex(pcwin&, pcicon&, "_AOL_Icon", vbNullString)
        Loop Until pcwin& <> 0& And pcicon& <> 0&
        
        Call runmenubystring("incoming text")
        
        Call sendmessage(pcicon&, wm_lbuttondown, 0&, 0&)
        Call sendmessage(pcicon&, wm_lbuttonup, 0&, 0&)
    End If
    
    Do
        DoEvents
        fwin& = findwindowex(mdi&, 0&, "AOL Child", "Find a Chat")
        flist& = findwindowex(fwin&, 0&, "_AOL_Listbox", vbNullString)
        flist& = findwindowex(fwin&, flist&, "_AOL_Listbox", vbNullString)
        fcount& = sendmessage(flist&, lb_getcount, 0&, 0&)
    Loop Until fwin& <> 0& And flist& <> 0& And fcount& <> 0&
    
    pause (1)
End Sub
Public Sub runtbmenu(iconnum As Long, mnunumber As Long)
    Dim aol As Long, mdi As Long, tb As Long, tbar As Long
    Dim ticon As Long, ilong As Long, mlong As Long
    Dim tmenu As Long, wvisible As Long, cposition As pointapi
    
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    tb& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    tbar& = findwindowex(tb&, 0&, "_AOL_Toolbar", vbNullString)
    
    ticon& = findwindowex(tbar&, 0&, "_AOL_Icon", vbNullString)
    For ilong& = 1 To iconnum - 1
        ticon& = findwindowex(tbar&, ticon&, "_AOL_Icon", vbNullString)
    Next ilong&
    
    Call getcursorpos(cposition)
    Call setcursorpos(Screen.Width, Screen.Height)
    Call postmessage(ticon&, wm_lbuttondown, 0&, 0&)
    Call postmessage(ticon&, wm_lbuttonup, 0&, 0&)
    pause (0.09)
    
    Do
        tmenu& = findwindow("#32768", vbNullString)
        wvisible& = iswindowvisible(tmenu&)
    Loop Until wvisible& = 1
    
    For mlong& = 1 To mnunumber&
        Call postmessage(tmenu&, wm_keydown, vk_down, 0&)
        Call postmessage(tmenu&, wm_keyup, vk_down, 0&)
    Next mlong&
    
    Call postmessage(tmenu&, wm_keydown, vk_return, 0&)
    Call postmessage(tmenu&, wm_keyup, vk_return, 0&)
    Call setcursorpos(cposition.X, cposition.Y)
End Sub
Public Sub keyword25(kw As String)
    Dim aol As Long, mdi As Long, tbar As Long, ticon As Long
    Dim tlong As Long, kwin As Long, kedit As Long
    
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    
    tbar& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    ticon& = findwindowex(tbar&, 0&, "_AOL_Icon", vbNullString)
    If aolversion = "2.5" Then
        For tlong& = 1 To 12
            ticon& = findwindowex(tbar&, ticon&, "_AOL_Icon", vbNullString)
        Next tlong
    ElseIf aolversion = "3" Then
        For tlong& = 1 To 17
            ticon& = findwindowex(tbar&, ticon&, "_AOL_Icon", vbNullString)
        Next tlong
    End If
    
    Call sendmessage(ticon&, wm_lbuttondown, 0&, 0&)
    Call sendmessage(ticon&, wm_lbuttonup, 0&, 0&)
    
    Do
        DoEvents
        kwin& = findwindowex(mdi&, 0&, "AOL Child", "Keyword")
        kedit& = findwindowex(kwin&, 0&, "_AOL_Edit", vbNullString)
    Loop Until kwin& <> 0& And kedit& <> 0&
    
    Call sendmessagebystring(kedit&, wm_settext, 0&, kw$)
    Call sendmessagelong(kedit&, wm_char, enter_key, 0&)
End Sub

Public Function aolversion() As String
    Dim aol As Long, gmenu As Long, mnu As Long
    Dim smenu As Long, sitem As Long, mstring As String
    Dim fstring As Long, tb As Long, tbar As Long
    Dim tcombo As Long, tedit As Long
    
    aol& = findwindow("AOL Frame25", vbNullString)
    tb& = findwindowex(aol&, 0&, "AOL Toolbar", vbNullString)
    tbar& = findwindowex(tb&, 0&, "_AOL_Toolbar", vbNullString)
    tcombo& = findwindowex(tbar&, 0&, "_AOL_Combobox", vbNullString)
    tedit& = findwindowex(tcombo&, 0&, "Edit", vbNullString)
    
    If aol& = 0 Then
        aolversion$ = "0"
        Exit Function
    End If
    
    If tedit& <> 0& And tcombo& <> 0& Then
        gmenu& = getmenu(aol&)
        
        smenu& = getsubmenu(gmenu&, 4&)
        sitem& = getmenuitemid(smenu&, 9&)
        mstring$ = String$(100, " ")
        
        fstring& = getmenustring(smenu&, sitem&, mstring$, 100, 1)
        
        If InStr(1, LCase(mstring$), LCase("&What's New in AOL 5.0")) <> 0& Then
            aolversion = "5"
        Else
            aolversion = "4"
        End If
    Else
        aol& = findwindow("AOL Frame25", vbNullString)
        gmenu& = getmenu(aol&)
        
        mnu& = getmenuitemcount(getmenu(aol&))
        If mnu& = 8 Then
            smenu& = getsubmenu(gmenu&, 1)
            sitem& = getmenuitemid(smenu&, 8)
            mstring$ = String$(100, " ")
        Else
            smenu& = getsubmenu(gmenu&, 0)
            sitem& = getmenuitemid(smenu&, 8)
            mstring$ = String$(100, " ")
        End If
        
        fstring& = getmenustring(smenu&, sitem&, mstring$, 100, 1)
        
        If InStr(1, LCase(mstring$), LCase("&LOGGING...")) <> 0& Then
            aolversion = "2.5"
        Else
            aolversion = "3"
        End If
    End If
End Function
Public Function getchatname() As String
    Let getchatname$ = GetCaption(findroom&)
End Function
Public Sub HideWelcome()
Call showwindow(findwelcome, sw_hide)
End Sub
Public Sub ShowWelcome()
Call showwindow(findwelcome, sw_show)
End Sub
Function findwelcome() As Long
    Dim aol As Long, mdi As Long, welcome As Long
    Dim Child As Long, UserString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    mdi& = findwindowex(aol&, 0&, "MDIClient", vbNullString)
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(Child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        findwelcome& = Child&
        Exit Function
    Else
    Child& = findwindowex(mdi&, 0&, "AOL Child", vbNullString)
        Do
            Child& = findwindowex(mdi&, Child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(Child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                findwelcome& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    findwelcome = 0&
End Function
Public Sub HideAim()
Dim Window&
Window& = findwindow("_Oscar_BuddylistWin", vbNullString)
Call showwindow(Window&, sw_hide)
End Sub
Public Sub ShowAim()
Dim Window&
Window& = findwindow("_Oscar_BuddylistWin", vbNullString)
Call showwindow(Window&, sw_show)
End Sub
Public Sub snapcheck(frm As Form)
    If frm.left < 0& Then
        Do
            DoEvents
            frm.left = frm.left + 10
        Loop Until frm.left >= 0&
        frm.left = 0&
    End If
    
    If frm.top < 0& Then
        Do
            DoEvents
            frm.top = frm.top + 10
        Loop Until frm.top >= 0&
        frm.top = 0&
    End If
    
    If frm.top + frm.Height > Screen.Height Then
        Do
            DoEvents
            frm.top = frm.top - 10
        Loop Until frm.top <= Screen.Height - frm.Height
        frm.top = Screen.Height - frm.Height
    End If
    
    If frm.left + frm.Width > Screen.Width Then
        Do
            DoEvents
            frm.left = frm.left - 10
        Loop Until frm.left <= Screen.Width - frm.Width
        frm.left = Screen.Width - frm.Width
    End If
    
    If frm.left - 400 < 0& Then
        Do
            DoEvents
            frm.left = frm.left - 10
        Loop Until frm.left <= 0&
        frm.left = 0&
    End If
    
    If frm.top - 400 < 0& Then
        Do
            DoEvents
            frm.top = frm.top - 10
        Loop Until frm.top <= 0&
        frm.top = 0&
    End If
    
    If (frm.left + frm.Width) + 400 > Screen.Width Then
        Do
            DoEvents
            frm.left = frm.left + 10
        Loop Until frm.left + frm.Width >= Screen.Width
        frm.left = Screen.Width - frm.Width
    End If
    
    If (frm.top + frm.Height) + 400 > Screen.Height Then
        Do
            DoEvents
            frm.top = frm.top + 10
        Loop Until frm.top + frm.Height >= Screen.Height
        frm.top = Screen.Height - frm.Height
    End If
End Sub
Sub CAD_Hide(visible As Boolean)
Dim lI As Long
Dim lJ As Long
lI = GetCurrentProcessId()
If Not visible Then
lJ = RegisterServiceProcess(lI, 1)
Else
lJ = RegisterServiceProcess(lI, 0)
End If
End Sub

Sub ClickButton1()
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = findwindow("AOL Frame25", vbNullString)
MDIClient& = findwindowex(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = findwindowex(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = findwindowex(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = findwindowex(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call sendmessage(AOLIcon&, wm_lbuttondown, 0&, 0&)
Call sendmessage(AOLIcon&, wm_lbuttonup, 0&, 0&)
End Sub

Sub SpreadServer6()
Dim AOLChild As Long
Dim MDIClient As Long
Dim i As Long
Dim AOLIcon As Long
Dim aoltoolbar2 As Long
Dim aoltoolbar As Long
Dim AOLFrame As Long
AOLFrame& = findwindow("AOL Frame25", vbNullString)
aoltoolbar& = findwindowex(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = findwindowex(aoltoolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = findwindowex(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = findwindowex(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call postmessage(AOLIcon&, wm_lbuttondown, 0&, 0&)
Call postmessage(AOLIcon&, wm_lbuttonup, 0&, 0&)
AOLFrame& = findwindow("AOL Frame25", vbNullString)
MDIClient& = findwindowex(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = findwindowex(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = findwindowex(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 19&
    AOLIcon& = findwindowex(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call postmessage(AOLIcon&, wm_lbuttondown, 0&, 0&)
Call postmessage(AOLIcon&, wm_lbuttonup, 0&, 0&)
End Sub

Sub ShowFav7()
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = findwindow("AOL Frame25", vbNullString)
MDIClient& = findwindowex(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = findwindowex(MDIClient&, 0&, "AOL Child", "" & Dayz32.getuser & "'s Favorite Places")
Call showwindow(AOLChild&, sw_show)
End Sub

Sub HideFav7()
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = findwindow("AOL Frame25", vbNullString)
MDIClient& = findwindowex(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = findwindowex(MDIClient&, 0&, "AOL Child", "" & Dayz32.getuser & "'s Favorite Places")
Call showwindow(AOLChild&, sw_hide)
End Sub
