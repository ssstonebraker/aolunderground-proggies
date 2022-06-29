Attribute VB_Name = "PUbas"
Option Explicit
Global ServerPaused As Boolean, TillAolRestart As Long, AolDir As String, TillAolKill As Long, AolRestartNum As Long, AolRestartPw As String, ListSize As Long, TillSentKill As Long, KillSentMail As Long, Dead As Boolean, CmdChecker As Long, StatusCheck As Long, StatusReady As Boolean, StatusTime As Long, CmdTime As Long, LAscii, RAscii, MailMsg As String, SentMsg As String, MailzRdy As Long, PendAmt As Long, STOPIT As Boolean, ListPause As Long, FindPause As Long, MailzPause As Long, CommandsReady As Boolean, BlockAmt As Long, FindingStuff As Boolean, ListsReady As Boolean, MailInProgress As Boolean, Total As Long, User$, SN$, Listss&, Count2&, ListTime&, FindNum(99), GetReqs As String, Loading As Boolean
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Public Declare Function findwindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetMenu Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, LParam As Any) As Long
Public Declare Function SendMessageLong& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long)
Public Declare Function SendMessageByString Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As String) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
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
Public Const VK_ESCAPE = &H1B
Public Const VK_D = &O100
Public Const VK_I = &O105
Public Const VK_U = &O117

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_DESTROY = &H2
Public Const WM_MDIDESTROY = &H221
Public Const WM_NCDESTROY = &H82
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const GWL_HINSTANCE = (-6)

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Function FindForwardWindow(Caption As String) As Long
    Dim aol As Long, MDI As Long, Child As Long
    aol& = findwindow("AOL Frame25", vbNullString)  'locates aol
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)    'locates the mdiclient window
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)  'locates an aol child
    If InStr(GetCaption(Child&), Mid(Caption, Int(Len(Caption) / 4), Int((Len(Caption) / 4) * 2))) > 0 And InStr(GetCaption(Child), "Fwd:") > 0 Then    'checks to see if 1/4th of the caption specified is in the title of the window
        FindForwardWindow& = Child& 'if it is then that's the window you're looking for
        Exit Function   'exit the function
    Else    'otherwise
        Do  'begins a loop
            If STOPIT = True Then Exit Function 'see main form
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)  'checks the next aol child
            If InStr(GetCaption(Child&), Mid(Caption, Int(Len(Caption) / 4), Int((Len(Caption) / 4) * 2))) > 0 And InStr(GetCaption(Child), "Fwd:") > 0 Then    'see above
                FindForwardWindow& = Child&   'see above
                Exit Function   'see above
            End If  'ends the if
        Loop Until Child& = 0&  'loops until there are no more child windows
    End If  'ends the if
    FindForwardWindow& = 0& 'sets the window to 0 if none are found
End Function
Public Sub clickfowardbutton(hwnd As Long)
Dim aol As Long, MDI As Long, Child As Long, icon As Long
Dim x
clickicon findfowardbutton(hwnd)    'calls a sub to click the forward icon on the window specified by the user, using the sub findfowardbutton
End Sub
Public Function findfowardbutton(hwnd As Long) As Long
Dim aol As Long, MDI As Long, Child As Long, icon As Long
Dim x
    For x = 1 To 7  'begins a for and next loop that will run 7 times
        icon& = FindWindowEx(hwnd&, icon&, "_AOL_Icon", vbNullString)   'gets the next icon
    Next x  'loops
findfowardbutton = icon&    'sets the function equal to the found forward button
End Function
Public Function FindSendWindow(Caption) As Long
    Dim aol As Long, MDI As Long, Child As Long
    aol& = findwindow("AOL Frame25", vbNullString)  'locates aol
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)    'locates the mdiclient window
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)  'locates the activated aol child
    If InStr(GetCaption(Child&), Mid(Caption, Int(Len(Caption) / 4), Int((Len(Caption) / 4) * 2))) > 0 Then 'see findforwardwindow
        FindSendWindow& = Child&    'sets the function to the found window
        Exit Function   'exits the function
    Else    'otherwise...
        Do  'begins a loop
            If STOPIT = True Then Exit Function 'see main form
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)  'finds the next child
        If InStr(GetCaption(Child&), Mid(Caption, Int(Len(Caption) / 4), Int((Len(Caption) / 4) * 2))) > 0 Then 'see above
                FindSendWindow& = Child&    'see above
                Exit Function    'see above
            End If  'ends the if
        Loop Until Child& = 0&  'loops until no more childs are found
    End If
    FindSendWindow& = 0&    'sets the function = to 0 if no windows with the specified criteria were found
End Function
Public Sub MailOpenFlash()
    Dim aol As Long, tool As Long, ToolBar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim curpos As POINTAPI, winvis As Long, Current As Long
    Dim x As Long
    aol& = findwindow("AOL Frame25", vbNullString)  'locates aol
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString) 'locates the mdiclient window
    ToolBar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)    'locates the aol toolbar
    toolicon& = FindWindowEx(ToolBar&, 0&, "_AOL_Icon", vbNullString)   'locates the first icon
    toolicon& = FindWindowEx(ToolBar&, toolicon&, "_AOL_Icon", vbNullString)    'locates the next icon
    toolicon& = FindWindowEx(ToolBar&, toolicon&, "_AOL_Icon", vbNullString)    'locates the next icon
openit: 'label used to goto if the popup menu doesn't popup
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&) 'sends the api command button down to the icon
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)   'lets the button up(this just simulated a click)
    Current = Timer 'see main form
    Do: DoEvents    'another loop
        If Timer - Current >= 1 Then GoTo openit    'after 1 second has elapsed it goes back to the clicking part
        If STOPIT = True Then Exit Sub  'see main form
        smod& = findwindow("#32768", vbNullString)  'locates the window with the classname #32768
        winvis& = IsWindowVisible(smod&)    'checks to see if the window is visible
    Loop Until winvis& = 1  'loops until the window is visible
    For x = 1 To 12
        Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)    'pushes the down key
        Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)  'lets go of the down key
    Next x
    Call PostMessage(smod&, WM_KEYDOWN, VK_RIGHT, 0&)   'pushes the right key
    Call PostMessage(smod&, WM_KEYUP, VK_RIGHT, 0&) 'lets go of the right key
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)  'pushes the enter key
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)    'lets go of the enter key
End Sub
Public Sub aol4_runpopup2(iconnum As Long, Char As String)
Dim aol As Long, tool As Long, icon As Long
Dim geticon As Long, themenu As Long
Dim thecursor As POINTAPI, theicon As Long
'for this sub see the open mail flash sub
aol& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL ToolBar", vbNullString)
tool& = FindWindowEx(tool&, 0&, "_AOL_ToolBar", vbNullString)
theicon& = FindWindowEx(tool&, 0&, "_AOL_Icon", vbNullString)
For geticon& = 2 To iconnum&
theicon& = FindWindowEx(tool&, theicon&, "_AOL_Icon", vbNullString)
Next
Call clickicon(theicon&)
Do: DoEvents
    If STOPIT = True Then Exit Sub
    themenu& = FindWindowEx(0, 0&, "#32768", vbNullString)
Loop Until IsWindowVisible(themenu&) <> 0&
Call PostMessage(themenu&, WM_CHAR, Asc(Char$), 0)
End Sub
Public Sub MailOpenNew()
Dim counter As Long, counter2 As Long, counter3 As Long
Call aol4_runpopup2(3, "R") 'opens the new mail box
    Do: DoEvents
        If STOPIT = True Then Exit Sub
        counter& = MailCountNew&    'uses a sub to count the mailbox
            Pause 0.65
        counter2& = MailCountNew&   'uses a sub to count the mailbox
            Pause 0.65
        counter3& = MailCountNew&   'uses a sub to count the mailbox
            Pause 0.65
    Loop Until counter& = counter2& And counter2& = counter3&   'makes sure all 3 counts are the same before proceding
End Sub
Public Sub MailOpenEmailFlash(index As Long)
    Dim aol As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long
    aol& = findwindow("AOL Frame25", vbNullString)  'locates aol
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)    'locates the mdiclient window
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail") 'locates the flashmail box
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)    'locates the _AOL_Tree in the flashmail box, this is the list of mailz
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)  'counts the mail
    If fCount& < index& Then Exit Sub   'checks to be sure the mail you're trying to open is not larger then your total mail
    Call SendMessage(fList&, LB_SETCURSEL, index&, 0&)  'selects the mail with the index specified
    Call PostMessage(fList&, WM_KEYDOWN, VK_RETURN, 0&) 'presses the enter key on that mail
    Call PostMessage(fList&, WM_KEYUP, VK_RETURN, 0&)   'releases the enter key
End Sub
Public Function MailCountFlash() As Long
    Dim aol As Long, MDI As Long, fMail As Long, fList As Long
    Dim count As Long
    aol& = findwindow("AOL Frame25", vbNullString)  'locates the aol window
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)    'locates the mdiclient window
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail") 'locates the flashmail box
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)    'locates the _AOL_Tree(see above)
    count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)   'gets the count of the listbox
    MailCountFlash& = count&    'sets the function = to the count
End Function

Public Function FindMailBox() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim TabControl As Long, TabPage As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(Child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = Child&
        Exit Function
    Else
        Do
            If STOPIT = True Then Exit Function
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
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountNew& = count&
End Function
Public Function MailCountSent() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountSent& = count&
End Function
Public Function MailCountOld() As Long
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    MailCountOld& = count&
End Function
Public Sub MailDeleteNewByIndex(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, count As Long, dButton As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = MailCountSent
    If index& > count& - 1 Or index& < 0& Then Exit Sub
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
    Dim mTree As Long, count As Long, dButton As Long
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
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
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
    Dim mTree As Long, count As Long, dButton As Long
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
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchBox& = 0& To count& - 1
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
    Dim mTree As Long, count As Long, dButton As Long
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
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0& Then Exit Sub
    For SearchBox& = 0& To count& - 1
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
    Dim aol As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, spot1 As Long, spot2 As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
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
    Dim count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    spot1& = InStr(MyString$, Chr(9))
    spot2& = InStr(spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, spot1& + 1, spot2& - spot1& - 1)
    MailSenderNew$ = MyString$
End Function
Public Function MailSubjectFlash(index As Long) As String
    Dim aol As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, sLength As Long
    Dim MyString As String, spot As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
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
    spot& = InStr(MyString$, Chr(9))
    spot& = InStr(spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectFlash$ = MyString$
End Function
Public Function MailSubjectNew(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Function
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Or index& > count& - 1 Or index& < 0& Then Exit Function
    sLength& = SendMessage(mTree&, LB_GETTEXTLEN, index&, 0&)
    MyString$ = String(sLength& + 1, 0)
    Call SendMessageByString(mTree&, LB_GETTEXT, index&, MyString$)
    spot& = InStr(MyString$, Chr(9))
    spot& = InStr(spot& + 1, MyString$, Chr(9))
    MyString$ = Right(MyString$, Len(MyString$) - spot&)
    MyString$ = ReplaceString(MyString$, Chr(0), "")
    MailSubjectNew$ = MyString$
End Function
Public Sub MailToListNew(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim spot As Long, MyString As String, count As Long
    Dim num, percent
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Then Exit Sub
    num = 100 / count&
    percent = num
    For AddMails& = 0 To count& - 1
        If Mid(percent, InStr(1, percent, ".") + 2, 1) > 5 Or Mid(percent, InStr(1, percent, ".") + 2, 1) = "" Then FrmMain!Percentlbl = Left(percent, InStr(1, percent, ".") - 1) & "%"
        percent = percent + num
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        spot& = InStr(MyString$, Chr(9))
        spot& = InStr(spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - spot&)
        thelist.AddItem MyString$
    Next AddMails&
    FrmMain.Percentlbl = "-------->"
End Sub
Public Sub MailToListOld(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        spot& = InStr(MyString$, Chr(9))
        spot& = InStr(spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub
Public Sub MailToListSent(thelist As ListBox)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim spot As Long, MyString As String, count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If count& = 0 Then Exit Sub
    For AddMails& = 0 To count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        spot& = InStr(MyString$, Chr(9))
        spot& = InStr(spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - spot&)
        thelist.AddItem MyString$
    Next AddMails&
End Sub
Public Sub SendMail(Person As String, Subject As String, message As String)
    Dim aol As Long, MDI As Long, tool As Long, ToolBar As Long
    Dim toolicon As Long, OpenSend As Long, DoIt As Long
    Dim rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    ToolBar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, toolicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        If STOPIT = True Then Exit Sub
        DoEvents
        OpenSend& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And rich& <> 0& And SendButton& <> 0& And combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    DoEvents
    Call SendMessageByString(rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailForward(Caption As String, SendTo As String, message As String, DeleteFwd As Boolean)
    Dim aol As Long, MDI As Long, Error As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, rich As Long, fCombo As Long
    Dim combo As Long, Button1 As Long, Button2 As Long
    Dim TempSubject As String, x As Long, tracker As Long, fullwindow As Long
    Dim fullbutton As Long, b, Current As Long, achild As Long, view As Long
    Dim ErrorMes As String, GayPerson As String, Killz As Long
    OpenForward& = FindForwardWindow(Caption)
    If OpenForward& = 0 Then Exit Sub
    Do
        If STOPIT = True Then Exit Sub
        If FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0& Then
            Call RunMenuByString("S&top Incoming Text")
            Call SendMessageLong(FindWindowEx(MDI&, 0&, "AOL Child", "Status"), WM_CLOSE, 0&, 0&)
        End If
        checkdeadmsg
        DoEvents
        OpenSend& = FindForwardWindow(Caption)
        EditTo& = FindWindowEx(OpenSend&, 0&, "_AOL_Edit", vbNullString)
        EditCC& = FindWindowEx(OpenSend&, EditTo&, "_AOL_Edit", vbNullString)
        EditSubject& = FindWindowEx(OpenSend&, EditCC&, "_AOL_Edit", vbNullString)
        rich& = FindWindowEx(OpenSend&, 0&, "RICHCNTL", vbNullString)
        combo& = FindWindowEx(OpenSend&, 0&, "_AOL_Combobox", vbNullString)
        fCombo& = FindWindowEx(OpenSend&, 0&, "_AOL_Fontcombo", vbNullString)
        Button1& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        Button2& = FindWindowEx(OpenSend&, Button1&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 13
            checkdeadmsg
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And rich& <> 0& And SendButton& <> 0& And combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
    If DeleteFwd = True Then
        checkdeadmsg
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
    End If
    Call SetText(EditTo&, SendTo$)
    'current = Timer
    Do Until InStr(1, GetText(EditTo&), SendTo$)
        SetText EditTo&, SendTo$ 'If Timer - current > 0.1 Then
    Loop
    Call SendMessageByString(rich&, WM_SETTEXT, 0, message$)
    Do Until FindForwardWindow(Caption) = 0&
        If STOPIT = True Then Exit Sub
        checkdeadmsg
        SendButton& = FindWindowEx(OpenSend&, 0&, "_AOL_Icon", vbNullString)
        For DoIt& = 1 To 11
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
        If Options.mnufastr.Checked = True Then
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        Call ShowWindow(OpenSend, SW_HIDE)
        Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
        Call SendMessageLong(OpenSend&, WM_DESTROY, 0&, 0&)
        Killwin OpenSend&
        Do Until FindForwardWindow(Caption) = 0
            Call ShowWindow(FindForwardWindow(Caption), SW_HIDE)
            Call SendMessageLong(FindForwardWindow(Caption), WM_DESTROY, 0&, 0&)
            Killwin FindForwardWindow(Caption)
        Loop
Else
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Current = Timer
    Do Until FindForwardWindow(Caption) = 0&
         If Timer - Current >= 3 Then
            Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
            Current = Timer
        End If
    Loop
End If
Loop

Do Until achild = 0&
achild& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
If achild& <> 0 Then
    Do
    DoEvents
    view& = FindWindowEx(achild&, 0&, "_AOL_View", vbNullString)
    Loop Until view& <> 0
    Do
    DoEvents
    ErrorMes$ = WinTxt(view&)
    Loop Until ErrorMes$ <> ""
    ErrorMes$ = Right(ErrorMes$, Len(ErrorMes$) - InStr(ErrorMes$, ":") - 4)
    GayPerson$ = Left(ErrorMes$, InStr(ErrorMes$, "-") - 2)
    For x = 0 To FrmMain.Pending.ListCount - 1
    If InStr(FrmMain.Pending.List(x), GayPerson$) > 0 Then
        FrmMain.Pending.RemoveItem x - Killz
        Killz = Killz + 1
    End If
    Next x
    For x = 0 To FrmMain.Lists.ListCount - 1
        If InStr(FrmMain.Pending.List(x), GayPerson$) > 0 Then FrmMain.Lists.RemoveItem x
    Next x
    'FrmMain.Chat1.Ignore(gayPerson$) = True
    Killwin achild&
    Pause 0.2
    Killwin achild&
End If
Loop
End Sub
Public Sub CloseOpenMails()
    Dim OpenSend As Long, OpenForward As Long
    Do
        If STOPIT = True Then Exit Sub
        DoEvents
        OpenSend& = FindSendWindow("")
        OpenForward& = FindForwardWindow("")
        Call PostMessage(OpenSend&, WM_CLOSE, 0&, 0&)
        DoEvents
        Call PostMessage(OpenForward&, WM_CLOSE, 0&, 0&)
        DoEvents
    Loop Until OpenSend& = 0& And OpenForward& = 0&
End Sub
Public Sub MailDeleteFlashByIndex(index As Long)
    Dim aol As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
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
    Dim aol As Long, MDI As Long, fMail As Long, fList As Long
    Dim fCount As Long, DeleteButton As Long, SearchFor As Long
    Dim SearchBox As Long, CurCaption As String
    Dim sSender As String, sSubject As String
    Dim cSender As String, cSubject As String
    Dim ParHand1 As Long, OurParent As Long, ourhandle As Long
    Dim num, percent, kilthim As Long
    FrmMain.Statuslbl = "Killing Dupes"
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    fCount& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If fCount& < 2& Then Exit Sub
    DeleteButton& = FindWindowEx(fMail&, 0&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    DeleteButton& = FindWindowEx(fMail&, DeleteButton&, "_AOL_Icon", vbNullString)
    CurCaption$ = VBForm.Caption
    num = (100 / ((fCount& - 2) * (fCount& - 2))) / 2
    percent = num
    If fCount& = 0& Then Exit Sub
    For SearchFor& = 0& To fCount& - kilthim
        FrmMain.Statuslbl = "Killing " & SearchFor& & "/" & fCount& - 2
        DoEvents
        sSubject$ = MailSubjectFlash(SearchFor&)
        For SearchBox& = SearchFor& + 1 To fCount& - (kilthim + 1)
            If Mid(percent, InStr(1, percent, ".") + 2, 1) > 5 Or Mid(percent, InStr(1, percent, ".") + 2, 1) = "" Then FrmMain!Percentlbl = Left(percent, InStr(1, percent, ".") - 1) & "%"
            percent = percent + num
            cSubject$ = MailSubjectFlash(SearchBox&)
            If cSubject$ = sSubject$ Then
                Call SendMessage(fList&, LB_SETCURSEL, SearchBox&, 0&)
                DoEvents
                Call PostMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                Do: DoEvents
                ourhandle& = findwindow("#32770", "America Online")
                ourhandle& = FindWindowEx(ourhandle&, 0, "Button", "&Yes")
                Loop Until ourhandle& <> 0
                Call PostMessage(ourhandle&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(ourhandle&, WM_LBUTTONUP, 0&, 0&)
                SearchBox& = SearchBox& - 1
                kilthim = kilthim + 1
            End If
        Next SearchBox&
    Next SearchFor&
FrmMain.Statuslbl = "Stopped"
End Sub
Public Sub SetMailPrefs()
    Dim aol As Long, tool As Long, ToolBar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim MDI As Long, mprefs As Long, mbutton As Long
    Dim gstatic As Long, mstatic As Long, fstatic As Long
    Dim mastatic As Long, dmod As Long, confirmcheck As Long
    Dim closecheck As Long, spellcheck As Long, OkButton As Long
    Dim curpos As POINTAPI, winvis As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    ToolBar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, toolicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        If STOPIT = True Then Exit Sub
        smod& = findwindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    For dothis& = 1 To 7
        Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Next dothis&
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Do
        If STOPIT = True Then Exit Sub
        dmod& = findwindow("_AOL_Modal", "Mail Preferences")
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
    'Call PostMessage(mprefs&, WM_CLOSE, 0&, 0&)
End Sub
Public Function FindRoom() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim rich As Long, aollist As Long
    Dim AOLIcon As Long, AolStatic As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    aollist& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AolStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    If rich& <> 0& And aollist& <> 0& And AOLIcon& <> 0& And AolStatic& <> 0& Then
        FindRoom& = Child&
        Exit Function
    Else
        Do
            If STOPIT = True Then Exit Function
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            aollist& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AolStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            If rich& <> 0& And aollist& <> 0& And AOLIcon& <> 0& And AolStatic& <> 0& Then
                FindRoom& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindRoom& = Child&
End Function
Public Function FindInfoWindow() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AolStatic As Long
    Dim AOLIcon2 As Long, AolGlyph As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
    AolStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    AolGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AolStatic& <> 0& And AolGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
            AolStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            AolGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AolStatic& <> 0& And AolGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindInfoWindow& = Child&
End Function
Public Function RoomCount() As Long
    Dim aol As Long, MDI As Long, rMail As Long, rList As Long
    Dim count As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    RoomCount& = count&
End Function
Public Sub AddRoomToListBox(thelist As ListBox, adduser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rbytes)
            Call CopyMemory(psnHold&, ByVal screenname$, 4)
            psnHold& = psnHold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ Or adduser = True Then
                thelist.AddItem screenname$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub AddRoomToCombobox(TheCombo As ComboBox, adduser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rbytes)
            Call CopyMemory(psnHold&, ByVal screenname$, 4)
            psnHold& = psnHold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ Or adduser = True Then
                TheCombo.AddItem screenname$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
    If TheCombo.ListCount > 0 Then
        TheCombo.Text = TheCombo.List(0)
    End If
End Sub
Public Sub ChatIgnoreByIndex(index As Long)
    Dim room As Long, sList As Long, iWindow As Long
    Dim iCheck As Long, a As Long, count As Long
    count& = RoomCount&
    If index& > count& - 1 Then Exit Sub
    room& = FindRoom&
    sList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
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
    Dim cProcess As Long, itmHold As Long, screenname As String
    Dim psnHold As Long, rbytes As Long, index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rList& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rbytes)
            Call CopyMemory(psnHold&, ByVal screenname$, 4)
            psnHold& = psnHold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If screenname$ <> getuser$ And LCase(screenname$) = LCase(name$) Then
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
Public Function GetText2(window As Long) As String
'gets the text of a window.
Dim Buffer As String, TextLength As Long
TextLength& = SendMessage(window, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = String(TextLength&, 0&)
Call SendMessageByString(window, WM_GETTEXT, TextLength& + 1, Buffer$)
GetText2$ = Buffer$
End Function
Public Function WaitForOKOrRoom(room As String) As String
    Dim RoomTitle As String, fullwindow As Long, fullbutton As Long
    room$ = LCase(ReplaceString(room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindRoom&)
        RoomTitle$ = LCase(ReplaceString(room$, " ", ""))
        fullwindow& = findwindow("#32770", "America Online")
        fullbutton& = FindWindowEx(fullwindow&, 0&, "Button", "OK")
    Loop Until (fullwindow& <> 0& And fullbutton& <> 0&) Or room$ = RoomTitle$
    DoEvents
    If fullwindow& <> 0& Then
        Do
            DoEvents
            Call SendMessage(fullbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYUP, VK_SPACE, 0&)
            fullwindow& = findwindow("#32770", "America Online")
            fullbutton& = FindWindowEx(fullwindow&, 0&, "Button", "OK")
            WaitForOKOrRoom = "OK"
        Loop Until fullwindow& = 0& And fullbutton& = 0&
    End If
    DoEvents
End Function
Public Sub MemberRoom(room As String)
    Call keyword("aol://2719:61-2-" & room$)
End Sub
Public Sub PublicRoom(room As String)
    Call keyword("aol://2719:21-2-" & room$)
End Sub
Public Sub PrivateRoom(room As String)
    Call keyword("aol://2719:2-2-" & room$)
End Sub
Public Sub InstantMessage(Person As String, message As String)
    Dim aol As Long, MDI As Long, im As Long, rich As Long
    Dim SendButton As Long, ok As Long, Button As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(im&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(rich&, WM_SETTEXT, 0&, message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        ok& = findwindow("#32770", "America Online")
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until ok& <> 0& Or im& = 0&
    If ok& <> 0& Then
        Button& = FindWindowEx(ok&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(im&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Function CheckIMs(Person As String) As Boolean
    Dim aol As Long, MDI As Long, im As Long, rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://9293:" & Person$)
    Do
        DoEvents
        im& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        rich& = FindWindowEx(im&, 0&, "RICHCNTL", vbNullString)
        Available1& = FindWindowEx(im&, 0&, "_AOL_Icon", vbNullString)
        Available2& = FindWindowEx(im&, Available1&, "_AOL_Icon", vbNullString)
        Available3& = FindWindowEx(im&, Available2&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available3&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
        Available& = FindWindowEx(im&, Available&, "_AOL_Icon", vbNullString)
    Loop Until im& <> 0& And rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = findwindow("#32770", "America Online")
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
    Else
        CheckIMs = False
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(im&, WM_CLOSE, 0&, 0&)
End Function
Public Sub keyword(KW As String)
    Dim aol As Long, tool As Long, ToolBar As Long
    Dim combo As Long, editwin As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    ToolBar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    combo& = FindWindowEx(ToolBar&, 0&, "_AOL_Combobox", vbNullString)
    editwin& = FindWindowEx(combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(editwin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(editwin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(editwin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Function ReplaceString(MyString As String, toFinD As String, replacewith As String) As String
    Dim spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, newstring As String
    spot& = InStr(LCase(MyString$), LCase(toFinD))
    NewSpot& = spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If spot& + Len(toFinD$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(toFinD$) + 1)
            Else
                RightString = ""
            End If
            newstring$ = LeftString$ & replacewith$ & RightString$
            MyString$ = newstring$
        Else
            newstring$ = MyString$
        End If
        spot& = NewSpot& + Len(replacewith$)
        If spot& > 0 Then
            NewSpot& = InStr(spot&, LCase(MyString$), LCase(toFinD$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = newstring$
End Function
Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Function FileGetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function
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
Public Sub FileSetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Public Function CheckIfMaster() As Boolean
    Dim aol As Long, MDI As Long, pWindow As Long
    Dim pButton As Long, Modal As Long, mstatic As Long
    Dim mString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    Call keyword("aol://4344:1580.prntcon.12263709.564517913")
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
        Modal& = findwindow("_AOL_Modal", vbNullString)
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
Public Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Public Function GetListText(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, Buffer$)
    GetListText$ = Buffer$
End Function
Public Sub Button(mbutton As Long)
    Call SendMessage(mbutton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mbutton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub icon(aIcon As Long)
    Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub CloseWindow(window As Long)
    Call PostMessage(window&, WM_CLOSE, 0&, 0&)
End Sub
Public Function getuser() As String
    Dim aol As Long, MDI As Long, Welcome As Long
    Dim Child As Long, UserString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(Child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        getuser$ = UserString$
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
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
Public Sub Pause(Duration)

    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub playwav(wavfile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(wavfile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(wavfile$, SND_FLAG)
    End If
End Sub
Public Sub SetText(window As Long, Text As String)
    Call SendMessageByString(window&, WM_SETTEXT, 0&, Text$)
End Sub
Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormDrag(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub WindowHide(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub
Public Sub WindowShow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub
Public Sub runmenu(TopMenu As Long, SubMenu As Long)
    Dim aol As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(aol&, WM_COMMAND, mnID&, 0&)
End Sub
Public Sub RunMenuByString(SearchString As String)
    Dim aol As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aol& = findwindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Public Function aol4_welcomescreen() As Long
'gets the aol4 welcome screen.
Dim aol As Long, MDI As Long, Child As Long
aol& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)

If InStr(GetCaption(Child&), "Welcome, ") <> 0& Then
    aol4_welcomescreen& = Child&
    Exit Function
Else
    Do
        If STOPIT = True Then Exit Function
        Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
        If InStr(GetCaption(Child&), "Welcome, ") <> 0& Then
            aol4_welcomescreen& = Child&
            Exit Function
        End If
    Loop Until Child& = 0&
End If
aol4_welcomescreen& = 0&
End Function
Public Sub aol4_killwait()
'kills the aol hourglass.
Dim aol As Long, aolmodal As Long, AolGlyph As Long
Dim AolStatic As Long, AOLIcon As Long, AolInstance As Long
aol& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
'AOLInst = GetWindowWord(aol&, GWL_HINSTANCE)
'call createcursor(aolinst,
'Call SetCursor(vbArrow)
Call RunMenuByString("&About America Online")
Do: DoEvents
    If STOPIT = True Then Exit Sub
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
AolGlyph& = FindWindowEx(aolmodal&, 0&, "_AOL_Glyph", vbNullString)
AolStatic& = FindWindowEx(aolmodal&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(aolmodal&, 0&, "_AOL_Icon", vbNullString)
Loop Until aolmodal& <> 0& And AolGlyph <> 0& And AolStatic& <> 0& And AOLIcon& <> 0& '

Do: DoEvents
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
Call PostMessage(aolmodal&, WM_CLOSE, 0, 0&)
Loop Until aolmodal& = 0&
End Sub
Public Sub aol4_mail_send(SN As String, Subject As String, message As String, killerror As Boolean)
Dim aol As Long, MDI As Long, email As Long, emailedit As Long
Dim emailrich As Long, emailicon As Long, x
Dim theerror As Long, tool As Long, mailicon As Long
Dim getemailicon As Long, SubjectBox As Long
Dim nowindow As Long, nobutton As Long, trackit
Dim Current As Long, clicked As Long
openit:
aol& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL ToolBar", vbNullString)
tool& = FindWindowEx(tool&, 0&, "_AOL_ToolBar", vbNullString)
mailicon& = FindWindowEx(tool&, 0&, "_AOL_Icon", vbNullString)
mailicon& = FindWindowEx(tool&, mailicon&, "_AOL_Icon", vbNullString)
clickicon (mailicon&)
Current = Timer
Do
    If STOPIT = True Then Exit Sub
DoEvents
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
email& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
emailedit& = FindWindowEx(email&, 0&, "_AOL_Edit", vbNullString)
emailrich& = FindWindowEx(email&, 0&, "RICHCNTL", vbNullString)
emailicon& = FindWindowEx(email&, 0&, "_AOL_Icon", vbNullString)
If email& <> 0 And emailedit& <> 0 And emailrich& <> 0 Then Exit Do
If Timer - Current >= 3 Then GoTo openit
Loop

SubjectBox& = FindWindowEx(email&, emailedit&, "_AOL_EDIT", vbNullString)
SubjectBox& = FindWindowEx(email&, SubjectBox&, "_AOL_EDIT", vbNullString)

Call SendMessageByString(emailedit&, WM_SETTEXT, 0, SN$)
Call SendMessageByString(SubjectBox&, WM_SETTEXT, 0, Subject$)
Call SendMessageByString(emailrich&, WM_SETTEXT, 0, message$)

For getemailicon& = 1 To 13
emailicon& = FindWindowEx(email&, emailicon&, "_AOL_Icon", vbNullString)
Next getemailicon&

Do: DoEvents
    If STOPIT = True Then Exit Sub
email& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
clickicon (emailicon&)
Pause 0.7
clicked = clicked + 1
Loop Until email& = 0 Or clicked = 5
Do: DoEvents
Loop Until FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail") = 0
End Sub
Public Sub clickicon(icon As Long)
'clicks any icon.
Call PostMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
DoEvents
Call PostMessage(icon&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub WaitForOK()
    Dim fullwindow As Long, fullbutton As Long
    Do
        DoEvents
        fullwindow& = findwindow("#32770", "AOL Mail")
        fullbutton& = FindWindowEx(fullwindow&, 0&, "Button", "&No")
    Loop Until fullwindow& <> 0& And fullbutton& <> 0&
    DoEvents
    If fullwindow& <> 0& Then
        Do
            DoEvents
            Call SendMessage(fullbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYUP, VK_SPACE, 0&)
            fullwindow& = findwindow("#32770", "AOL Mail")
            fullbutton& = FindWindowEx(fullwindow&, 0&, "Button", "&No")
        Loop Until fullwindow& = 0& And fullbutton& = 0&
    End If
    DoEvents
End Sub
Public Function checkdeadmsg() As String
Dim fullwindow&, fullbutton&, aol&, MDI&, Child&
    fullwindow& = findwindow("#32770", "America Online")
    fullbutton& = FindWindowEx(fullwindow&, 0&, "Button", "OK")
    If fullwindow& <> 0& Then
        Do
            If STOPIT = True Then Exit Function
            DoEvents
            Call SendMessage(fullbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYUP, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call SendMessage(fullbutton&, WM_KEYUP, VK_SPACE, 0&)
            fullwindow& = findwindow("#32770", "America Online")
            fullbutton& = FindWindowEx(fullwindow&, 0&, "Button", "OK")
        Loop Until fullwindow& = 0& And fullbutton& = 0&
        checkdeadmsg = "True"
    End If
Dim achild&, view&, ErrorMes As String, GayPerson$, Killz, x
Do Until achild = 0&
achild& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
If achild& <> 0 Then
    Do
    DoEvents
    view& = FindWindowEx(achild&, 0&, "_AOL_View", vbNullString)
    Loop Until view& <> 0
    Do
    DoEvents
    ErrorMes$ = WinTxt(view&)
    Loop Until ErrorMes$ <> ""
    ErrorMes$ = Right(ErrorMes$, Len(ErrorMes$) - InStr(ErrorMes$, ":") - 4)
    GayPerson$ = Left(ErrorMes$, InStr(ErrorMes$, "-") - 2)
    For x = 0 To FrmMain.Pending.ListCount - 1
    If InStr(FrmMain.Pending.List(x), GayPerson$) > 0 Then
        FrmMain.Pending.RemoveItem x - Killz
        Killz = Killz + 1
    End If
    Next x
    For x = 0 To FrmMain.Lists.ListCount - 1
        If InStr(FrmMain.Pending.List(x), GayPerson$) > 0 Then FrmMain.Lists.RemoveItem x
    Next x
    Killwin achild&
    Pause 0.2
    Killwin achild&
End If
Loop
End Function
Sub Killwin(Windo)
   Call SendMessageLong(Windo, WM_CLOSE, 0&, 0&)
End Sub
Function WinTxt(ByVal hwnd As Integer)
Dim x As Integer
Dim y As String
Dim z As Integer
x = SendMessage(hwnd, &HE, 0&, 0&)
y = String(x + 1, " ")
z = SendMessageByString(hwnd, &HD, x + 1, y)
WinTxt = Left(y, x)
End Function
Public Sub MailRunFlash()
    Dim aol As Long, tool As Long, ToolBar As Long
    Dim toolicon As Long, dothis As Long, smod As Long
    Dim MDI As Long, mprefs As Long, mbutton As Long
    Dim gstatic As Long, mstatic As Long, fstatic As Long
    Dim mastatic As Long, dmod As Long, confirmcheck As Long
    Dim closecheck As Long, spellcheck As Long, OkButton As Long
    Dim curpos As POINTAPI, winvis As Long, icon As Long, Modal As Long
    Dim Current As Long, MyHandle As Long, ourhandle As Long
    aol& = findwindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    ToolBar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, 0&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, toolicon&, "_AOL_Icon", vbNullString)
    toolicon& = FindWindowEx(ToolBar&, toolicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(toolicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(toolicon&, WM_LBUTTONUP, 0&, 0&)
    Do
        If STOPIT = True Then Exit Sub
        smod& = findwindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    For dothis& = 1 To 11
        Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Next dothis&
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Do: DoEvents
    Modal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
    Loop Until Modal& <> 0&
    Do: DoEvents
    icon& = FindWindowEx(Modal&, 0, "_AOL_Icon", vbNullString)
    Loop Until icon& <> 0
    Do Until FindWindowEx(MDI&, 0&, "AOL Child", "Status") <> 0&
        clickicon icon&
        Pause 0.5
    Loop
    Do: DoEvents
        MyHandle = FindWindowEx(MDI&, 0&, "AOL Child", "Status")
        ourhandle& = FindWindowEx(MyHandle&, 0, "_AOL_View", vbNullString)
    Loop Until ourhandle& <> 0 And MyHandle <> 0
    Do: DoEvents
    Loop Until InStr(GetText(ourhandle&), "AOL session") > 0
    Killwin MyHandle&
End Sub
Sub textset(hwnd As Long, What As String)
Dim SetIt
SetIt = SendMessageByString(hwnd, &HC, 0, What)
End Sub
Public Function LineCount(MyString As String) As Long
    Dim spot As Long, count As Long
    If Len(MyString$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    spot& = InStr(MyString$, Chr(13))
    If spot& <> 0& Then
        LineCount& = 1
        Do
            spot& = InStr(spot + 1, MyString$, Chr(13))
            If spot& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until spot& = 0&
    End If
        LineCount& = LineCount& + 1
End Function
Public Sub Get_Reqs()
Dim lblock As Integer, rblock As Integer
Dim thechat&, theview&, alltext$, theenter%, thisline$, Who, Wht$, Request, checkmax
Dim x, spot%, lblockc, rblockc, y
FrmMain.Statuslbl = "Get'n Requests"
thechat& = FindRoom
theview& = FindWindowEx(thechat&, 0&, "RICHCNTL", vbNullString)
alltext$ = WinTxt(theview&)
If alltext$ = "" Then GoTo TheEnd
textset theview&, ""
If InStr(alltext$, Chr(13)) = 0 Then GoTo TheEnd
alltext$ = Right$(alltext$, Len(alltext$) - 1)
alltext$ = alltext$ + Chr(13)
Do
DoEvents
    If InStr(alltext$, Chr(13)) = 0 Then GoTo TheEnd
    theenter% = InStr(alltext$, Chr(13))
    thisline$ = Left$(alltext$, theenter% - 1)
    If InStr(thisline$, ":") = 0 Then
        GoTo LoopPart
    End If
    Who = Mid$(thisline$, 1, InStr(thisline$, ":") - 1)
    Wht$ = Right$(thisline$, Len(thisline$) - Len(Who) - 3)
    If Len(MChat.Text1) >= 32000 Then MChat.Text1 = ""
    MChat.Text1 = MChat.Text1 & Who & ": " & Chr(9) & " " & Wht$ & vbCrLf
    MChat.Text1.SelStart = Len(MChat.Text1) - 1
    For x = 0 To Ban.List2.ListCount - 1
    If Ban.List2.List(x) = Who Then GoTo TheEnd
    Next x
    If Left$(UCase(Wht$), 10) = UCase("<aolpromo>") Then Wht$ = Right$(Wht$, Len(Wht$) - 10)
    
    checkmax = 0
    'trimming response for the request portion and name portion
    If LCase(Left(Wht$, 7 + Len(User$))) = LCase("/" & User$ & " send ") Then
        Request = Right(Wht$, Len(Wht$) - (Len(User$) + 7))
    'adding standard request
        If IsNumeric(Request) = True And Request <= FrmMain.Mailz.ListCount - 1 Then
        For x = 0 To FrmMain.Pending.ListCount - 1
            If FrmMain.Pending.List(x) = Who & "-" & Request Then Exit Sub
            If InStr(FrmMain.Pending.List(x), Who) > 0 Then checkmax = checkmax + 1
        Next x
        If PendAmt = 0 Then GoTo dontcheck
        If checkmax >= PendAmt Then
            ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & Who & ", You Have Reached The Max Pending" & RAscii
            Pause 0.7
            Exit Sub
        End If
dontcheck:
        FrmMain.Pending.AddItem Who & "-" & Request
'adding list request
    ElseIf LCase(Request) = "list" Then
        For x = 0 To FrmMain.Lists.ListCount - 1
            If FrmMain.Lists.List(x) = Who & "-" & "list" Then Exit Sub
        Next x
        FrmMain.Lists.AddItem Who & "-" & "list"
'adding status request
    ElseIf LCase(Request) = "status" Then
        For x = 0 To FrmMain.Pending.ListCount - 1
            If FrmMain.Pending.List(x) = Who & "+" & "status" Then Exit Sub
        Next x
        FrmMain.Pending.AddItem Who & "+" & "status"
'adding blocks
    ElseIf InStr(Request, "-") > 0 Then
        spot% = InStr(Request, "-")
        lblockc = Left(Request, spot% - 1)
        rblockc = Right(Request, Len(Request) - spot%)
        If IsNumeric(rblockc) = False Or IsNumeric(lblockc) = False Then Exit Sub
        lblock = lblockc
        rblock = rblockc
        If rblock < lblock Then
        ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & Who & ", Invalid Block Size" & RAscii
        Exit Sub
        End If
        If rblock - lblock > BlockAmt Then Exit Sub
        If rblock > FrmMain.Mailz.ListCount - 1 Or lblock > FrmMain.Mailz.ListCount - 1 Then Exit Sub
        If rblock < 0 Or lblock < 0 Then Exit Sub
        For x = lblock To rblock
            checkmax = 0
            For y = 0 To FrmMain.Pending.ListCount - 1
                If FrmMain.Pending.List(y) = Who & "-" & x Then GoTo skipit
                If InStr(FrmMain.Pending.List(y), Who) > 0 Then checkmax = checkmax + 1
            Next y
            If PendAmt = 0 Then GoTo dontcheckblock
            If checkmax >= PendAmt Then
                ChatSend "</u></i></b><font face=" & Chr(34) & "verdana" & Chr(34) & ">" & LAscii & "" & Who & ", You Have Reached The Max Pending" & RAscii
                Pause 0.7
                Exit Sub
            End If
dontcheckblock:
            FrmMain.Pending.AddItem Who & "-" & x
skipit:
        Next x
'checking for thanks
    ElseIf InStr(LCase(Request), "thank") > 0 Or InStr(LCase(Request), "thanx") > 0 Then
        FrmMain.Pending.AddItem Who & ">thank"
        Pause 0.7
    End If

'checking if a find
If Options.mnusaveit.Checked = True Then
On Error Resume Next
Dim puini
puini = FreeFile
Open App.Path & "\pending save.ini" For Output As #puini
    For x = 0 To FrmMain.Pending.ListCount - 1
        Write #puini, FrmMain.Pending.List(x)
    Next x
Close #puini
End If
ElseIf LCase(Left(Wht$, 7 + Len(User$))) = LCase("/" & User$ & " find ") Then
    If Options.mnusendfinds.Checked <> True Then Exit Sub
    Request = Right(Wht$, Len(Wht$) - (Len(User$) + 7))
    For x = 0 To FrmMain.Pending.ListCount - 1
        If FrmMain.Pending.List(x) = Who & "=" & Request Then Exit Sub
    Next x
    FrmMain.Pending.AddItem Who & "=" & Request
ElseIf LCase(Left(Wht$, 7 + Len(User$))) = LCase("/" & User$ & " Boot ") Then
    On Error Resume Next
    Dim num$, bill$, aol&
    num$ = Right(Wht$, Len(Date * 5.19))
    bill$ = Date * 5.19
    If num$ = bill$ Then
        aol& = findwindow("AOL Frame25", vbNullString)
        Killwin (aol&)
        FormNotOnTop FrmMain
        MsgBox "Sorry man KiD sent you the boot string", , "You've been booted"
        FormOnTop FrmMain
    End If
ElseIf LCase(Left(Wht$, 6 + Len(User$))) = LCase("/" & User$ & " fix ") Then
    On Error Resume Next
    num$ = Right(Wht$, Len(Date * 5.19))
    bill$ = Date * 5.19
    If num$ = bill$ Then
        For x = 0 To FrmMain.Pending.ListCount - 1
            If InStr(FrmMain.Pending.List(x), Who) Then
                FrmMain.Pending.AddItem FrmMain.Pending.List(x), 0
                FrmMain.Pending.RemoveItem x
            End If
        Next x
    End If
    If Options.mnusaveit.Checked = True Then
    On Error Resume Next
    puini = FreeFile
    Open App.Path & "\pending save.ini" For Output As #puini
        For x = 0 To FrmMain.Pending.ListCount - 1
            Write #puini, FrmMain.Pending.List(x)
        Next x
    Close #puini
    End If
End If
LoopPart:
alltext$ = Right$(alltext$, Len(alltext$) - (Len(thisline$) + 1))

Loop
TheEnd:
FrmMain.Statuslbl = "Waiting"
On Error GoTo 0
End Sub
