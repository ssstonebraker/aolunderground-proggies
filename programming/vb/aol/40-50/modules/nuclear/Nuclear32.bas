Attribute VB_Name = "Nuclear32"
'Nuclear32.bas by Hider
'E-mail: t3t@hider.com
'for AOL 4.0
'any questions go to www.hider.com
'and post on my board
'more to be added later this is version1
Option Explicit
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ChangeMenu Lib "user32" Alias "ChangeMenuA" (ByVal hMenu As Long, ByVal cmd As Long, ByVal lpszNewItem As String, ByVal cmdInsert As Long, ByVal FLAGS As Long) As Long
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetCursor Lib "user32" () As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "user32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Const GWL_WNDPROC = (-4)
Public Const WM_NCDESTROY = &H82

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
Public Const PROCESS_VM_READ = &H10

Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_SETCURSOR = &H20

Global Const CC_RGBINIT = &H1&
Global Const CF_BOTH = &H3&

Private Type HOOKINFO
    hwnd As Long
    OldWndProc As Long
End Type

Private HookArray() As HOOKINFO
Private NumHooks As Integer

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
Public Sub Add_Bud_ToList(Who As String)
    Dim ao As Long, mdi As Long, bList As Long
    Dim bIcon As Long, gEtit As Integer
    Dim edit As Long, iCon2 As Long, gEtit2 As Long
    Dim eDit2 As Long, eDitri As Long
    Dim eIcon As Long, eIcon2 As Long, oKw As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    bList& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    bIcon& = FindWindowEx(bList&, 0&, "_AOL_Icon", vbNullString)
    For gEtit% = 1 To 4
        bIcon& = GetWindow(bIcon&, 2)
    Next gEtit%
        TimeOut (0.05)
        Icon bIcon&
        TimeOut (2)
        edit& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
        iCon2& = FindWindowEx(edit&, 0&, "_AOL_Icon", vbNullString)
        iCon2& = FindWindowEx(edit&, iCon2&, "_AOL_Icon", vbNullString)
        TimeOut (0.01)
        Icon iCon2&
        Closer edit&
        TimeOut (3)
        eDit2& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List Buddies")
        gEtit2& = FindWindowEx(eDit2&, 0&, "_AOL_Edit", vbNullString)
        gEtit2& = FindWindowEx(eDit2&, gEtit2&, "_AOL_Edit", vbNullString)
        Call SendMessageByString(gEtit2&, WM_SETTEXT, 0&, Who$)
        TimeOut (1)
        eDitri& = FindWindowEx(eDit2&, 0&, "_AOL_Icon", vbNullString)
        Icon eDitri&
        eIcon& = FindWindowEx(eDit2&, eDitri&, "_AOL_Icon", vbNullString)
        eIcon2& = FindWindowEx(eDit2&, eIcon&, "_AOL_Icon", vbNullString)
        TimeOut (1)
        Icon eIcon2&
        TimeOut (2)
        oKw& = FindWindow("#32770", "America Online")
        Closer oKw&
'This is to add a buddy to your bl
End Sub
Public Sub AddRoomToCombo(ListBox As ListBox, ComboBox As ComboBox)
    Dim Add As Integer
    ComboBox.Clear
    For Add% = 0 To ListBox.ListCount
        ComboBox.AddItem (ListBox.List(Add%))
    Next Add%
End Sub
Public Sub AddRoomToList(ListBox As ListBox)
    On Error Resume Next
    ListBox.Clear
    Dim AOLProcess As Long, lItem As Long, Person As String, lPerson As Long
    Dim ReadBytes As Long, Room As Long, aoHandle As Long, aoThread As Long
    Dim aoProThread As Long, index As Integer
    Room& = InRoom&
    aoHandle& = FindWindowEx(Room&, 0&, "_AOL_ListBox", vbNullString)
    aoThread& = GetWindowThreadProcessId(aoHandle&, AOLProcess)
    aoProThread& = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
    If aoProThread& Then
       For index% = 0 To SendMessage(aoHandle&, LB_GETCOUNT, 0, 0) - 1
           Person$ = String$(4, vbNullChar)
           lItem& = SendMessage(aoHandle&, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
           lItem& = lItem& + 24
           Call ReadProcessMemory(aoProThread&, lItem&, Person$, 4, ReadBytes&)
           Call CopyMemory(lPerson&, ByVal Person$, 4)
           lPerson& = lPerson + 6
           Person$ = String$(16, vbNullChar)
           Call ReadProcessMemory(aoProThread&, lPerson&, Person$, Len(Person$), ReadBytes&)
           ListBox.AddItem Person$
       Next index%
           Call CloseHandle(aoProThread&)
    End If
'This is pretty much the same in every bas
End Sub
Public Function ChangeCaption(hwnd As Long, Caption As String) As Long
    Call SendMessageByString(hwnd&, WM_SETTEXT, 0&, Caption$)
'change the caption of a window
End Function
Public Function Chat(What_Say As String) As String
    Dim cRoom As Long, aoR As Long, aor2 As Long
    cRoom& = InRoom&
        aoR& = FindWindowEx(cRoom&, 0&, "RICHCNTL", vbNullString)
        aor2& = FindWindowEx(cRoom&, aoR&, "RICHCNTL", vbNullString)
        Call SendMessageByString(aor2&, WM_SETTEXT, 0&, What_Say$)
        Call SendMessageLong(aor2&, WM_CHAR, ENTER_KEY, 0&)
'This is the chat sender
End Function
Public Function Chat2(What_to_Say As String) As String
    Dim ao As Long, md As Long, aoC As Long, aoR As Long, aor2 As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    md& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    aoC& = FindWindowEx(md&, 0&, "AOL Child", vbNullString)
    aoR& = FindWindowEx(aoC&, 0&, "RICHCNTL", vbNullString)
    aor2& = FindWindowEx(aoC&, aoR&, "RICHCNTL", vbNullString)
    SetFocus (aor2&)
    Call SendMessageByString(aor2&, WM_SETTEXT, 0&, What_to_Say$)
    Call SendMessageLong(aor2&, WM_CHAR, ENTER_KEY, 0&)
'This is an extra chat send
End Function
Public Function ChatLine() As String
    Dim cText As String, cTnum As Long
    Dim cTrim As String
    cText$ = ChatLineWithSN
    cTnum& = Len(ChatLineSN)
    cTrim$ = Mid$(cText$, cTnum& + 4, Len(cText$) - Len(ChatLineSN))
    ChatLine$ = cTrim$
'gets the last line of chat without the sn
End Function
Public Function ChatLineSN() As String
    Dim cText As String, cTrim As String
    Dim gEtit As Long, tSN As String
    cText$ = ChatLineWithSN
    cTrim$ = Left$(cText$, 11)
    For gEtit& = 1 To 11
        If Mid$(cTrim$, gEtit&, 1) = ":" Then
            tSN$ = Left$(cTrim$, gEtit& - 1)
        End If
    Next gEtit&
    ChatLineSN$ = tSN$
'gets the sn from the last line of chat
End Function
Public Function ChatLineWithSN() As String
    Dim cText As String, fChar As Long, tCTxt As String
    Dim tChar As String, tChars As String
    Dim lChat As Long, cLast As String
    cText$ = ChatText
    For fChar& = 1 To Len(cText$)
        tChar$ = Mid(cText$, fChar&, 1)
        tChars$ = tChars$ & tChar$
        If tChar$ = Chr(13) Then
            tCTxt$ = Mid(tChars$, 1, Len(tChars$) - 1)
            tChars$ = ""
        End If
    Next fChar&
    lChat& = Val(fChar&) - Len(tChars$)
    cLast$ = Mid(cText$, lChat&, Len(tChars$))
    ChatLineWithSN$ = cLast$
'gets the lastline of chat with the sn
End Function
Public Function ChatText() As String
    Dim rO As Long, aoR As Long, cText As String
    rO& = InRoom
    aoR& = FindWindowEx(rO&, 0&, "RICHCNTL", vbNullString)
    cText$ = GetText(aoR&)
    ChatText$ = cText$
'gets all the chat text in a room
End Function
Public Sub Closer(hwnd As Long)
    Call SendMessageLong(hwnd&, WM_CLOSE, 0&, 0&)
'closes the window you want to close
End Sub
Public Function InRoom() As Long
    Dim ao As Long, md As Long, aoC As Long
    Dim aoR As Long, aol As Long
    Dim aoI As Long, aoS As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    md& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    aoC& = FindWindowEx(md&, 0&, "AOL Child", vbNullString)
    aoR& = FindWindowEx(aoC&, 0&, "RICHCNTL", vbNullString)
    aol& = FindWindowEx(aoC&, 0&, "_AOL_Listbox", vbNullString)
    aoI& = FindWindowEx(aoC&, 0&, "_AOL_Icon", vbNullString)
    aoS& = FindWindowEx(aoC&, 0&, "_AOL_Static", vbNullString)
    If aoR& <> 0& And aol& <> 0& And aoI& <> 0& And aoS& <> 0& Then
        InRoom& = aoC&
'The function stopped here but I found that
'it only found the room if no other windows
'were open.I tried to get focus and for/next
'to get the window when other windows were
'open.But that didn't work so I did a loop
'and it seems to work ok.Also had trouble
'with the im if other windows were open.
        Exit Function
    Else
        Do
            aoC& = FindWindowEx(md&, aoC&, "AOL Child", vbNullString)
            aoR& = FindWindowEx(aoC&, 0&, "RICHCNTL", vbNullString)
            aol& = FindWindowEx(aoC&, 0&, "_AOL_Listbox", vbNullString)
            aoI& = FindWindowEx(aoC&, 0&, "_AOL_Icon", vbNullString)
            aoS& = FindWindowEx(aoC&, 0&, "_AOL_Static", vbNullString)
            If aoR& <> 0& And aol& <> 0& And aoI& <> 0& And aoS& <> 0& Then
                InRoom& = aoC&
            End If
            Exit Function
        Loop Until InRoom& = aoC&
    End If
    InRoom& = aoC&
'This is find room
End Function
Public Function IsImOpen() As Long
    Dim ao As Long, mdi As Long, mi As Long, cap As String
    ao& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    mi& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    cap$ = GetCaption(mi&)
    If InStr(cap$, "Instant Message") = 1 Or InStr(cap$, "Instant Message") = 2 Or InStr(cap$, "Instant Message") = 3 Then
        IsImOpen& = mi&
'This function stopped here but I also
'had trouble setting focus to the im when
'other windows were open so I added the code below
        Exit Function
    Else
        Do
            mi& = FindWindowEx(mdi&, mi&, "AOL Child", vbNullString)
            cap$ = GetCaption(mi&)
            If InStr(cap$, "Instant Message") = 1 Or InStr(cap$, "Instant Message") = 2 Or InStr(cap$, "Instant Message") = 3 Then
                IsImOpen& = mi&
                Exit Function
            End If
        Loop Until mi& = 0&
    End If
    IsImOpen& = mi&
'This was added because I had trouble
'getting focus on the im window when
'other windows were open
End Function
Public Function ChangeAOLRoomCaption(New_Caption As String) As String
    Dim ao As Long, md As Long, aoC As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    md& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    aoC& = FindWindowEx(md&, 0&, "AOL Child", vbNullString)
    ChangeCaption aoC&, New_Caption$
'This will change AOL room caption in the chat room
End Function
Public Sub Button(aoBut As Long)
    Dim but As Long
    but& = SendMessage(aoBut&, WM_KEYDOWN, VK_SPACE, 0&)
    but& = SendMessage(aoBut&, WM_KEYUP, VK_SPACE, 0&)
'Cliacks the aol buttons
End Sub
Public Function ChangeAOLCaption(newcaption As String) As String
    Dim cap As Long
    cap& = FindWindow("AOL Frame25", vbNullString)
    ChangeCaption cap&, newcaption$
'This will change AOL caption
End Function
Public Sub IM_BuddyByIndex(Who As Long, What_to_Say As String)
    Dim ao As Long, mdi As Long, aoC As Long
    Dim bList As Long, gEtit As Integer, bIcon As Long
    Dim blBox As Long, index As Long, iMwin As Long
    Dim imR As Long, gIcon As Integer
    ao& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    bList& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    blBox& = FindWindowEx(bList&, 0&, "_AOL_Listbox", vbNullString)
    index& = SendMessage(blBox&, LB_GETCOUNT, 0&, 0&)
    Call SendMessage(blBox&, LB_SETCURSEL, Who&, 0&)
    Call PostMessage(blBox&, WM_LBUTTONDBLCLK, 0&, 0&)
    TimeOut (2)
    iMwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
    imR& = FindWindowEx(iMwin&, 0&, "RICHCNTL", vbNullString)
    Call SendMessageByString(imR&, WM_SETTEXT, 0&, What_to_Say$)
    bIcon& = FindWindowEx(iMwin&, 0&, "_AOL_Icon", vbNullString)
    For gIcon% = 1 To 9
        bIcon& = GetWindow(bIcon&, 2)
    Next gIcon%
    TimeOut (0.01)
    Icon bIcon&
'This wii send an im to a buddy on you list
'call it like this  IM_BuddyByIndex index#, whatto say
End Sub
Public Sub Icon(aoIcon As Long)
    Dim aoI As Long
    aoI& = SendMessage(aoIcon&, WM_LBUTTONDOWN, 0&, 0&)
    aoI& = SendMessage(aoIcon&, WM_LBUTTONUP, 0&, 0&)
'clicks the aol icons
End Sub
Public Function ImSn() As String
    Dim ao&, mdi&, mi&, imCa As String
    Dim tSN As String, cap As String
    ao& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    mi& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    SetFocus (mi&)
    cap$ = GetCaption(mi&)
       If InStr(cap$, "Instant Message From:") = 1 Then GoTo Finish
Finish:
       imCa$ = GetCaption(mi&)
       tSN$ = Mid(imCa$, InStr(imCa$, ":") + 1)
       ImSn$ = tSN$
'This gets the sn from the im
End Function
Public Sub HideWindow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
'Hides the window you want
End Sub
Public Sub SeeWindow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
'shows a hidden window
End Sub
Public Function GetText(hwnd As Long) As String
    Dim gTrim As Long, tSpace As String, geString As Long
    gTrim& = SendMessageByNum(hwnd&, 14, 0&, 0&)
    tSpace$ = Space$(gTrim&)
    geString& = SendMessageByString(hwnd&, 13, gTrim& + 1, tSpace$)
    GetText$ = tSpace$
'gets text of the window you want
End Function
Public Function GetCaption(hwnd As Long) As String
    Dim hLenght As Long, hTitle As String, a As Long
    hLenght& = GetWindowTextLength(hwnd&)
    hTitle$ = String$(hLenght&, 0)
    a& = GetWindowText(hwnd&, hTitle$, (hLenght& + 1))
    GetCaption$ = hTitle$
'Gets the caption of the window you want
End Function
Public Sub HideWelcomeWindow()
    Dim aol As Long, wel As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    wel& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    HideWindow (wel&)
'Hides the aol welcome window
End Sub
Public Sub ShowWelcomeWindow()
    Dim aol As Long, wel As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    wel& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    SeeWindow (wel&)
'Shows the aol welcome window after you hide it
End Sub
Public Sub Sign_Off(YouSure As Boolean)
    Dim ao As Long, Off As String
    Dim Respond As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    Off$ = "Sign Off"
    Respond& = MsgBox("Exit AOL Now?", vbYesNo, "Nuclear32.bas")
    If Respond = vbYes Then
        YouSure = True
        Call Run_MenuByString(ao&, Off$)
    Else
        YouSure = False
        Exit Sub
    End If
'Shuts down aol but will ask first
End Sub
Public Sub KeyWord(KeyWord As String)
    Dim ao As Long, aoT As Long, aoed As Long
    Dim ao2 As Long, aoT2 As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    aoT& = FindWindowEx(ao&, 0&, "AOL Toolbar", vbNullString)
    aoT2& = FindWindowEx(aoT&, 0&, "_AOL_Toolbar", vbNullString)
    ao2& = FindWindowEx(aoT2&, 0&, "_AOL_Combobox", vbNullString)
    aoed& = FindWindowEx(ao2&, 0&, "Edit", vbNullString)
    Call SendMessageByString(aoed&, WM_SETTEXT, 0&, KeyWord$)
    Call SendMessageLong(aoed&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(aoed&, WM_CHAR, ENTER_KEY, 0&)
'This keyword sub uses the toolbar on AOL
End Sub
Public Sub KeyWord2(KeyWord As String)
    Dim ao As Long, aoT As Long, aoT2 As Long
    Dim aoI As Long, mdi As Long, GI As Integer
    Dim kwWin As Long, aoE As Long, aoI2 As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    aoT& = FindWindowEx(ao&, 0&, "AOL Toolbar", vbNullString)
    aoT2& = FindWindowEx(aoT&, 0&, "_AOL_Toolbar", vbNullString)
    aoI& = FindWindowEx(aoT2&, 0&, "_AOL_Icon", vbNullString)
    For GI% = 1 To 20
        aoI& = GetWindow(aoI&, 2)
    Next GI%
    Icon (aoI&)
         Do: DoEvents
             mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
             kwWin& = FindWindowEx(mdi&, 0&, "AOL Child", "Keyword")
             aoE& = FindWindowEx(kwWin&, 0&, "_AOL_Edit", vbNullString)
             aoI2& = FindWindowEx(kwWin&, 0&, "_AOL_Icon", vbNullString)
         Loop Until kwWin& <> 0& And aoE& <> 0& And aoI2 <> 0&
             Call SendMessageByString(aoE&, WM_SETTEXT, 0&, KeyWord$)
             Call TimeOut(0.05)
             Icon (aoI2&)
             Icon (aoI2&)
'This uses the aol toolbar to show the keyword box
'I use this in the im sub keyword can be used also
'The other keyword sub would be better for a buster
End Sub
Public Sub MailSender(Who As String, subject As String, message As String)
    Dim aol As Long, aoT As Long, aoT2 As Long
    Dim aoI As Long, mdi As Long, aoMa As Long
    Dim aoE As Long, aoR As Long, aoI2 As Long
    Dim gIcon As Integer, aoEr As Long
    Dim aoMo As Long, aoi3 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    aoT& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    aoT2& = FindWindowEx(aoT&, 0&, "_AOL_Toolbar", vbNullString)
    aoI& = FindWindowEx(aoT2&, 0&, "_AOL_Icon", vbNullString)
    aoI& = FindWindowEx(aoT2&, aoI&, "_AOL_Icon", vbNullString)
    Icon aoI&
    TimeOut (1#)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    aoMa& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
    aoE& = FindWindowEx(aoMa&, 0&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(aoE&, WM_SETTEXT, 0&, Who$)
    aoE& = FindWindowEx(aoMa&, aoE&, "_AOL_Edit", vbNullString)
    aoE& = FindWindowEx(aoMa&, aoE&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(aoE&, WM_SETTEXT, 0&, subject$)
    aoR& = FindWindowEx(aoMa&, 0&, "RICHCNTL", vbNullString)
    Call SendMessageByString(aoR&, WM_SETTEXT, 0&, message$)
    aoI2& = FindWindowEx(aoMa&, 0&, "_AOL_Icon", vbNullString)
    For gIcon% = 1 To 18
        aoI2& = GetWindow(aoI2&, 2)
    Next gIcon%
    Icon aoI2&
    TimeOut (2#)
    aoMo& = FindWindowEx(aol&, 0&, "_AOL_Modal", vbNullString)
    aoEr& = FindWindowEx(aoMo&, 0&, "_AOL_Modal", vbNullString)
    aoi3& = FindWindowEx(aoEr&, 0&, "_AOL_Icon", vbNullString)
    Icon aoi3&
'uses the toolbar mail icon to send mail
End Sub
Public Sub PR_Bust(RoomName As String)
    Call KeyWord("aol://2719:2-2-" & RoomName$)
    Wait_For_OK
    If InRoom& Then
        Exit Sub
    Else
        Do
            DoEvents
            Call KeyWord("aol://2719:2-2-" & RoomName$)
            Wait_For_OK
        Loop Until InRoom&
        Exit Sub
    End If
'bust in a private room
End Sub
Public Sub Member_Bust(RoomName As String)
    Call KeyWord("aol://2719:21-2-" & RoomName$)
    Wait_For_OK
    If InRoom& Then
        Exit Sub
    Else
        Do
            DoEvents
            Call KeyWord("aol://2719:21-2-" & RoomName$)
            Wait_For_OK
        Loop Until InRoom&
        Exit Sub
    End If
'Busts in a member room
End Sub
Public Sub Run_Menu(aMenu As Long, bMenu As Long)
    Dim aol As Long, Mnu As Long, SubMnu As Long
    Dim rMnu As Long, MnuID As Long, Clicker As Long
    Static DoIt As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    Mnu& = GetMenu(aol&)
    SubMnu& = GetSubMenu(Mnu&, aMenu&)
    MnuID& = GetMenuItemID(SubMnu&, bMenu&)
    rMnu& = CLng(0&) * &H10000 Or DoIt
    Clicker& = SendMessageByNum(aol&, 273, MnuID&, 0&)
End Sub
Public Sub Run_MenuByString(App As Long, sString As String)
    Dim tSearch As Long, mnuCount As Integer
    Dim fString As Integer, theSearch As Long
    Dim itmCount As Long, getStr As Integer
    Dim sCount As Long, buffer As String
    Dim strMnu As Long, mnuItem As Long
    Dim rIt As Long
    tSearch& = GetMenu(App&)
    mnuCount% = GetMenuItemCount(tSearch&)
    For fString% = 0 To mnuCount% - 1
        theSearch& = GetSubMenu(tSearch&, fString%)
        itmCount& = GetMenuItemCount(theSearch&)
        For getStr% = 0 To itmCount& - 1
            sCount& = GetMenuItemID(theSearch&, getStr%)
            buffer$ = String$(100, " ")
            strMnu& = GetMenuString(tSearch&, sCount&, buffer$, 100, 1)
            If InStr(UCase(buffer$), UCase(sString$)) Then
                mnuItem& = sCount&
                GoTo Same
           End If
        Next getStr%
    Next fString%
Same:
    rIt& = SendMessageLong(App&, WM_COMMAND, mnuItem&, 0&)
End Sub
Public Sub TimeOut(HowLong As Long)
    Dim sTime As Long
    sTime& = Timer
    Do While Timer - sTime& < HowLong&
        DoEvents
    Loop
End Sub
Public Sub IM(Who As String, What As String)
    Dim ao As Long, mdi As Long, iMwin As Long
    Dim aoE As Long, aoR As Long, aoIc As Long
    Dim X As Integer, oKw As Long, closer2 As Long
    ao& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    Call KeyWord2("aol://9293:")
    Do: DoEvents
        iMwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
        aoE& = FindWindowEx(iMwin&, 0&, "_AOL_Edit", vbNullString)
        aoR& = FindWindowEx(iMwin&, 0&, "RICHCNTL", vbNullString)
        aoIc& = FindWindowEx(iMwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until aoE& <> 0& And aoR& <> 0& And aoIc& <> 0&
        Call SendMessageByString(aoE&, WM_SETTEXT, 0&, Who$)
        Call SendMessageByString(aoR&, WM_SETTEXT, 0&, What$)
    For X% = 1 To 9
        aoIc& = GetWindow(aoIc&, 2)
    Next X%
        Call TimeOut(0.01)
        Icon (aoIc&)
    Do: DoEvents
        ao& = FindWindow("AOL Frame25", vbNullString)
        mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
        iMwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
        oKw& = FindWindow("#32770", "America Online")
        If oKw& <> 0& Then Closer oKw&
        Closer iMwin&
        Exit Do
        If iMwin& = 0& Then Exit Do
    Loop
'This uses the im keyword box
End Sub
Public Sub IM_On()
    IM "$IM_ON", "HEHE"
'turns ims on if there off
End Sub
Public Sub IM_Off()
    IM "$IM_OFF", "HOHO"
'turns im's off
End Sub
Public Sub IM2(Who As String, WhatSay As String)
    Dim ao As Long, mdi As Long, bList As Long
    Dim bIcon As Long, sWin As Long, sIcon As Long
    Dim Xoom As Integer, sEdit As Long, sRich As Long
    Dim Boom As Integer
    ao& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    bList& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    bIcon& = FindWindowEx(bList&, 0&, "_AOL_Icon", vbNullString)
    For Xoom% = 1 To 2
        bIcon& = GetWindow(bIcon&, 2)
    Next Xoom%
    Icon bIcon&
    TimeOut (2)
    sWin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
    sEdit& = FindWindowEx(sWin&, 0&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(sEdit&, WM_SETTEXT, 0&, Who$)
    sRich& = FindWindowEx(sWin&, 0&, "RICHCNTL", vbNullString)
    Call SendMessageByString(sRich&, WM_SETTEXT, 0&, WhatSay$)
    sIcon& = FindWindowEx(sWin&, 0&, "_AOL_Icon", vbNullString)
    For Boom% = 1 To 9
        sIcon& = GetWindow(sIcon&, 2)
    Next Boom%
    Icon sIcon&
'This uses the buddylist to send an IM
'it works even if the list is minimized
End Sub
Public Function SN() As String
    On Error Resume Next
    Dim ao As Long, mdi As Long, welC As Long
    Dim wLength As Long, wTitle As String, gTex As Long
    Dim Who As String
    ao& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(ao&, 0&, "MDIClient", vbNullString)
    welC& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    wLength& = GetWindowTextLength(welC&)
    wTitle$ = String$(200, 0)
    gTex& = GetWindowText(welC&, wTitle$, (wLength& + 1))
    Who$ = Mid$(wTitle$, 10, (InStr(wTitle$, "!") - 10))
    SN$ = Who$
'This gets the user SN from the welcome screen
End Function
Public Function IMLastMessageWithSN() As String
    Dim fIm As Long, aoR As Long, iText As String, tline As String
    Dim tLines As String, tIMtxt As String
    Dim fText As Long, lLen As Long, Msg As String
    fIm& = IsImOpen&
    aoR& = FindWindowEx(fIm&, 0&, "RICHCNTL", vbNullString)
    iText$ = GetText(aoR&)
    For fText& = 1 To Len(iText$)
         tline$ = Mid(iText$, fText&, 1)
         tLines$ = tLines$ & tline$
         If tline$ = Chr(13) Then
            tLines$ = ""
         End If
    Next fText&
    lLen& = Val(fText&) - Len(tLines$)
    Msg$ = Mid(iText$, lLen&, Len(tLines$))
    IMLastMessageWithSN$ = Msg$
'This gets the last message from the im with sn
End Function
Public Function IMLastMessage() As String
    Dim iText As String, iTrimnum As Long
    Dim iTrim As String
    iText$ = IMLastMessageWithSN
    iTrimnum& = Len(ImSn)
    iTrim$ = Mid$(iText$, iTrimnum + 4, Len(iText$) - Len(ImSn))
    IMLastMessage$ = iTrim$
'This gets the last im message
End Function
Public Sub Wait_For_OK()
    Dim okWin As Long, yes As Long, status As String
    Do
        okWin& = FindWindow("#32770", "America Online")
        If status$ = "Off" Then
            Exit Sub
            Exit Do
        End If
        DoEvents
    Loop Until okWin& <> 0&
        yes& = FindWindowEx(okWin&, 0&, "Button", "OK")
        Button yes&
'wait for the aol ok window before proceeding
'with a sub
End Sub
Public Sub Chat_Wavy(WhatSay As String)
    Dim Boom As String, a As Integer
    Dim Zoom As Integer, H As String
    Dim i As String, d As String, e As String
    Dim r As String
    Boom$ = WhatSay$
    a% = Len(Boom$)
    For Zoom% = 1 To a% Step 4
        H$ = Mid$(Boom$, Zoom%, 1)
        i$ = Mid$(Boom$, Zoom% + 1, 1)
        d$ = Mid$(Boom$, Zoom% + 2, 1)
        e$ = Mid$(Boom$, Zoom% + 3, 1)
        r$ = r$ & "<sup>" & H$ & "</sup>" & i$ & "<sub>" & d$ & "</sub>" & e$
    Next Zoom%
    Chat r$
'Sends wavy chat to a room
End Sub
'Here starts the form effects section
'of the bas.This is something I like
'to do.It's cool to make a form have
'different effects.
Public Sub Form_Explode(Form As Form, Movement As Integer)
'Call this in form load or unload
    Dim myRect As RECT
    Dim formWidth As Integer, formHeight As Integer, i As Integer
    Dim X As Integer, Y As Integer
    Dim cx As Integer, cy As Integer
    Dim TheScreen As Long, Brush As Long
    GetWindowRect Form.hwnd, myRect
        formWidth% = (myRect.Right - myRect.Left)
        formHeight% = myRect.Bottom - myRect.Top
            TheScreen& = GetDC(0)
            Brush& = CreateSolidBrush(Form.BackColor)
            For i% = 1 To Movement%
                cx% = formWidth * (i% / Movement%)
                cy% = formHeight * (i% / Movement%)
                X% = myRect.Left + (formWidth% - cx%) / 2
                Y = myRect.Top + (formHeight% - cy%) / 2
                Rectangle TheScreen, X%, Y%, X% + cx%, Y% + cy%
            Next i%
                X% = ReleaseDC(0, TheScreen&)
                DeleteObject (Brush&)
End Sub
Public Sub Form_Implode(Form As Form, Movement As Integer)
'Call this in form load or unload
'Example ImplodeForm FormName,1000
'the bigger the interval the more
'effect you get
    Dim myRect As RECT
    Dim formWidth As Integer, formHeight As Integer
    Dim i As Integer, X As Integer
    Dim Y As Integer, cx As Integer, cy As Integer
    Dim TheScreen As Long, Brush As Long
    GetWindowRect Form.hwnd, myRect
        formWidth% = (myRect.Right - myRect.Left)
        formHeight% = myRect.Bottom - myRect.Top
            TheScreen& = GetDC(0)
            Brush& = CreateSolidBrush(Form.BackColor)
            For i% = Movement% To 1 Step -1
                cx% = formWidth% * (i% / Movement%)
                cy% = formHeight% * (i% / Movement%)
                X% = myRect.Left + (formWidth% - cx%) / 2
                Y% = myRect.Top + (formHeight% - cy%) / 2
                Rectangle TheScreen&, X%, Y%, X% + cx%, Y% + cy%
            Next i%
                X% = ReleaseDC(0, TheScreen&)
                DeleteObject (Brush&)
End Sub
Public Function Form_FlashTitleBar(Flash As Form, Howmany As Integer, HowLong As Single) As Form
    Dim i As Integer, start As Single
    For i% = 0 To Howmany%
        Call FlashWindow(Flash.hwnd, True)
        start = Timer
        Do While Timer < start = HowLong
            DoEvents
        Loop
    Next i%
    Call FlashWindow(Flash.hwnd, False)
'uses flashwindow api to flash the title bar
End Function
Public Sub Form_CoolExit(Form As Form)
    Dim sStart As Integer, GoNow As Long
    GoNow& = Form.Height / 2
    For sStart% = 1 To GoNow&
    DoEvents
        Form.Height = Form.Height - 10
        Form.Top = (Screen.Height - Form.Height) \ 2
        If Form.Height <= 11 Then GoTo Finish
    Next sStart%
Finish:
        Form.Height = 30
        GoNow& = Form.Width / 2
    For sStart% = 1 To GoNow&
    DoEvents
        Form.Width = Form.Width - 10
        Form.Left = (Screen.Width - Form.Width) \ 2
        If Form.Width <= 11 Then Exit Sub
    Next sStart%
    Unload Form
'this is the effect I like the best
End Sub
Public Sub Form_ExitDown(Form As Form)
    Do
        Form.Top = Trim(Str(Int(Form.Top) + 300))
        DoEvents
    Loop Until Form.Top > 10000
        If Form.Top > 10000 Then End
End Sub
Public Sub Form_ExitLeft(Form As Form)
    Do
        Form.Left = Trim(Str(Int(Form.Left) - 300))
        DoEvents
    Loop Until Form.Left < -6300
    If Form.Left < -6300 Then End
End Sub
Public Sub Form_ExitRight(Form As Form)
    Do
        Form.Left = Trim(Str(Int(Form.Left) + 300))
        DoEvents
    Loop Until Form.Left > 11000
    If Form.Left > 11000 Then End
End Sub
Public Sub Form_ExitUp(Form As Form)
    Do
        Form.Top = Trim(Str(Int(Form.Top) - 300))
        DoEvents
    Loop Until Form.Top < -4500
    If Form.Top < -4500 Then End
End Sub
Public Sub Form_Center(Form As Form)
    Form.Top = (Screen.Height * 0.85) / 2 - Form.Height / 2
    Form.Left = Screen.Width / 2 - Form.Width / 2
End Sub
Public Sub Form_OnTop(Form As Form)
    Form& = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub Form_MoveWithNoTitle(Form As Form)
    Call ReleaseCapture
    Call SendMessage(Form.hwnd, &H112, &HF012, 0)
End Sub
'More to be added soon

