Attribute VB_Name = "Shaggy"
' Some code was made using code
' from DoS32.bas, Caloric, Source
' Myst, and DaCrazyOne
' All Credit Is Worth Noteing
' by the way just thought it
' should be known that i made
' the getchatetext and msg and sn
' subs all because everyone else
' was trying to find the richcntl
' and its actually richcntlreadonly
' thanks -Shaggy
Option Explicit

Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const WM_CLEAR = &H303
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
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
Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E
Public Const GW_HWNDFIRST = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Const SW_SHOWNORMAL = 1
Public Const Op_Flags = PROCESS_READ Or RIGHTS_REQUIRED
Public Const SW_RESTORE = 9
Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCURSEL& = &H188
Public Const LB_INSERTSTRING = &H181
Public Const LB_RESETCONTENT& = &H184
Public Const CB_ADDSTRING& = &H143
Public Const CB_DELETESTRING& = &H144
Public Const CB_FINDSTRINGEXACT& = &H158
Public Const CB_GETITEMDATA = &H150
Public Const CB_RESETCONTENT& = &H14B
Global Const SND_SYNC = &H0
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Public Const Sys_Add = &H0
Public Const Sys_Delete = &H2
Public Const Sys_Message = &H1
Public Const Sys_Icon = &H2
Public Const Sys_Tip = &H4
Public Const Snd_Flag2 = SND_ASYNC Or SND_LOOP
Public Const WM_MOUSEMOVE = &H200
Public Const MF_BYPOSITION = &H400&
Public Const EM_GETLINECOUNT& = &HBA

Public Enum MAILTYPE
        mailFLASH
        mailNEW
        mailOLD
        mailSENT
End Enum

Public systray As NOTIFYICONDATA

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type

Public Enum OnScreen
    scon
    scoff
End Enum
Public Declare Function FindParent& Lib "user32" Alias "FindWindowA" (ByVal lpClassName$, ByVal lpWindowName$)
Public Declare Function FindChild& Lib "user32" Alias "FindWindowExA" (ByVal hWnd1&, ByVal hWnd2&, ByVal lpsz1$, ByVal lpsz2$)
Public Declare Function osQueryPerformanceCounter Lib "kernel32" Alias "QueryPerformance" (lpPerformanceCount As Currency) As Long
Public Declare Function osQueryPerformanceFrequency Lib "kernel32" Alias "QueryFrequency" (lpFrequency As Currency) As Long
Public Declare Function SendIt& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Public Declare Function SenditByString& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$)
Global Lines&
Global NewLineCount&



Private Function FindRoom() As Long
    Dim aol As Long, mdi As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
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

Public Sub AddRoomToList(TheList As ListBox, AddUser As Boolean)
' Only use this sub if you know that you it's AOL 6
    On Error Resume Next
    Dim cprocess As Long, itmhold As Long, ScreenName As String
    Dim psnhold As Long, rbytes As Long, Index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = GetWindowThreadProcessId(rlist, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0 To SendMessage(rlist, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            'this was the orginal line for AOL 5: itmHold& = itmHold& + 24
            itmhold& = itmhold& + 28
            Call ReadProcessMemory(mthread&, itmhold&, ScreenName$, 4, rbytes)
            Call CopyMemory(psnhold&, ByVal ScreenName$, 4)
            psnhold& = psnhold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, ScreenName$, Len(ScreenName$), rbytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser Or AddUser = True Then
                TheList.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mthread)
    End If
End Sub

Private Sub OldAddRoomToList(TheList As ListBox, AddUser As Boolean)
' Only use this sub if you know it's AOL 5 or below..
    On Error Resume Next
    Dim cprocess As Long, itmhold As Long, ScreenName As String
    Dim psnhold As Long, rbytes As Long, Index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = GetWindowThreadProcessId(rlist, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0 To SendMessage(rlist, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, ScreenName$, 4, rbytes)
            Call CopyMemory(psnhold&, ByVal ScreenName$, 4)
            psnhold& = psnhold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, ScreenName$, Len(ScreenName$), rbytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser Or AddUser = True Then
                TheList.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mthread)
    End If
End Sub

Public Sub NewAddRoomToList(TheList As ListBox, AddUser As Boolean)
' You can use this sub with any AOL
    On Error Resume Next
    Dim cprocess As Long, itmhold As Long, ScreenName As String
    Dim psnhold As Long, rbytes As Long, Index As Long, room As Long
    Dim rlist As Long, sthread As Long, mthread As Long, itmNum As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    itmNum& = 24
Top:
    rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = GetWindowThreadProcessId(rlist, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For Index& = 0 To SendMessage(rlist, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmhold& = itmhold& + itmNum&
            Call ReadProcessMemory(mthread&, itmhold&, ScreenName$, 4, rbytes)
            Call CopyMemory(psnhold&, ByVal ScreenName$, 4)
            psnhold& = psnhold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, ScreenName$, Len(ScreenName$), rbytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            'This is used because if you use the old way it returns either
            'a blank or just a 'p'
            If Trim(ScreenName$) = "" Or Trim(ScreenName$) = "p" Then
                itmNum& = 28
                Call CloseHandle(mthread)
                GoTo Top
            End If
            If ScreenName$ <> GetUser Or AddUser = True Then
                TheList.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mthread)
    End If
End Sub



Public Function LineCount&(ByVal hwnd&)
    If hwnd& = 0& Then Let LineCount& = 0: Exit Function
    Dim findchar&
    Dim TheChar$
    Dim lineNum&
    Dim TextLength&
    Dim Text$
    
    Let Text$ = GetText$(hwnd&)
    Let TextLength& = Len(Text$)
    If TextLength& = 0 Then Exit Function
    
    For findchar& = 1 To TextLength&
        Let TheChar$ = Mid(Text$, findchar&, 1)
        If TheChar$ = Chr(13) Then lineNum& = lineNum& + 1
    Next
    
    If Mid(Text$, TextLength&, 1) = Chr(13) Then
        Let LineCount& = lineNum&
    Else
        Let LineCount& = lineNum& + 1
    End If
End Function

Public Function GetLineCount&(ByVal Text$)
    Dim findchar&
    Dim TheChar$
    Dim lineNum&
    Dim TextLength&
    
    Let TextLength& = Len(Text$)
    If TextLength& = 0 Then Exit Function
    
    For findchar& = 1 To TextLength&
        Let TheChar$ = Mid(Text$, findchar&, 1)
        If TheChar$ = Chr(13) Then lineNum& = lineNum& + 1
    Next
    
    If Mid(Text$, TextLength&, 1) = Chr(13) Then
        Let GetLineCount& = lineNum&
    Else
        Let GetLineCount& = lineNum& + 1
    End If
End Function
Public Function ReplaceText$(ByVal Text$, ByVal find$, ByVal Replace$)
    Dim FindIt&
    Dim txtBefore$
    Dim txtAfter$
    Dim txtNew$
        Let FindIt& = InStr(Text$, find$)
        If FindIt& = 0 Then Let ReplaceText$ = Text$: Exit Function
            Do
                DoEvents
                Let txtBefore$ = Left(Text$, FindIt& - 1)
                Let txtAfter$ = Mid(Text$, FindIt& + Len(find))
                Let txtNew$ = txtBefore$ & Replace$ & txtAfter$
                Let Text$ = txtNew$
                Let FindIt& = InStr(Text$, find$)
            Loop Until FindIt& = 0
    Let ReplaceText$ = Text$
End Function

Public Function LineText$(ByVal hwnd&, ByVal theline&)
    Dim findchar&
    Dim TheChar$
    Dim TheChars$
    Dim TempNum&
    Dim TheText$
    Dim TextLength&
    Dim TheCharsLength&
    Dim Text$
    
    Let Text$ = GetText$(hwnd&)
    Let TextLength& = Len(Text$)
    For findchar& = 1 To TextLength&
        Let TheChar$ = Mid$(Text$, findchar&, 1)
        Let TheChars$ = TheChars$ & TheChar$
            If TheChar$ = Chr(13) Then
                TempNum& = TempNum& + 1
                Let TheCharsLength& = Len(TheChars$)
                Let TheText$ = Mid$(TheChars$, 1, TheCharsLength& - 1)
                If theline& = TempNum& Then GoTo skipit
                Let TheChars = ""
            End If
    Next
        Let LineText$ = TheChars$
    Exit Function
skipit:
    Let TheText$ = ReplaceText$(TheText$, Chr(13), "")
    Let LineText$ = TheText$
End Function



Public Function GetText$(ByVal hwnd&)
    Dim TextLength&
    Dim NullString$
    Dim Text$
    
    Let TextLength& = SendIt&(hwnd&, WM_GETTEXTLENGTH, 0&, 0&)
    Let NullString$ = String$(TextLength&, 0&)
    Call SenditByString&(hwnd&, WM_GETTEXT, TextLength& + 1, NullString$)
    Let Text$ = NullString$
    Let GetText$ = Text$
End Function

Public Function ShorterText$(ByVal hwnd&, ByVal theline&)
    Dim findchar&
    Dim TheChar$
    Dim TheChars$
    Dim TempNum&
    Dim TheText$
    Dim TextLength&
    Dim TheCharsLength&
    Dim Text$
    Dim SumNum&
    
    Let Text$ = GetText$(hwnd&)
    Let TextLength& = Len(Text$)
    For findchar& = 1 To TextLength&
        Let TheChar$ = Mid$(Text$, findchar&, 1)
        Let TheChars$ = TheChars$ & TheChar$
            If TheChar$ = Chr(13) Then
                TempNum& = TempNum& + 1
                Let TheCharsLength& = Len(TheChars$)
                Let SumNum& = TheCharsLength& + SumNum&
                Let TheText$ = Mid$(Text$, SumNum&)
                If theline& = TempNum& Then GoTo skipit
                TheChars$ = ""
            End If
    Next
        Let ShorterText$ = TheChars$
    Exit Function
skipit:
    Let TheText$ = ReplaceText$(TheText$, Chr(13), "")
    Let ShorterText$ = TheText$
End Function

Public Function LineFromText$(ByVal Text$, ByVal theline&)

Dim findchar&
Dim TheChar$
Dim TheChars$
Dim TempNum&
Dim TheText$
Dim TextLength&
Dim TheCharsLength&

Let TextLength& = Len(Text$)
For findchar& = 1 To TextLength&
    Let TheChar$ = Mid(Text$, findchar&, 1)
    Let TheChars$ = TheChars$ & TheChar$
        If TheChar$ = Chr(13) Then
            TempNum& = TempNum& + 1
            Let TheCharsLength& = Len(TheChars$)
            Let TheText$ = Mid(TheChars$, 1, TheCharsLength& - 1)
            If theline& = TempNum& Then GoTo skipit
            Let TheChars = ""
        End If
Next
    Let LineFromText$ = TheChars$
Exit Function

skipit:
Let TheText$ = ReplaceText(TheText$, Chr(13), "")
Let LineFromText$ = TheText$

End Function



Public Sub clickToolbar(IconNumber&, Letter$)

Dim AOLFrame As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim AOLIcon As Long
Dim Count As Long
Dim found As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(clickToolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
Letter = Asc(Letter)
Call PostMessage(menu, WM_CHAR, Letter, 0&)
End Sub
Public Sub clickToolbar2(IconNumber&, Letter$, letter2$)

Dim AOLFrame As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim AOLIcon As Long
Dim Count As Long
Dim found As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(clickToolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
Letter = Asc(Letter)
letter2 = Asc(letter2)
Call PostMessage(menu, WM_CHAR, Letter, 0&)
Call PostMessage(menu, WM_CHAR, letter2, 0&)
End Sub

Public Function OpenScreenNames()
Call clickToolbar("9", "S")
End Function

Public Function OpenPrefrences()
Call clickToolbar("9", "P")
End Function

Public Function OpenMyDirectoryListing()
Call clickToolbar("9", "L")
End Function

Public Function OpenPasswords()
Call clickToolbar("9", "A")
End Function

Public Function OpenBilling()
Call clickToolbar("9", "B")
End Function
Public Function OpenParentalControls()
Call clickToolbar("9", "C")
End Function
Public Function OpenAccessNumber()
Call clickToolbar("9", "A")
End Function
Public Function OpenMYAOL()
Call clickToolbar("9", "M")
End Function
Public Function OpenAOL_Quick_Checkout()
Call clickToolbar("9", "Q")
End Function
Public Function OpenAOL_devices()
Call clickToolbar("9", "D")
End Function
Public Function OpenAOL_ACCESS_PhoneNumbers()
Call clickToolbar("9", "A")
End Function

Public Function OpenMailCenter()
Call clickToolbar("0", "M")
End Function

Public Function OpenRecentlyDeletedMail()
Call clickToolbar("0", "D")
End Function

Public Function OpenFilingCabinet()
Call clickToolbar("0", "F")
End Function

Public Function OpenMailWaiting2besent()
Call clickToolbar("0", "B")
End Function

Public Function OpenAutoAOL()
Call clickToolbar("0", "U")
End Function

Public Function OpenMailSignatures()
Call clickToolbar("0", "S")
End Function

Public Function OpenMailPrefrences()
Call clickToolbar("0", "P")
End Function
Public Function OpenWriteMail()
Call clickToolbar("0", "W")
End Function
Public Function OpenMailControls()
Call clickToolbar("0", "C")
End Function
Public Function OpenMailWaiting_toBeSent()
Call clickToolbar("0", "B")
End Function
Public Function OpenGreetings_Mail_extras()
Call clickToolbar("0", "G")
End Function
Public Function OpenNewsLetters()
Call clickToolbar("0", "N")
End Function
Public Function OpenRead_NewMail()
Call clickToolbar2("0", "R", "N")
End Function
Public Function OpenRead_OLDMail()
Call clickToolbar2("0", "R", "O")
End Function
Public Function OpenRead_SentMail()
Call clickToolbar2("0", "R", "S")
End Function

Public Function OpenChatNow()
Call clickToolbar("3", "N")
End Function
Public Function OpenSendInstantMessage()
Call clickToolbar("3", "I")
End Function
Public Function Open_Chat_PeopleConnection()
Call clickToolbar("3", "C")
End Function
Public Function OpenGetMemberProfile()
Call clickToolbar("3", "G")
End Function
Public Function OpenFindAChat()
Call clickToolbar("3", "F")
End Function
Public Function OpenCreateHomePage()
Call clickToolbar("3", "H")
End Function
Public Function OpenStartYourOwnChat()
Call clickToolbar("3", "S")
End Function
Public Function OpenJoinOnlineGroup()
Call clickToolbar("3", "J")
End Function
Public Function OpenLiveEvents()
Call clickToolbar("3", "E")
End Function
Public Function OpenSignOnAFriend()
Call clickToolbar("3", "O")
End Function
Public Function OpenBuddylist()
Call clickToolbar("3", "B")
End Function
Public Function OpenInvitations()
Call clickToolbar("3", "V")
End Function
Public Function OpenLocateMemOnline()
Call clickToolbar("3", "L")
End Function
Public Function OpenMemberDirectory()
Call clickToolbar("3", "N")
End Function
Public Function OpenMessage2Pager()
Call clickToolbar("3", "M")
End Function
Public Function OpenPersonals()
Call clickToolbar("3", "P")
End Function
Public Function OpenWhitePages()
Call clickToolbar("3", "W")
End Function

Public Function opentvkistings()
Call clickToolbar("6", "T")
End Function
Public Function openshopataol()
Call clickToolbar("6", "S")
End Function
Public Function addtocalender()
Call clickToolbar("6", "A")
End Function

Public Function OpenCalender()
Call clickToolbar("6", "C")
End Function
Public Function openCarBuying()
Call clickToolbar("6", "B")
End Function
Public Function openDownloadcenter()
Call clickToolbar("6", "D")
End Function
Public Function openHomeWorkHelp()
Call clickToolbar("6", "K")
End Function
Public Function openMapsnDirections()
Call clickToolbar("6", "M")
End Function
Public Function openGovermentGuide()
Call clickToolbar("6", "U")
End Function
Public Function openMedicalReferences()
Call clickToolbar("6", "N")
End Function
Public Function openMovieShowtimes()
Call clickToolbar("6", "W")
End Function
Public Function OpenPersonals2()
Call clickToolbar("6", "P")
End Function
Public Function openRadio()
Call clickToolbar("6", "R")
End Function
Public Function openRecipeFinder()
Call clickToolbar("6", "F")
End Function
Public Function openSportsScores()
Call clickToolbar("6", "O")
End Function
Public Function openStockPortfolios()
Call clickToolbar("6", "L")
End Function
Public Function openStockQuotes()
Call clickToolbar("6", "Q")
End Function
Public Function openTravelReservations()
Call clickToolbar("6", "V")
End Function
Public Function openYellowPages()
Call clickToolbar("6", "E")
End Function
Public Function openYouveGotPictures()
Call clickToolbar("6", "Y")
End Function

Public Function OpenFavorites()
Call clickToolbar("11", "F")
End Function

Public Function ADD_Top_Window_to_Favorites()
Call clickToolbar("11", "A")
End Function
Public Function Go_to_keyword()
Call clickToolbar("11", "G")
End Function
Public Function My_Hot_Keys()
Call clickToolbar2("11", "M", "E")
End Function

Public Sub WavyText(Text As TextBox)
Dim i%, x%, Wave$
ReDim Waver(1 To 4) As String
Waver(1) = "<sup>"
Waver(2) = "</sup>"
Waver(3) = "<sub>"
Waver(4) = "</sub>"
For i = 1 To Len(Text)
    x = x + 1
    If x > 4 Then x = 1
    Wave$ = Wave$ & Waver(x) & Mid$(Text, i, 1)
Next
Text = Wave$
End Sub
Public Sub Hex2RGB(strHexColor As String, r As Byte, g As Byte, b As Byte)
    Dim HexColor As String
    Dim i As Byte
    On Error Resume Next
    ' make sure the string is 6 characters l
    '     ong
    ' (it may have been given in &H###### fo
    '     rmat, we want ######)
    strHexColor = Right((strHexColor), 6)
    ' however, it may also have been given a
    '     s or #***** format, so add 0's in front


    For i = 1 To (6 - Len(strHexColor))
        HexColor = HexColor & "0"
    Next
    HexColor = HexColor & strHexColor
    ' convert each set of 2 characters into
    '     bytes, using vb's cbyte function
    r = CByte("&H" & Right$(HexColor, 2))
    g = CByte("&H" & Mid$(HexColor, 3, 2))
    b = CByte("&H" & Left$(HexColor, 2))
End Sub


Public Function RGB2Hex(r As Byte, g As Byte, b As Byte) As String
    On Error Resume Next
    ' convert to long using vb's rgb functio
    '     n, then use the long2rgb function
    RGB2Hex = Long2Hex(RGB(r, g, b))
End Function





Public Function RGB2Long(r As Byte, g As Byte, b As Byte) As Long
    On Error Resume Next
    ' use vb's rgb function
    RGB2Long = RGB(r, g, b)
End Function


Public Function Long2Hex(LongColor As Long) As String
    On Error Resume Next
    ' use vb's hex function
    Long2Hex = Hex(LongColor)
End Function


Public Function Hex2Long(strHexColor As String) As Long
    Dim r As Byte
    Dim g As Byte
    Dim b As Byte
    On Error Resume Next
    ' use the hex2rgb function to get the re
    '     d green and blue bytes
    Hex2RGB strHexColor, r, g, b
    ' convert to long using vb's rgb functio
    '     n
    Hex2Long = RGB(r, g, b)
End Function

Public Sub AOLSearch(Txt As String)
Dim AOLFrame As Long, AOLToolbar As Long, AOLEdit As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, 0&, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
AOLEdit = FindWindowEx(AOLToolbar, AOLEdit, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0&, Txt$)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Function SendChat(txtChat As String)
Dim RICHCNTL As Long, AOLChild As Long, MDIClient As Long
Dim AOLFrame As Long, i As Long, AOLIcon As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, txtChat$)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 5&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Function SendInstantMessage(txtSN As String, txtMSG As String)
Dim i As Long
Dim x As Long
Dim AOLIcon As Long
Dim AolToolBar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLIcon2 As Long
Dim RICHCNTL As Long

AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AolToolBar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AolToolBar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AolToolBar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
' lower pause to increase speed
' to low will make it leave the
' im blank until you call the
' instantmessage again
Pause 1
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, txtSN$)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, txtMSG$)
AOLIcon2& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For x& = 1& To 9&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next x&
Call SendMessageLong(AOLIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon2&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function Pause(Time As Long)
Dim Current As Long
Current = Timer
Do Until Timer - Current >= Time
DoEvents
Loop
End Function

Public Function FindRoomFull()

Dim x As Long
x = FindWindow("#32770", vbNullString)
Call SendMessageLong(x, WM_CLOSE, 0&, 0&)
End Function
Public Sub SendIM(Person As String, Message As String)

Dim IM&, Text&, sn&, send&, errorwin&, Count&, errorbut&
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long, AOLEdit As Long
Call Keyword7("im")
Do
DoEvents
IM& = FindWindowEx(MDIClient, 0&, "aol child", "Send Instant Message")
AOLEdit = FindWindowEx(AOLChild, 0&, "_aol_edit", "Send Instant Message")
Text& = FindWindowEx(IM&, 0&, "richcntl", "Send Instant Message")
send& = FindWindowEx(IM&, 0&, "_AOL_Icon", "Send Instant Message")
sn& = FindWindowEx(IM&, 0&, "_AOL_Edit", "Send Instant Message")
For Count& = 0 To 7
send& = FindWindowEx(IM&, send&, "_AOL_Icone", vbNullString)
Next Count&
Loop Until IM& <> 0& And send& <> 0& And Text& <> 0&
Call SendMessageByString(sn&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(Text&, WM_SETTEXT, 0&, Message$)
Call SendMessage(send&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(send&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
errorwin& = FindWindow("#32770", "America Online")
IM& = FindWindowEx(MDIClient, 0&, "aol child", "Send Instant Message")
Loop Until errorwin& <> 0 Or IM& = 0
If errorwin <> 0 Then
errorbut& = FindWindowEx(errorwin&, 0&, "Button", vbNullString)
Call PostMessage(errorbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(errorbut&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End If
End Sub
Public Sub ChatSend(Chat As String)

Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim RICHCNTL As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Chat$)
Call WaitForTextToLoad(RICHCNTL)
Call SendMessageLong(RICHCNTL, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub ChatSend2(Chat As String)

Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim RICHCNTL As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
Dim TheText As String, TL As Long
TL = SendMessageLong(RICHCNTL, WM_GETTEXTLENGTH, 0&, 0&)
TheText = String(TL + 1, " ")
Call SendMessageByString(RICHCNTL, WM_GETTEXT, TL + 1, TheText)
TheText = Left(TheText, TL)
If TheText = "" Then GoTo justsendchat
Call SendMessageByString(RICHCNTL, WM_CLEAR, 0&, 0&)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Chat$)
Call SendMessageLong(RICHCNTL, WM_CHAR, ENTER_KEY, 0&)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, TheText)
Exit Sub
justsendchat:
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Chat$)
Call SendMessageLong(RICHCNTL, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub SendToChat(Message As String)
Dim RICHCNTL As Long, textlen As Long, RICHCNTLTxt As String
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
textlen& = SendMessage(RICHCNTL&, WM_GETTEXTLENGTH, 0&, 0&)
RICHCNTLTxt$ = String(textlen&, 0&)
Call SendMessageByString(RICHCNTL&, WM_GETTEXT, textlen& + 1&, RICHCNTLTxt$)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, Message$)
Call SendMessageByNum(RICHCNTL&, WM_CHAR, 13&, 0&)
If Len(RICHCNTLTxt$) <> 0& Then Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, RICHCNTLTxt$)
End Sub
Public Sub getchattext(Txt As TextBox)
Dim RICHCNTL As Long, textlen As Long, RICHCNTLTxt As String
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long, txt1 As String
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0&, "richcntlreadonly", vbNullString)
Txt = GetText(RICHCNTL&)
End Sub
Public Function ChatBox() As String
Dim RICHCNTL As Long, textlen As Long, RICHCNTLTxt As String
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long, txt1 As String
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0&, "richcntlreadonly", vbNullString)
If (RICHCNTL&) = 0 Then
ChatBox = ""
Else
ChatBox = (RICHCNTL&)
End If
End Function
Public Sub Ims_On()
Call SendInstantMessage("$IM_OFF", "Bye")
End Sub

Function GetLastChatLine()
Dim Length, length1, chr1, txt1 As String
txt1 = GetChatText1
Do
DoEvents
chr1 = InStr(txt1, Chr(13))
If Not chr1 = 0 Then
txt1 = Mid(txt1, chr1 + 1, Len(txt1))
Else
GoTo done
End If
Loop
done:
If InStr(txt1, "Link -1") = 0 Then GoTo skip:
txt1 = ReplaceString(txt1, Mid(txt1, InStr(txt1, "Link -1"), Len(txt1)), "")
skip:
GetLastChatLine = txt1
End Function
Function GetLastSN()
Dim chatline, Txt
chatline = GetLastChatLine
Txt = Left(chatline, InStr(chatline, ":") - 1)
GetLastSN = Txt
End Function
Function GetLastMSG()
Dim chatline, txt1 As String, space, Txt
chatline = GetLastChatLine
txt1 = Left(chatline, InStr(chatline, ":"))
txt1 = Mid(chatline, Len(txt1) + 2, Len(chatline))
Txt = ReplaceString(txt1, vbTab, "")
GetLastMSG = Txt
End Function


Public Sub WaitForTextToLoad(hwnd As Long)
Dim Count1 As Long, Count2 As Long, Count3 As Long
Do: DoEvents
    Count1& = Len(GetText(hwnd&))
    Call TimeOut(0.5)
    Count2& = Len(GetText(hwnd&))
    Call TimeOut(0.5)
    Count3& = Len(GetText(hwnd&))
Loop Until Count2& = Count1& And Count3& = Count1& And Count3& <> 0&
End Sub
Public Sub WriteToINI(AppName As String, KeyName As String, KeyValue As String, FileName As String)
Call WritePrivateProfileString(AppName$, KeyName$, KeyValue$, FileName$)
End Sub
Public Function GetFromINI(AppName As String, KeyName As String, FileName As String) As String

Dim Buffer As String
Buffer$ = String(255&, Chr(0))
KeyName$ = LCase(KeyName$)
GetFromINI$ = Left(Buffer$, GetPrivateProfileString(AppName$, ByVal KeyName$, "", Buffer$, Len(Buffer$), FileName$))
End Function

Function GetChatText1()

Dim ChatText
Dim AORich As Long
Dim room As Long
room& = FindChat
AORich& = FindChildByClass(room&, "RICHCNTLREADONLY")
GetChatText1 = GetText(AORich&)
End Function
Public Function FindChildByClass(ByVal hParent As Long, ByVal sClassName As String, Optional ByVal nIndex) As Long
   Dim hChild As Long
   Dim i As Integer

   If IsMissing(nIndex) Then
      nIndex = 1
   ElseIf nIndex < 1 Then
      Exit Function
   End If
   hChild = GetWindow(hParent, GW_CHILD)
   While i < nIndex And hChild
      If GetWindowClassName(hChild) = sClassName Then
         i = i + 1
      End If
      
      If i < nIndex Then
         hChild = GetWindow(hChild, GW_HWNDNEXT)
      End If
   Wend
   FindChildByClass = hChild
   Exit Function
End Function
Public Function GetWindowClassName(ByVal hWindow As Long) As String
      Dim sClassName As String * 100
      Dim ret As Long
   ret = GetClassName(hWindow, sClassName, 100)
   GetWindowClassName = Trim$(Left(sClassName, ret))
End Function
Public Sub Save_ListBox(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub Load_ListBox(Directory As String, TheList As ListBox)
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
Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub Keyword7(KW As String)

Dim AOLFrame As Long, AOLToolbar As Long, aolcombobox As Long
Dim editx As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
aolcombobox = FindWindowEx(AOLToolbar, 0&, "_aol_combobox", vbNullString)
editx = FindWindowEx(aolcombobox, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, KW$)
Call SendMessageLong(editx&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(editx&, WM_CHAR, VK_RETURN, 0&)

End Sub
Public Sub EnterPR(room As String)

Call Keyword7("aol://2719:2-2-" & room)
End Sub
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(MyString$, (ToFind))
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
            NewSpot& = InStr(Spot&, (MyString$), (ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
Public Sub WaitForOKOrRoom(room As String)

    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    room$ = (ReplaceString(room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindChat&)
        RoomTitle$ = (ReplaceString(room$, " ", ""))
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or room$ = RoomTitle$
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

Public Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function

Public Sub ChatNow()
Call clickToolbar("3", "N")
End Sub
Public Sub addroom(List As String)

Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim aollistBox As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
aollistBox = FindWindowEx(AOLChild, 0&, "_aol_listbox", vbNullString)
'Call AddAOLListToListbox(aollistbox, list)
End Sub
Function AddString(num As String, str As String) As String
Dim add
For add = 0 To num
AddString = AddString & str
Next
End Function
Public Sub AddAOLListToListbox(ListToGet As Long, ListToPut As ListBox)
  ' Use ADDROOM
    On Error Resume Next
    Dim cprocess As Long, itmhold As Long, ListItem As String
    Dim psnhold As Long, rbytes As Long, i As Integer
    Dim sthread As Long, mthread As Long
    ' Obtain the identifiers of a thread and process that are associated
    ' with the window. A process is a running application and a thread
    ' is a task that the program is doing (like a program could be doing
    ' several things, each of these things would be a thread).
    sthread = GetWindowThreadProcessId(ListToGet, cprocess)
    ' Open the handle to the existing process
    mthread = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess)
    If mthread <> 0 Then
        For i = 0 To SendMessage(ListToGet, LB_GETCOUNT, 0, 0) - 1
            ListItem = String(4, vbNullChar)
            itmhold = SendMessage(ListToGet, LB_GETITEMDATA, ByVal CLng(i), ByVal 0&)
            itmhold = itmhold + 24
            ' Read memory from the address space of the process
            Call ReadProcessMemory(mthread, itmhold, ListItem, 4, rbytes)
            Call CopyMemory(psnhold, ByVal ListItem, 4)
            psnhold = psnhold + 6
            ListItem = String(16, vbNullChar)
            Call ReadProcessMemory(mthread, psnhold, ListItem, Len(ListItem), rbytes)
            ' cut nulls off
            ListItem = Left(ListItem, InStr(ListItem, vbNullChar) - 1)
            ListToPut.AddItem ListItem
        Next i
        Call CloseHandle(mthread)
    End If
End Sub


Public Sub ListRemoveBlanks(TheList As ListBox)
' Self-explanitory
Dim Count&, Count2&
If TheList.ListCount = 0 Then Exit Sub
Do
DoEvents
Count& = 1
Do
DoEvents
If TheList.List(Count&) = "" Then TheList.RemoveItem (Count&)
Count& = Count& + 1
Count2& = TheList.ListCount
Loop Until Count& >= Count2&
Loop Until InStr(TheList.hwnd, "") = 0
End Sub
Public Sub KillDupes(TheList As ListBox)
' Kills duplicates in a listbox.
Dim Count&, Count2&, Count3&
If TheList.ListCount = 0 Then Exit Sub
For Count& = 0 To TheList.ListCount - 1
DoEvents
For Count2& = Count& + 1 To TheList.ListCount - 1
DoEvents
If TheList.List(Count&) = TheList.List(Count2&) Then TheList.RemoveItem (Count2&)
Next Count2&
Next Count&
End Sub
Public Sub TimeOut(Length&)
Dim Time As Long
Time = Timer
Do
DoEvents
Loop Until Timer - Time >= Length
End Sub
Public Function direxists(Search As String) As Boolean

'If Right(Search$, 1) <> "" + "\" Then
'Search$ = Search$ + "\"
'End If
If Dir(Search$) <> "" Then
direxists = True
Else
direxists = False
End If
End Function
Public Sub ReadNew()
' Opens New Mail
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub WriteMail()
' Opens Write Mail
Dim AOLFrame As Long, AOLToolbar As Long, AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
AOLToolbar = FindWindowEx(AOLFrame, 0&, "aol toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar, 0&, "_aol_toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Function FindChat() As Long
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim aollistBox As Long, AOLStatic As Long, AOLIcon As Long
Dim RICHCNTL As Long
AOLFrame& = FindWindow("aol frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild& = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
aollistBox& = FindWindowEx(AOLChild, 0&, "_aol_listbox", vbNullString)
AOLStatic& = FindWindowEx(AOLChild, AOLStatic, "_aol_static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
If RICHCNTL& <> 0& And aollistBox& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
FindChat& = AOLChild&
Exit Function
Else
Do

AOLChild& = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
aollistBox& = FindWindowEx(AOLChild, 0&, "_aol_listbox", vbNullString)
AOLStatic& = FindWindowEx(AOLChild, AOLStatic, "_aol_static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild, AOLIcon, "_aol_icon", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild, 0&, "richcntl", vbNullString)
      If RICHCNTL& <> 0& And aollistBox& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
         FindChat& = AOLChild&
         Exit Function
         End If
         Loop Until AOLChild& = 0&
         End If
         FindChat& = AOLChild&
         If FindChat& <> 0& Then MsgBox "chat not found"
      End Function

Public Sub AddListToListbox(TheList As Long, NewList As ListBox)
    ' This sub will only work with standard listboxes.
    Dim lCount As Long, Item As String, i As Integer, TheNull As Integer
    ' get the item count in the list
    lCount = SendMessageLong(TheList, LB_GETCOUNT, 0&, 0&)
    For i = 0 To lCount - 1
        Item = String(255, Chr(0))
        Call SendMessageByString(TheList, LB_GETTEXT, i, Item)
        TheNull = InStr(Item, Chr(0))
        ' remove any null characters that might be on the end of the string
        If TheNull <> 0 Then
            NewList.AddItem Mid$(Item, 1, TheNull - 1)
        Else
            NewList.AddItem Item
        End If
    Next
End Sub
Public Function GetUser()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim UserString As String
AOLFrame& = FindWindow("aol frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild& = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLChild& = FindWindowEx(MDIClient, AOLChild, "aol child", vbNullString)
UserString$ = GetCaption(AOLChild&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser = UserString
        Exit Function
    Else
        Do
            AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
            UserString$ = GetCaption(AOLChild&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser = UserString
                Exit Function
            End If
        Loop Until AOLChild& = 0&
    End If
    GetUser = ""
End Function
Public Sub ClickIdleOff()

Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon, WM_KEYUP, VK_SPACE, 0&)
End Sub
