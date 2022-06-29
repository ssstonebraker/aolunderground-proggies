Attribute VB_Name = "Kinger2000"
Option Explicit

'KingEr²ººº by søpøn
'you should've received this module with a text document giving you a brief
'description on how to use it.  if you didn't, sucks to be you but it's not
'rocket science so use your scrolling, ieet0 caps, bas file using head to figure
'it out. ;P -søpøn  ("Dim sopon as int@aol.com")

'<start declerations>
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetParent& Lib "user32" (ByVal hwnd As Long)
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetTickCount& Lib "kernel32" ()
Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByVal lpFileSizeHigh As Long) As Long
Declare Function GetWindow& Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long)
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Declare Function ShowWindow& Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long)
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function PostMessage& Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any)
Declare Function PostMessageByString Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SetCursorPos& Lib "user32" (ByVal X As Long, ByVal Y As Long)
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long
'<end declerations>

'<begin constants>

Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158

Public Const EM_SETSEL& = &HB1

Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5

Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT& = &H18B
Public Const LB_SETCURSEL& = &H186

Public Const SW_MINIMIZE& = 6

Public Const WM_CLEAR& = &H303
Public Const WM_CLOSE& = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CHAR = &H102
Public Const WM_ENABLE = &HA
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN& = &H201
Public Const WM_LBUTTONUP& = &H202
Public Const WM_SETTEXT& = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOVE = &HF012
Public Const WM_USER = &H400

Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_UP = &H26
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SPACE = &H20

Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const Flag = SWP_NOMOVE Or SWP_NOSIZE

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

'<end constants>

'<begin types>
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
'<end types>


Function Is40() As Boolean
Dim hIcon As Long
Dim hBar As Long
Dim hFrame As Long
hFrame& = FindWindow("AOL Frame25", vbNullString)
hBar& = FindWindowEx(hFrame, 0&, "AOL Toolbar", vbNullString)
hBar& = FindWindowEx(hBar&, 0&, "_AOL_Toolbar", vbNullString)

If hBar <> 0 Then
    Is40 = True
Else
    Is40 = False
End If
End Function

Function MDI() As Long
Dim hMain&, hMDI&
hMain = FindWindow("AOL Frame25", vbNullString)
hMDI = FindWindowEx(hMain, 0, "MDIClient", vbNullString)
MDI = hMDI
End Function

Function RunMenuByString(SearchString As String) As Long
'taken from dos32.bas (slightly modified to return the id of the menu)
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
                RunMenuByString = sID&
                Exit Function
            End If
        Next LookSub&
    Next LookFor&
End Function

Sub KeepOnTop(frmTop As Form)
SetWindowPos frmTop.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flag
End Sub

Function Menu40(nWhich As Integer) As Long
Dim hAOL&, hTool&, hBar&, hIcon&, hMenu&
Dim nNext%, nVisible%

hAOL = FindWindow("AOL Frame25", vbNullString)
hTool = FindWindowEx(hAOL, 0, "AOL Toolbar", vbNullString)
hBar = FindWindowEx(hTool, 0, "_AOL_Toolbar", vbNullString)
hIcon = FindWindowEx(hBar, 0, "_AOL_Icon", vbNullString)

For nNext = 0 To nWhich - 1
    hIcon = GetWindow(hIcon, 2)
Next nNext

Call SetCursorPos(0, 0)
Call PostMessage(hIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(hIcon, WM_LBUTTONUP, 0&, 0&)

Select Case nWhich
    Case 0:
        Exit Function
    Case 1:
        Exit Function
    Case 3:
        Exit Function
End Select
Do
    DoEvents
    hMenu = FindWindow("#32768", vbNullString)
    nVisible = IsWindowVisible(hMenu)
Loop Until nVisible = 1
Menu40 = hMenu
End Function

Sub ClearText(hwnd As Long)
Call SendMessage(hwnd, EM_SETSEL, -1, 0)
Call SetText(hwnd, "")
End Sub

Sub SetText(hwnd As Long, sText As String)
Call SendMessageByString(hwnd, WM_SETTEXT, 0&, sText)
End Sub

Private Function Room40() As Long
Dim hChild&, hRICH&, hList&, hCombo&, hIcon&
Do
    hChild = FindWindowEx(MDI, hChild, "AOL Child", vbNullString)
    hRICH = FindWindowEx(hChild, 0, "RICHCNTL", vbNullString)
    hList = FindWindowEx(hChild, 0, "_AOL_Listbox", vbNullString)
    hCombo = FindWindowEx(hChild, 0, "_AOL_Combobox", vbNullString)
    hIcon = FindWindowEx(hChild, 0, "_AOL_Combobox", vbNullString)
Loop Until (hIcon <> 0 And hCombo <> 0 And hList <> 0 And hRICH <> 0) Or hChild = 0
Room40 = hChild
End Function

Function GetText(hwnd As Long) As String
Dim sBuffer As String, hLen&
hLen = SendMessageByNum(hwnd, 14, 0&, 0&)
sBuffer$ = Space$(hLen)
Call SendMessageByString(hwnd, 13, hLen + 1, sBuffer$)
GetText = sBuffer

End Function

Private Function UserName40() As String
Dim hChild&, sText$
Do
    hChild = FindWindowEx(MDI, hChild, "AOL Child", vbNullString)
    sText = GetText(hChild)
    If (Left(sText, Len("Welcome,")) = "Welcome,") And (Right(sText, 1) = "!") Then
        sText = Mid(sText, Len("welcome, "), Len(sText) - Len("welcome, "))
        Exit Do
    Else
        GoTo lblNext
    End If
lblNext:
Loop Until hChild = 0
UserName40 = Trim(sText)
End Function

Private Sub OpenFlash40()
Dim hMenu As Long

Dim lpPoint As POINTAPI
Call GetCursorPos(lpPoint)
hMenu = Menu40(2)
Call PostMessage(hMenu, WM_KEYDOWN, &H26, 0&)
Call PostMessage(hMenu, &H101, &H26, 0&)
Call PostMessage(hMenu, WM_KEYDOWN, &H27, 0&)
Call PostMessage(hMenu, &H101, &H27, 0&)
Call PostMessage(hMenu, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(hMenu, &H101, VK_RETURN, 0&)
Call SetCursorPos(lpPoint.X, lpPoint.Y)

Do: DoEvents
Loop Until FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail") <> 0
End Sub

Private Sub OpenNew40()
Dim lpPoint As POINTAPI
Call GetCursorPos(lpPoint)
Call Menu40(0)
Call SetCursorPos(lpPoint.X, lpPoint.Y)
End Sub

Private Sub OpenOld40()
Dim hMenu&, lpPoint As POINTAPI
Call GetCursorPos(lpPoint)
hMenu = Menu40(2)
Call PostMessage(hMenu, WM_CHAR, Asc("o"), 0)
Call SetCursorPos(lpPoint.X, lpPoint.Y)
End Sub

Private Sub OpenSent40()
Dim hMenu&, lpPoint As POINTAPI
Call GetCursorPos(lpPoint)
hMenu = Menu40(2)
Call PostMessage(hMenu, WM_CHAR, Asc("s"), 0)
Call SetCursorPos(lpPoint.X, lpPoint.Y)
End Sub

Private Function CountFlash40() As Integer
Dim hMain&, hTree&
hMain = FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail")
If hMain = 0 Then
    OpenFlash
    Pause 2
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail")
End If
Do
    DoEvents
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0
CountFlash40 = SendMessage(hTree, LB_GETCOUNT, 0, 0)
End Function

Private Function CountNew40() As Integer
Dim hMain&, hTree&, hTab&, hControl&
Dim CountOne%, CountTwo%
If FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox") = 0 Then Call OpenNew40
Do
    DoEvents
    hMain& = FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox")
    hControl = FindWindowEx(hMain, 0, "_AOL_TabControl", vbNullString)
    hTab = FindWindowEx(hControl, 0, "_AOL_TabPage", vbNullString)
    hTree = FindWindowEx(hTab, 0, "_AOL_Tree", vbNullString)
    If hTree <> 0 Then Exit Do
Loop Until hMain = 0
Do
    CountOne = SendMessage(hTree, LB_GETCOUNT, 0, 0)
    Pause 1
    CountTwo = SendMessage(hTree, LB_GETCOUNT, 0, 0)
    Pause 1
Loop Until CountOne = CountTwo
CountNew40 = SendMessage(hTree, LB_GETCOUNT, 0, 0)
End Function

Sub Pause(hInterval As Long)
Dim hCurrent As Long

hInterval = hInterval * 1000
hCurrent = GetTickCount
Do While GetTickCount - hCurrent < Val(hInterval)
DoEvents
Loop
End Sub

Private Function CreateName40(sName As String, sPassword As String) As Boolean
Dim hMain&, hIcon&
Dim hModal&, hText&
'-
Dim hMsgbox&

'-------- main screen -------------
hMain = FindWindowEx(MDI, 0, "AOL Child", "AOL Screen Names")
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
'-----------------------------------

'-------- "try again screen" ----------
hModal& = FindWindow("_AOL_Modal", vbNullString)
hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
hIcon = FindWindowEx(hModal, 0&, "_AOL_Icon", vbNullString)
'--------------------------------------

If hMain <> 0 And hModal <> 0 Then
    'Call PostMessage(hMain, WM_CLOSE, 0, 0)
ElseIf hMain <> 0 And hModal = 0 Then
    Call Pause(1)
    Call Icon(hIcon)
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0, "AOL Child", "Create A Screen Name")
        hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    Loop Until hIcon <> 0
    Call Pause(1)
    Call Icon(hIcon)
    Do
        DoEvents
        hModal& = FindWindow("_AOL_Modal", vbNullString)
        hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
    Loop Until hText <> 0
ElseIf hMain = 0 And hModal = 0 Then
    Call Keyword("names")
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0, "AOL Child", "AOL Screen Names")
        hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
        hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
    Loop Until hIcon <> 0
    Call Pause(1)
    Call Icon(hIcon)
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0, "AOL Child", "Create A Screen Name")
        hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    Loop Until hIcon <> 0
    Call Pause(1)
    Call Icon(hIcon)
    Do
        DoEvents
        hModal& = FindWindow("_AOL_Modal", vbNullString)
        hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
    Loop Until hText <> 0
End If
    
hModal& = FindWindow("_AOL_Modal", vbNullString)
hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
hIcon = FindWindowEx(hModal, 0&, "_AOL_Icon", vbNullString)
Call SetText(hText, sName)
Call Pause(1)
Call Icon(hIcon)

Do
    DoEvents
    hMsgbox = FindWindow("#32770", "America Online")
Loop Until hMsgbox <> 0 Or CreateNamePW40 Or IsWindow(hIcon) = 0

If hMsgbox <> 0 Then 'Or (hIcon <> 0 And hText <> 0) Then
    Do
        DoEvents
        Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
    Loop Until IsWindow(hMsgbox) = 0
    CreateName40 = False
    Exit Function
ElseIf CreateNamePW40 <> 0 Then
    hText = FindWindowEx(CreateNamePW40, 0, "_AOL_Edit", vbNullString)
    hIcon = FindWindowEx(CreateNamePW40, 0, "_AOL_Icon", vbNullString)
    Call SetText(hText, sPassword)
        hText = NextOfClass(hText)
    Call SetText(hText, sPassword)
    Call Icon(hIcon)
    CreateName40 = True
ElseIf hMsgbox = 0 And CreateNamePW40 = 0 Then
    Do
        DoEvents
        hModal& = FindWindow("_AOL_Modal", vbNullString)
        hText& = FindWindowEx(hModal&, 0&, "_AOL_Edit", vbNullString)
        hIcon = FindWindowEx(hModal, 0&, "_AOL_Icon", vbNullString)
    Loop Until hText <> 0 And hIcon <> 0
    CreateName40 = False
    Exit Function
End If
End Function

Private Function CreateName25(sName As String, sPassword As String) As Boolean
Dim hMain&, hBox&
Dim hAlternate&

If FindWindow("_AOL_Modal", "Create a Screen Name") <> 0 Then GoTo lbl_One
If FindWindow("_AOL_Modal", "Create an Alternate Screen Name") <> 0 Then GoTo lbl_One

Call Keyword25("names")
Do
    DoEvents
    hMain& = FindWindowEx(MDI, 0&, "AOL Child", "Create or Delete Screen Names")
    hBox& = FindWindowEx(hMain&, 0&, "_AOL_Listbox", vbNullString)
Loop Until hBox <> 0

Call SendMessage(hBox, (WM_USER + 7), 3, 0)
Call SendMessage(hBox, WM_CHAR, 13, 0)

lbl_One:
Do
    DoEvents
    hMain = FindWindow("_AOL_Modal", "Create a Screen Name")
    hAlternate = FindWindow("_AOL_Modal", "Create an Alternate Screen Name")
    hBox = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    If hBox = 0 Then hBox = FindWindowEx(hAlternate, 0, "_AOL_Edit", vbNullString)
Loop Until hBox <> 0

Call SetText(hBox, sName)
Call SendMessage(hBox, WM_CHAR, 13, 0)
Call Pause(1)

Do
    DoEvents
    hMain = FindWindow("_AOL_Modal", "Set Password")
    hBox = FindWindow("#32770", "America Online")
    hAlternate = FindWindow("_AOL_Modal", "Create an Alternate Screen Name")
    If hMain <> 0 Then
        hBox = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
        Call SetText(hBox, sPassword)
        hBox = NextOfClass(hBox)
        Call SetText(hBox, sPassword)
        Call SendMessage(hBox, WM_CHAR, 13, 0)
        CreateName25 = True
        Exit Function
    End If
    If hBox <> 0 Then
        Call PostMessage(hBox, WM_CLOSE, 0, 0)
        CreateName25 = False
        Exit Function
    End If
    If hAlternate <> 0 Then
        CreateName25 = False
        Exit Function
    End If
Loop Until Len(UserName) = 0

End Function

Private Function CreateNamePW40() As Long
Dim hMain&
Dim hText&, hRICH&, hStatic&

hMain = FindWindow("_AOL_Modal", vbNullString)
hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
hText = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)

If hStatic <> 0 And hRICH <> 0 And hText <> 0 And NextOfClass(hText) <> 0 Then
    CreateNamePW40 = hMain
Else
    CreateNamePW40 = 0
End If

End Function

Function CreateName(sName As String, sPassword As String) As Boolean
If Is25 Then
    CreateName = CreateName25(sName, sPassword)
Else
    CreateName = CreateName40(sName, sPassword)
End If
End Function

Private Function KillModal40() As Boolean
Dim Bval As Boolean
Dim hModal&
Do
    DoEvents
    hModal = FindWindow("_AOL_Modal", vbNullString)
    If hModal <> 0 Then
        Call PostMessage(hModal, WM_CLOSE, 0, 0)
        Bval = True
    End If
    hModal = FindWindow("_AOL_Modal", vbNullString)
Loop Until hModal = 0
KillModal40 = Bval
End Function

Private Sub SendMail40(sWho As String, sSubject As String, sText As String, Optional sBackgroundPic As String, Optional sCarbonCopy As String)
Dim hMain&, hText&, hIcon&
Dim hBox&, hEdit&, hOpen&
Dim nFor%
Call Menu40(1)
LockWindowUpdate (MDI)
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Write Mail")
    hText = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
Loop Until hText <> 0
Call SetText(hText, sWho)
hText = FindWindowEx(hMain, hText, "_AOL_Edit", vbNullString)
Call SetText(hText, sCarbonCopy)
hText = FindWindowEx(hMain, hText, "_AOL_Edit", vbNullString)
Call SetText(hText, sSubject)
hText = FindWindowEx(hMain, hText, "RICHCNTL", vbNullString)
Call SetText(hText, sText)
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
If IsMissing(sBackgroundPic) = True Then
    If Dir(sBackgroundPic) = "" Then Exit Sub
    For nFor = 1 To 8
        hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
    Next nFor
    Call Icon(hIcon)
    Call PostMessage(hIcon, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(hIcon, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(hIcon, WM_CHAR, Asc("b"), 0)
    Do
        DoEvents
        hBox = FindWindow("#32770", "Open")
        hEdit = FindWindowEx(hBox, 0, "Edit", vbNullString)
        hOpen = FindWindowEx(hBox, 0, vbNullString, "&Open")
    Loop Until hEdit <> 0 And hOpen <> 0
    Call SetText(hEdit, sBackgroundPic)
    Call SendMessage(hOpen, WM_KEYDOWN, VK_SPACE, 0)
    Call SendMessage(hOpen, WM_KEYDOWN, VK_SPACE, 0)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
End If
For nFor = 1 To 18
    hIcon = GetWindow(hIcon, 2)
Next nFor
Icon hIcon
Do
    DoEvents
Loop Until KillModal40 = True Or IsWindow(hIcon) = 0
LockWindowUpdate (0)
End Sub

Sub Icon(hIcon As Long)
DoEvents
Call PostMessage(hIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(hIcon, WM_LBUTTONUP, 0&, 0&)
End Sub

Function MailList(lpBox As ListBox, Optional bBCC As Boolean) As String
Dim nNext%, sReturn$
If bBCC = True Then
    sReturn = "("
End If
For nNext = 0 To lpBox.ListCount - 1
    sReturn = sReturn & lpBox.List(nNext) & ","
Next nNext
sReturn = Mid(sReturn, 1, Len(sReturn) - 1)
If bBCC = True Then
    sReturn = sReturn & ")"
End If
MailList = sReturn
End Function

Private Function FwdNew40(lpBox As ListBox, nIndex As Integer, sMessage As String, lpDeadBox As ListBox)
Dim hMain&, hControl&, hTab&, hTree&, hIcon&
Dim hEdit&, hRICH&, hSend&
Dim sString As String
Dim nCount%
Dim bReset As Boolean

bReset = False
If FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox") = 0 Then Call CountNew40
Do
    DoEvents
    Call PostMessage(FindMail40, WM_CLOSE, 0, 0)
    Call PostMessage(FindFwd40, WM_CLOSE, 0, 0)
Loop Until FindMail40 = 0 And FindFwd40 = 0
Do
    DoEvents
    hMain& = FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox")
    hControl = FindWindowEx(hMain, 0, "_AOL_TabControl", vbNullString)
    hTab = FindWindowEx(hControl, 0, "_AOL_TabPage", vbNullString)
    hTree = FindWindowEx(hTab, 0, "_AOL_Tree", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    If hTree <> 0 And hIcon <> 0 Then Exit Do
Loop
Call ShowWindow(hMain, SW_MINIMIZE)
Call PostMessage(hTree, LB_SETCURSEL, CLng(nIndex), 0)
Call LockWindowUpdate(MDI)
Call Icon(hIcon)
Do
    DoEvents
    nCount = nCount + 1
    hMain = FindMail40
    Call PicFix40
    If nCount > 200 Then
        nCount = 0
        Call Icon(hIcon)
    End If
Loop Until hMain <> 0
For nCount = 0 To 7
    hIcon = FindWindowEx(FindMail40, hIcon, "_AOL_Icon", vbNullString)
Next nCount
Call Icon(hIcon)
Reset:
If lpBox.ListCount = 0 Then
    Call PostMessage(FindFwd40, WM_CLOSE, 0, 0)
    Call LockWindowUpdate(0)
    Call KeepAsNew(nIndex)
    Do
        DoEvents
        hMain = FindWindow("#32770", "AOL Mail")
        hMain = FindWindowEx(hMain, 0, "Button", "&No")
    Loop Until hMain <> 0
    Do
        DoEvents
        Call SendMessage(hMain, WM_KEYDOWN, VK_SPACE, 0)
        Call SendMessage(hMain, WM_KEYUP, VK_SPACE, 0)
    Loop Until IsWindow(hMain) = 0
    FwdNew40 = -1
    Exit Function
End If
sString = MailList(lpBox, True)
nCount = 0
Do
    DoEvents
    nCount = nCount + 1
    If nCount = 500 Then
        nCount = 0
        Call Icon(hIcon)
    End If
Loop Until FindFwd40 <> 0
Do
    DoEvents2
    Call PostMessage(FindMail40, WM_CLOSE, 0, 0)
Loop Until FindMail40 = 0
hMain = FindFwd40
hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
Call SetText(hEdit, sString)

If bReset = True Then GoTo lbl_JustClick

hSend = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Call SetText(hRICH, sMessage)
hEdit = FindWindowEx(hMain, hEdit, "_AOL_Edit", vbNullString)
hEdit = FindWindowEx(hMain, hEdit, "_AOL_Edit", vbNullString)
sString = GetText(hEdit)
If LCase(Left(sString, Len("Fwd: "))) = LCase("fwd: ") Then
    sString = Right(sString, Len(sString) - Len("fwd: "))
End If
SetText hEdit, sString
For nCount = 1 To 11
    hSend = FindWindowEx(hMain, hSend, "_AOL_Icon", vbNullString)
Next nCount
Call LockWindowUpdate(0)

lbl_JustClick:
Call Icon(hSend)
nCount = 0
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Error")
    nCount = nCount + 1
    If hMain <> 0 Then
        Call Errored(lpBox, lpDeadBox)
        bReset = True
        GoTo Reset
    End If
    If nCount = 5000 Then
        nCount = 0
        Call Icon(hSend)
    End If
    If KillModal40 = True Then
        Call SendMessage(FindFwd40, WM_CLOSE, 0, 0)
        Call SetPreferences40
    End If
    If IsWindow(hEdit) = 0 Then Exit Do
Loop
Call KeepAsNew40(nIndex)
End Function

Function PicFix40() As Boolean
Dim hMain&, hIcon&, hBox&
PicFix40 = True
hMain = FindWindowEx(MDI, 0, vbNullString, "Picture in E-mail Warning from AOL Neighborhood Watch")
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
hBox = FindWindowEx(hMain, 0, "_AOL_Checkbox", vbNullString)
If hIcon = 0 Then Exit Function
Call SendMessage(hBox, &HF1, 1, 0)
Call Icon(hIcon)
PicFix40 = False
End Function

Private Function FindMail40() As Long
Dim hMain&, hRICH&, hView&, hStatic&, hIcon&
hMain = FindWindowEx(MDI, 0, "AOL Child", vbNullString)
If hMain = 0 Then Exit Function
hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
hView = FindWindowEx(hMain, 0, "_AOL_View", vbNullString)
hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
If GetChildCount(hMain) = 18 Then
    If hRICH <> 0 And hView <> 0 And hStatic <> 0 And hIcon <> 0 Then
        FindMail40 = hMain
        Exit Function
    End If
End If
Do
    DoEvents
    hMain = GetWindow(hMain, 2)
    If hMain = 0 Then Exit Function
    hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
    hView = FindWindowEx(hMain, 0, "_AOL_View", vbNullString)
    hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)

    If GetChildCount(hMain) = 18 Then
        If hRICH <> 0 And hView <> 0 And hStatic <> 0 And hIcon <> 0 Then
            FindMail40 = hMain
            Exit Function
        End If
    End If
Loop
End Function

Private Function FindFwd40() As Long
Dim hMain&, hEdit&, hStatic&, hRICH&, hFont&, hCombo&, hBox&, hIcon&
hMain = FindWindowEx(MDI, 0, "AOL Child", vbNullString)
If hMain = 0 Then Exit Function
hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
hFont = FindWindowEx(hMain, 0, "_AOL_FontCombo", vbNullString)
hCombo = FindWindowEx(hMain, 0, "_AOL_Combobox", vbNullString)
hBox = FindWindowEx(hMain, 0, "_AOL_Checkbox", vbNullString)
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
If hStatic <> 0 And hRICH <> 0 And hFont <> 0 Then
    If hCombo <> 0 And hBox <> 0 And hIcon <> 0 Then
        If GetChildCount(hMain) = 29 Then
            FindFwd40 = hMain
            Exit Function
        End If
    End If
End If
Do
    DoEvents
    hMain = GetWindow(hMain, 2)
    hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
    hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
    hFont = FindWindowEx(hMain, 0, "_AOL_FontCombo", vbNullString)
    hCombo = FindWindowEx(hMain, 0, "_AOL_Combobox", vbNullString)
    hBox = FindWindowEx(hMain, 0, "_AOL_Checkbox", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
'    Call LockWindowUpdate(0)
    If hEdit <> 0 And hStatic <> 0 And hRICH <> 0 And hFont <> 0 Then
        If hCombo <> 0 And hBox <> 0 And hIcon <> 0 Then
            If GetChildCount(hMain) = 29 Then
                FindFwd40 = hMain
                Exit Function
            End If
        End If
    End If
Loop Until hMain = 0
End Function

Function GetChildCount(hParent As Long) As Integer
Dim hMain&, nCount%
If IsWindow(hParent) = False Then Exit Function
hMain = GetWindow(hParent, GW_CHILD)
If hMain <> 0 Then nCount = 1
Do
    DoEvents
    hMain = GetWindow(hMain, GW_HWNDNEXT)
    If hMain <> 0 Then nCount = nCount + 1
Loop Until hMain = 0
GetChildCount = nCount
End Function

Private Sub SetPreferences40()
Dim hMenu&, hMain&, hBox&, hIcon&
    Dim lpPos As POINTAPI
    Call GetCursorPos(lpPos)
    hMenu = Menu40(2)
    Call PostMessage(hMenu, WM_CHAR, 80, 0)
    SetCursorPos lpPos.X, lpPos.Y
    Do
        DoEvents
        hMain = FindWindow("_AOL_Modal", "Mail Preferences")
        hBox = FindWindowEx(hMain, 0, "_AOL_Checkbox", vbNullString)
        hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    Loop Until hBox <> 0 And hIcon <> 0
    Call CheckIt(hBox, False)
    hBox = NextOfClass(hBox)
    Call CheckIt(hBox, True)
    hBox = NextOfClass(hBox)
    Call CheckIt(hBox, False)
    hBox = NextOfClass(hBox)
    Call CheckIt(hBox, False)
    hBox = NextOfClass(hBox)
    Call CheckIt(hBox, True)
    Do
        DoEvents
        Call Icon(hIcon)
    Loop Until IsWindow(hMain) = 0
    Pause 1
End Sub

Function NextOfClass(hwnd As Long) As Long
Dim sClass$, hParent&
sClass = GetClass(hwnd)
hParent = GetParent(hwnd)
NextOfClass = FindWindowEx(hParent, hwnd, sClass, vbNullString)
End Function

Function GetClass$(hwnd&)
Dim sBuffer$, nLen%
sBuffer$ = String$(250, 0)
nLen = GetClassName(hwnd, sBuffer$, 250)
GetClass = Mid(sBuffer, 1, nLen)
End Function

Sub CheckIt(hCheckbox&, bChecked As Boolean)
Call SendMessage(hCheckbox, &HF1, bChecked, 0)
End Sub

Public Function Errored(lpBox As ListBox, lpDeadBox As ListBox) As Integer
Dim hMain&, hView&, hRemoved&, nCount%, sText$, sTest$
Dim nInstr%
hMain = FindWindowEx(MDI, 0, "AOL Child", "Error")
hView = FindWindowEx(hMain, 0, "_AOL_View", vbNullString)
If hView = 0 Then
    Errored = 0
    Exit Function
End If
sText = LCase(GetText(hView))
sText = Mid(sText, 67)
While InStr(1, sText, " - ") <> 0
    nInstr = InStr(1, sText, " - ")
    lpDeadBox.AddItem Mid(sText, 1, nInstr - 1)
    sText = Mid(sText, nInstr + 3)
    nInstr = InStr(nInstr + 1, sText, Chr(13))
    sText = Mid(sText, nInstr + 2)
Wend

For nCount = 0 To lpDeadBox.ListCount - 1
    hRemoved = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, -1, lpDeadBox.List(nCount))
    If hRemoved <> -1 Then lpBox.RemoveItem hRemoved
Next nCount

Do
    DoEvents
    Call PostMessage(hMain, WM_CLOSE, 0, 0)
Loop Until IsWindow(hMain) = 0
End Function

Function RemoveSpaces(sText As String) As String
Dim nPos As Integer
Dim sLft, sRgt As String
While InStr(1, sText, " ") <> 0
    DoEvents
    nPos = InStr(1, sText, " ")
    sLft = Left(sText, nPos - 1)
    sRgt = Right(sText, Len(sText) - nPos)
    sText = sLft & sRgt
Wend
RemoveSpaces = sText
End Function

Sub DoEvents2()
Static nCurrent As Integer
nCurrent = nCurrent + 1
    If nCurrent = 5 Then
        DoEvents
        nCurrent = 0
    End If
End Sub

Private Function FwdFlash40(lpBox As ListBox, nIndex As Integer, sMessage As String, lpDeadBox As ListBox) As Integer
'make sure SetPreferences40 is set before you call this
Dim hMain&, hControl&, hTab&, hTree&, hIcon&
Dim hEdit&, hRICH&, hSend&, hMsgbox&
Dim sString As String
Dim nCount%
nCount = CountFlash40
Do
    DoEvents
    Call PostMessage(FindMail40, WM_CLOSE, 0, 0)
    Call PostMessage(FindFwd40, WM_CLOSE, 0, 0)
Loop Until FindMail40 = 0 And FindFwd40 = 0
hMain = FindWindow("#32770", "AOL Mail")
hMain = FindWindowEx(hMain, 0, "Button", "&No")
If hMain <> 0 Then
    Do
        DoEvents
        Call SendMessage(hMain, WM_KEYDOWN, VK_SPACE, 0)
        Call SendMessage(hMain, WM_KEYUP, VK_SPACE, 0)
    Loop Until IsWindow(hMain) = 0
End If

hMain = FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail")
hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Call ShowWindow(hMain, SW_MINIMIZE)
Call SendMessage(hTree, LB_SETCURSEL, CLng(nIndex), 0&)
Call LockWindowUpdate(MDI)
Call Icon(hIcon)
Do
    DoEvents2
Loop Until FindMail40 <> 0
hMain = FindMail40
For nCount = 0 To 7
    hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
Next nCount
Call Icon(hIcon)
Do
    DoEvents2
    nCount = nCount + 1
    If nCount > 200 Then
        nCount = 0
        Call Icon(hIcon)
    End If
    hMain = FindFwd40
Loop Until hMain <> 0
Do
    Call PostMessage(FindMail40, WM_CLOSE, 0&, 0&)
Loop Until FindMail40 = 0
Reset:
If lpBox.ListCount = 0 Then
    Call PostMessage(FindFwd40, WM_CLOSE, 0, 0)
    Call LockWindowUpdate(0)
    Do
        DoEvents
        hMain = FindWindow("#32770", "AOL Mail")
        hMain = FindWindowEx(hMain, 0, "Button", "&No")
    Loop Until hMain <> 0
    Do
        DoEvents
        Call SendMessage(hMain, WM_KEYDOWN, VK_SPACE, 0)
        Call SendMessage(hMain, WM_KEYUP, VK_SPACE, 0)
    Loop Until IsWindow(hMain) = 0
    FwdFlash40 = -1
    Exit Function
End If
    

sString = MailList(lpBox, True)
nCount = 0
hMain = FindFwd40
hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
hRICH = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
hSend = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)

Call SetText(hEdit, sString)
Call SetText(hRICH, sMessage)

hEdit = NextOfClass(hEdit)
hEdit = NextOfClass(hEdit)
sString = GetText(hEdit)
If LCase(Left(sString, Len("Fwd: "))) = LCase("fwd: ") Then
    sString = Right(sString, Len(sString) - Len("fwd: "))
End If
SetText hEdit, sString
For nCount = 1 To 11
    hSend = FindWindowEx(hMain, hSend, "_AOL_Icon", vbNullString)
Next nCount

Call Icon(hSend)
Call LockWindowUpdate(0)
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Error")
    hMsgbox = FindWindow("#32770", "America Online")
    nCount = nCount + 1
    If hMsgbox <> 0 Then
        Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
        Call PostMessage(FindFwd40, WM_CLOSE, 0, 0)
        Do
            DoEvents
            hMain = FindWindow("#32770", vbNullString)
            hMain = FindWindowEx(hMain, 0, "Button", "&No")
        Loop Until hMain <> 0
        Do
            DoEvents
            Call SendMessage(hMain, WM_KEYDOWN, VK_SPACE, 0)
            Call SendMessage(hMain, WM_KEYUP, VK_SPACE, 0)
        Loop Until IsWindow(hMain) = 0
        FwdFlash40 = -1
        Exit Function
    End If
    If hMain <> 0 Then
        Call Errored(lpBox, lpDeadBox)
        GoTo Reset
    End If
    If nCount = 5000 Then
        nCount = 0
        Call Icon(hSend)
    End If
    If KillModal40 = True Then
        Call SendMessage(FindFwd40, WM_CLOSE, 0, 0)
        Call SetPreferences40
        FwdFlash40 = 1
    End If
    If IsWindow(hEdit) = 0 Then Exit Do
Loop
End Function

Private Sub KeepAsNew40(nIndex As Integer)
Dim hMain&, hControl&, hTab&, hTree&, hIcon&

hMain& = FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox")
hControl = FindWindowEx(hMain, 0, "_AOL_TabControl", vbNullString)
hTab = FindWindowEx(hControl, 0, "_AOL_TabPage", vbNullString)
hTree = FindWindowEx(hTab, 0, "_AOL_Tree", vbNullString)
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
hIcon = NextOfClass(hIcon)
hIcon = NextOfClass(hIcon)

Call SendMessage(hTree, LB_SETCURSEL, CLng(nIndex), 0&)
Call Icon(hIcon)
End Sub

Private Sub ChatSend40(sText As String)
Dim hMain&, hBox&, nCount
hMain = Room
hBox = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
hBox = NextOfClass(hBox)
Call ClearText(hBox)
Call SetText(hBox, sText)
Do
    Call SendMessage(hBox, WM_CHAR, 13, 0&)
    nCount = nCount + 1
Loop Until Len(GetText(hBox)) > 0 Or nCount > 5
End Sub

Private Function IM40(sPerson As String, sText As String) As Boolean
Dim hAOL&, hTool&, hBar&, hIcon&, hMenu&
Dim hMain&, hText&
Dim nNext As Integer

hAOL = FindWindow("AOL Frame25", vbNullString)
hTool = FindWindowEx(hAOL, 0, "AOL Toolbar", vbNullString)
hBar = FindWindowEx(hTool, 0, "_AOL_Toolbar", vbNullString)
hIcon = FindWindowEx(hBar, 0, "_AOL_Icon", vbNullString)
For nNext = 0 To 9
    hIcon = NextOfClass(hIcon)
Next nNext
Call LockWindowUpdate(MDI)
Call PostMessage(hIcon, WM_CHAR, Asc("i"), 0&)
Do
    DoEvents2
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Send Instant Message")
    hText = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Loop Until hText <> 0
Call SetText(hText, sPerson)
hText = FindWindowEx(hMain, 0, "RICHCNTL", vbNullString)
Call SetText(hText, sText)
Call LockWindowUpdate(0)
For nNext = 0 To 7
    hIcon = NextOfClass(hIcon)
Next nNext
Call Icon(hIcon)
Do
    DoEvents2
    hMain = FindWindow("#32770", "America Online")
    nNext = nNext + 1
    If nNext > 1000 Then
        Call Icon(hIcon)
        nNext = 0
    End If
    If hMain <> 0 Then
        Call SendMessage(hMain, WM_CLOSE, 0, 0)
        Call SendMessage(GetParent(hIcon), WM_CLOSE, 0, 0)
        IM40 = False
        Exit Function
    End If
Loop Until IsWindow(hText) = 0
IM40 = True
End Function

Sub SetCD(bOpen As Boolean) 'opens or closes the cd
If bOpen = False Then
    DoEvents
    Call mciSendString("set CDAudio door closed", vbNullString, 0, 0)
Else
    DoEvents
    Call mciSendString("set CDAudio door open", vbNullString, 0, 0)
End If
End Sub

Sub Playwav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub

Sub FormDrag(frmMain As Form)
    Call ReleaseCapture
    Call SendMessage(frmMain.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Function GetFromINI(sPath As String, sSection As String, sKey As String) As String
If Dir(sPath) = "" Then 'file is not there so we'll create it for the user
    Open sPath For Append As #1
    Close #1
    GetFromINI = ""
    Exit Function
End If

Dim hLen As Long
Dim sBuffer As String
sBuffer = Space(255)
hLen = GetPrivateProfileString(sSection, sKey, "", sBuffer, 255, sPath)
GetFromINI = Mid(sBuffer, 1, hLen)
End Function

Private Sub Keyword40(sWord As String)
    Dim hMain&, hTool&, hBar&
    Dim hCombo&, hEdit&
    hMain = FindWindow("AOL Frame25", vbNullString)
    hTool = FindWindowEx(hMain, 0&, "AOL Toolbar", vbNullString)
    hBar = FindWindowEx(hTool, 0&, "_AOL_Toolbar", vbNullString)
    hCombo = FindWindowEx(hBar, 0&, "_AOL_Combobox", vbNullString)
    hEdit = FindWindowEx(hCombo, 0&, "Edit", vbNullString)
    Call SendMessageByString(hEdit, WM_SETTEXT, 0&, sWord)
    Call SendMessageLong(hEdit, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(hEdit, WM_CHAR, VK_RETURN, 0&)
End Sub

Sub WriteToINI(sPath As String, sSection As String, sKey As String, sValue As String)
Call WritePrivateProfileString(sSection, sKey, sValue, sPath)
End Sub

Public Function LBText(hwnd As Long, nIndex As Integer) As String
Dim sBuffer As String
Dim nLen As Integer

nLen = SendMessage(hwnd, LB_GETTEXTLEN, nIndex, 0)
If nLen = -1 Then Exit Function
sBuffer = String(nLen, 0)
nLen = SendMessageByString(hwnd, LB_GETTEXT, nIndex, sBuffer)
LBText = sBuffer
End Function

Function LBText16(hwnd As Long, nIndex As Integer) As String
'this function still doesn't work ;(
Dim sBuffer As String
Dim nLen As Integer

nLen = SendMessage(hwnd, WM_USER + 11, nIndex, 0)
If nLen = -1 Then Exit Function
sBuffer = Space(100)
nLen = SendMessageByString(hwnd, WM_USER + 10, nIndex, sBuffer)
LBText16 = Trim(sBuffer)
End Function

Private Sub Incoming²Box40(lpBox As ListBox)
Dim hMain&, hTree&
Dim nCount%, nFor%, nInstr%
Dim sParse

hMain = FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail")
If hMain = 0 Then
    Call OpenFlash40
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail")
    Loop Until FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString) <> 0
End If

hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
nCount = CountFlash40
Call ShowWindow(hMain, SW_MINIMIZE)

For nFor = 0 To nCount - 1
    sParse = LBText(hTree, nFor)
    nInstr = InStr(sParse, Chr(9))
    sParse = Mid(sParse, nInstr + 1)
    nInstr = InStr(sParse, Chr(9))
    sParse = Mid(sParse, nInstr + 1)
    sParse = Trim(sParse)
    lpBox.AddItem sParse
Next nFor
Call SendMessage(hMain, WM_CLOSE, 0, 0)
End Sub

Private Sub WriteMail40(sWho As String, sSubject As String, sMessage As String, Optional lpBox As ListBox, Optional lpDeadBox As ListBox)
Dim hMain&, hText&, nCount%

If FindWindowEx(MDI, 0, "AOL Child", "Write Mail") = 0 Then Call Menu40(1)

Call LockWindowUpdate(MDI)
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Write Mail")
    hText = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
Loop Until hText <> 0

Reset:
If Len(sWho) = 0 Then
    Call SetText(hText, MailList(lpBox, True))
Else
    Call SetText(hText, sWho)
End If
hText = NextOfClass(hText)
hText = NextOfClass(hText)
Call SetText(hText, sSubject)
hText = FindWindowEx(hMain, hText, "RICHCNTL", vbNullString)
Call SetText(hText, sMessage)
hText = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)

For nCount = 1 To 13
    hText = NextOfClass(hText)
Next nCount

Call Icon(hText)
Call LockWindowUpdate(0)

Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, vbNullString, "Error")
    nCount = nCount + 1
    If hMain <> 0 And IsMissing(lpBox) = True Then
        Call Errored(lpBox, lpDeadBox)
        GoTo Reset
    End If
    If nCount = 500 Then
        nCount = 0
        Call Icon(hText)
    End If
    If KillModal40 = True Then
        Call SendMessage(GetParent(hText), WM_CLOSE, 0, 0)
        Call SetPreferences40
    End If
    If IsWindow(hText) = 0 Then Exit Do
Loop
End Sub

Function SearchLB(lpBox As ListBox, sSearchString As String, ByRef sBuffer As String, Optional nMaxFinds As Integer) As Integer
Dim nCount As Integer, nLoop As Integer

For nLoop = 0 To lpBox.ListCount - 1
    If InStr(1, LCase(lpBox.List(nLoop)), LCase(sSearchString)) <> 0 Then
        sBuffer = sBuffer & nLoop + 1 & ")" & Space(5) & lpBox.List(nLoop) & Chr(13)
        nCount = nCount + 1
    End If
    If nCount = nMaxFinds Then
        SearchLB = nCount
        Exit Function
    End If
Next nLoop
SearchLB = nCount
End Function

Function LBDupe(lpBox As ListBox) As Integer
Dim nCount As Integer, nPos1 As Integer, nPos2 As Integer, nDelete As Integer
Dim sText As String
If lpBox.ListCount < 3 Then
    LBDupe = 0
    Exit Function
End If
For nCount = 0 To lpBox.ListCount - 1
    Do
        DoEvents2
        sText = LBText(lpBox.hwnd, nCount)
        nPos1 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nCount, sText)
        nPos2 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nPos1 + 1, sText)
            If nPos2 = -1 Or nPos2 = nPos1 Then Exit Do
        lpBox.RemoveItem nPos2
        nDelete = nDelete + 1
    Loop
Next nCount
LBDupe = nDelete
End Function

Function RandomInt(nMax As Integer) As Integer
Randomize
RandomInt = Int((Rnd * nMax) + 1)
End Function

Private Function CountOld40() As Integer
Dim hMain&, hTab&, hPage&
Dim hTree&, nCount%, nCount2%
Call SendMessage(Menu40(2), WM_CHAR, Asc("o"), 0)

Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox")
    hTab = FindWindowEx(hMain, 0, "_AOL_TabControl", vbNullString)
    hPage = FindWindowEx(hTab, 0, "_AOL_TabPage", vbNullString)
    hPage = NextOfClass(hPage)
    hTree = FindWindowEx(hPage, 0, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0
Do
    nCount = SendMessage(hTree, LB_GETCOUNT, 0, 0)
    Pause 1
    nCount2 = SendMessage(hTree, LB_GETCOUNT, 0, 0)
    Pause 1
Loop Until nCount = nCount2
CountOld40 = nCount2
End Function

Private Function CountSent40() As Integer
    Dim hMain&, hTab&, TabPage&, hPage&, hTree&
    Dim mTree&, nCount%, nCount2%
    Call SendMessage(Menu40(2), WM_CHAR, Asc("s"), 0)
    
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox")
    hTab = FindWindowEx(hMain&, 0&, "_AOL_TabControl", vbNullString)
    hPage = FindWindowEx(hTab, 0&, "_AOL_TabPage", vbNullString)
    hPage = FindWindowEx(hTab, hPage, "_AOL_TabPage", vbNullString)
    hPage = FindWindowEx(hTab, hPage, "_AOL_TabPage", vbNullString)
    hTree = FindWindowEx(hPage, 0&, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0

Do
    DoEvents
    nCount = SendMessage(hTree, LB_GETCOUNT, 0, 0)
    Pause 2
    nCount2 = SendMessage(hTree, LB_GETCOUNT, 0, 0)
    Pause 2
Loop Until nCount = nCount2
CountSent40 = nCount

End Function

Sub AIM_IM(sWho As String, sMessage As String)
'for whatever reason that i really didn't feel like figuring out, a
'box comes up askign to warn [sWho]. i don't know why, but it won't
'happen every single time. weird...

Dim hMain&, hGroup&, hIcon&, hSend&
Dim hClass&, hMessage&, hWaring&

hMain = FindWindow("_Oscar_BuddyListWin", vbNullString)
hGroup = FindWindowEx(hMain, 0, "_Oscar_TabGroup", vbNullString)
hIcon = FindWindowEx(hGroup, 0, "_Oscar_IconBtn", vbNullString)
Call Icon(hIcon)

Do
    DoEvents
    hMain = FindWindow("AIM_IMessage", "Instant Message")
    hGroup = FindWindowEx(hMain, 0, "_Oscar_PersistantCombo", vbNullString)
    hIcon = FindWindowEx(hGroup, 0, "Edit", vbNullString)
    hSend = FindWindowEx(hMain, 0, "_Oscar_IconBtn", vbNullString)
    hClass = FindWindowEx(hMain, 0, "WndAte32Class", vbNullString)
    hClass = NextOfClass(hClass)
    hMessage = FindWindowEx(hClass, 0, "Ate32Class", vbNullString)
Loop Until hIcon <> 0 And hSend <> 0

Call SetText(hIcon, sWho)
Call SetText(hMessage, sMessage)
Call Icon(hSend)
End Sub

Function CBDupe(lpBox As ComboBox) As Integer
Dim nCount As Integer, nPos1 As Integer, nPos2 As Integer, nDelete As Integer
Dim sText As String
If lpBox.ListCount < 3 Then
    CBDupe = 0
    Exit Function
End If
For nCount = 0 To lpBox.ListCount - 1
    Do
        DoEvents
        sText = lpBox.List(nCount)
        Debug.Print nCount
        nPos1 = SendMessageByString(lpBox.hwnd, CB_FINDSTRINGEXACT, nCount, sText)
        nPos2 = SendMessageByString(lpBox.hwnd, CB_FINDSTRINGEXACT, nPos1 + 1, sText)
            If nPos2 = -1 Or nPos2 = nPos1 Then Exit Do
        lpBox.RemoveItem nPos2
        nDelete = nDelete + 1
    Loop
Next nCount
CBDupe = nDelete
End Function

Private Sub Ghost40(bGhosting As Boolean)
Dim hMain&, hIcon&
Dim hSetup&, hPrivacy&, hMsgbox&
Dim bOpen As Boolean
Dim nNext As Integer

If Len(UserName40) = 0 Then Exit Sub
hMain = FindWindowEx(MDI, 0&, "AOL Child", "Buddy List Window")
bOpen = True
If hMain = 0 Then
    bOpen = False
    Call Keyword40("buddy view")
    Do
        DoEvents
        hMain = FindWindowEx(MDI, 0&, "AOL Child", "Buddy List Window")
        hIcon = FindWindowEx(hMain&, 0&, "_AOL_Icon", vbNullString)
        hIcon = FindWindowEx(hMain&, hIcon, "_AOL_Icon", vbNullString)
        hIcon = FindWindowEx(hMain&, hIcon, "_AOL_Icon", vbNullString)
        If GetText(hIcon) <> "Setup" Then hIcon = FindWindowEx(hMain&, hIcon, "_AOL_Icon", vbNullString)
    Loop Until hIcon <> 0
End If
hIcon = FindWindowEx(hMain&, 0&, "_AOL_Icon", vbNullString)
hIcon = FindWindowEx(hMain&, hIcon, "_AOL_Icon", vbNullString)
hIcon = FindWindowEx(hMain&, hIcon, "_AOL_Icon", vbNullString)
If GetText(hIcon) <> "Setup" Then hIcon = FindWindowEx(hMain&, hIcon, "_AOL_Icon", vbNullString)
Call Icon(hIcon)
Do
    DoEvents
    hSetup& = FindWindowEx(MDI, 0&, "AOL Child", "Soponizer's Buddy Lists")
    hIcon& = FindWindowEx(hSetup&, 0&, "_AOL_Icon", vbNullString)
    For nNext = 1 To 4
        hIcon& = FindWindowEx(hSetup, hIcon&, "_AOL_Icon", vbNullString)
    Next nNext
Loop Until hIcon <> 0

If bOpen = False Then Call PostMessage(hMain, WM_CLOSE, 0, 0)
Call Icon(hIcon)
Do
    DoEvents
    hPrivacy& = FindWindowEx(MDI, 0&, "AOL Child", "Privacy Preferences")
    hIcon& = FindWindowEx(hPrivacy&, 0&, "_AOL_Checkbox", vbNullString)
Loop Until hIcon <> 0
Call PostMessage(hSetup, WM_CLOSE, 0, 0)
If bGhosting = True Then
    For nNext = 1 To 4
        hIcon = FindWindowEx(hPrivacy&, hIcon, "_AOL_Checkbox", vbNullString)
    Next nNext
End If

Call Icon(hIcon)

Do
    DoEvents
    hIcon = FindWindowEx(hPrivacy&, hIcon, "_AOL_Checkbox", vbNullString)
Loop Until GetText(hIcon) = "Buddy List and Instant Message"

Call Icon(hIcon)

Do
    DoEvents
    hIcon = FindWindowEx(hPrivacy&, hIcon, "_AOL_Icon", vbNullString)
Loop Until GetText(hIcon) = "Save"

Do
    DoEvents
    Call Icon(hIcon)
    hMsgbox = FindWindow("#32770", "America Online")
Loop Until hMsgbox <> 0
Call SendMessage(hMsgbox, WM_CLOSE, 0, 0)
End Sub

Private Sub New²Box40(lpBox As ListBox)
Dim hMain&, hTree&, nCount%
Dim nFor%, nInstr%, sParse$
nCount = CountNew40
hMain = FindWindowEx(MDI, 0, "AOL Child", UserName40 & "'s Online Mailbox")
hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)

Call ShowWindow(hMain, SW_MINIMIZE)

For nFor = 0 To nCount - 1
    sParse = LBText(hTree, nFor)
    nInstr = InStr(sParse, Chr(9))
    sParse = Mid(sParse, nInstr + 1)
    nInstr = InStr(sParse, Chr(9))
    sParse = Mid(sParse, nInstr + 1)
    sParse = Trim(sParse)
    lpBox.AddItem sParse
Next nFor
Call SendMessage(hMain, WM_CLOSE, 0, 0)

End Sub

Private Function FwdFlash25(lpBox As ListBox, nIndex As Integer, sMessage As String, lpDeadBox As ListBox) As Integer
'before calling this function you had better make sure that
'you called the setpreferences first or else your program might
'get hung.

Dim hMain&, hTree&, hIcon&
Dim hFwd&, hEdit&, sSubject As String
Dim hMsgbox&
Dim nCount As Integer

If lpBox.ListCount = 0 Then Exit Function

hMain = FindWindowEx(MDI, 0, "A0L Child", "Incoming FlashMail")
nCount = CountFlash25

If nIndex > nCount Then
    FwdFlash25 = -1
    Exit Function
End If


Do
    DoEvents
    hMain& = FindWindowEx(MDI&, hMain&, "AOL Child", "Incoming FlashMail")
    hTree& = FindWindowEx(hMain&, 0&, "_AOL_Tree", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Button", "Open")
Loop Until hTree <> 0

Call SendMessage(hTree, (WM_USER + 7), nIndex, 0)
Call Button(hIcon)
Call LockWindowUpdate(MDI)

Do
    DoEvents
    hMain = FindMail25
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Loop Until hIcon <> 0

hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
Call Icon(hIcon)

Reset:
If lpBox.ListCount = 0 Then
    Call PostMessage(FindFwd25, WM_CLOSE, 0, 0)
'    Do
'        DoEvents
'        hMain = FindWindow("#32770", vbNullString)
'        hMain = FindWindowEx(hMain, 0, "Button", "&No")
'    Loop Until hMain <> 0
'    Do
'        DoEvents
'        Call Button(hMain)
'    Loop Until IsWindow(hMain) = 0
    FwdFlash25 = -1
    Exit Function
End If
Do
    DoEvents
    hFwd = FindFwd25
    hEdit = FindWindowEx(hFwd, 0, "_AOL_Edit", vbNullString)
    hIcon = FindWindowEx(hFwd, 0, "_AOL_Icon", vbNullString)
Loop Until hIcon <> 0 And hEdit <> 0

Call PostMessage(hMain, WM_CLOSE, 0, 0)

Call SetText(hEdit, MailList(lpBox, True))
    hEdit = FindWindowEx(hFwd, hEdit, "_AOL_Edit", vbNullString)
    hEdit = FindWindowEx(hFwd, hEdit, "_AOL_Edit", vbNullString)
    
sSubject = GetText(hEdit)
If LCase(Left(sSubject, Len("Fwd: "))) = LCase("fwd: ") Then
    sSubject = Right(sSubject, Len(sSubject) - Len("fwd: "))
End If

Call SetText(hEdit, sSubject)
    hEdit = FindWindowEx(hFwd, hEdit, "_AOL_Edit", vbNullString)

Call Icon(hIcon)
Call LockWindowUpdate(0)

Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Error")
    hMsgbox = FindWindow("#32770", "America Online")
    nCount = nCount + 1
    If hMsgbox <> 0 Then
        Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
        Call PostMessage(FindFwd25, WM_CLOSE, 0, 0)
        FwdFlash25 = -1
        Exit Function
    End If
    If hMain <> 0 Then
        Call Errored(lpBox, lpDeadBox)
        GoTo Reset
    End If
    If nCount = 5000 Then
        nCount = 0
        Call Icon(hIcon)
    End If
    If KillModal25 = True Then
        Call SendMessage(FindFwd25, WM_CLOSE, 0, 0)
        Call SetPreferences25
        FwdFlash25 = 1
        Exit Function
    End If
    If IsWindow(hEdit) = 0 Then Exit Do
Loop
End Function

Private Function CountFlash25() As Integer
Dim hMain&, hTree&
Dim CountTwo%, CountOne%
Dim bOpen As Boolean

bOpen = True
hMain& = FindWindowEx(MDI&, hMain&, "AOL Child", "Incoming FlashMail")
If hMain = 0 Then
    Call RunMenuByString("Read &Incoming Mail")
    bOpen = False
End If
Do
    DoEvents
    hMain& = FindWindowEx(MDI&, hMain&, "AOL Child", "Incoming FlashMail")
    hTree& = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0

If bOpen = True Then GoTo NoWait
Do
    Pause 1
    CountTwo = SendMessage(hTree, WM_USER + 12, 0, 0)
    Pause 1
    CountOne = SendMessage(hTree, WM_USER + 12, 0, 0)
Loop Until CountOne = CountTwo
NoWait:
CountFlash25 = SendMessage(hTree, WM_USER + 12, 0, 0)
End Function

Private Function CountNew25() As Integer
Dim hMain&, hTree&, hTab&, hControl&
Dim CountOne%, CountTwo%
If FindWindowEx(MDI, 0, "AOL Child", "New Mail") = 0 Then
    Dim bOpen As Boolean
    bOpen = OpenNew25
    If bOpen = False Then
        CountNew25 = 0
        Exit Function
    End If
End If

Do
    DoEvents
    hMain& = FindWindowEx(MDI, 0, "AOL Child", "New Mail")
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
    If hTree <> 0 Then Exit Do
Loop Until hMain = 0
Do

    Pause 1
    CountTwo = SendMessage(hTree, WM_USER + 12, 0, 0)
    Pause 1
    CountOne = SendMessage(hTree, WM_USER + 12, 0, 0)
Loop Until CountOne = CountTwo
CountNew25 = SendMessage(hTree, WM_USER + 12, 0, 0)

End Function

Private Function FindFwd25() As Long
Dim hMain&, hStatic&, hEdit&

Do
    DoEvents
    hMain = FindWindowEx(MDI, hMain, "AOL Child", vbNullString)
    hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
    hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
Loop Until (GetText(hStatic) = "Send Now" And hEdit <> 0) Or hMain = 0
FindFwd25 = hMain
End Function

Private Function FwdNew25(lpBox As ListBox, nIndex As Integer, sMessage As String, lpDeadBox As ListBox) As Integer
Dim hMain&, hTree&, hIcon&, hMsgbox&
Dim hTextbox&, bReseting As Boolean
Dim sText As String
Dim nCount As Integer

hMain = FindWindowEx(MDI, 0, "AOL Child", "New Mail")
hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)

If hTree = 0 Then If OpenNew25 = False Then Exit Function Else Pause 1

Do
    hMain = FindWindowEx(MDI, 0, "AOL Child", "New Mail")
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
    hIcon& = FindWindowEx(hMain, 0, "_AOL_Button", "Read")
Loop Until hTree <> 0 And hIcon <> 0

Call PostMessage(hTree, (WM_USER + 7), nIndex, 0)
Call Button(hIcon)
Call Icon(hIcon)
Call LockWindowUpdate(MDI)
Do
    DoEvents
    hMain = FindMail25
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    hIcon = NextOfClass(hIcon)
Loop Until hIcon <> 0

Call Icon(hIcon): Call Icon(hIcon)

lbl_Reset:
Do
    DoEvents
    nCount = nCount + 1
    hMain = FindFwd25
    hTextbox = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    If nCount > 100 Then
        nCount = 0
        Call Icon(hIcon): Call Icon(hIcon)
    End If
Loop Until hTextbox <> 0
If FindMail25 <> 0 Then
    Do
        DoEvents
        Call PostMessage(FindMail25, WM_CLOSE, 0, 0)
    Loop Until FindMail25 = 0
End If

sText = MailList(lpBox, True)
Call SetText(hTextbox, sText)
If bReseting = True Then GoTo lbl_Click
hTextbox = NextOfClass(hTextbox)
hTextbox = NextOfClass(hTextbox)

sText = GetText(hTextbox)
If LCase(Left(sText, Len("fwd: "))) = LCase("fwd: ") Then sText = Right(sText, Len(sText) - Len("fwd: "))
Call SetText(hTextbox, sText)

hTextbox = NextOfClass(hTextbox)
Call SetText(hTextbox, sMessage)

lbl_Click:
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)

Call LockWindowUpdate(0)
Call Icon(hIcon)
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Error")
    hMsgbox = FindWindow("#32770", "America Online")
    nCount = nCount + 1
    
    If nCount > 500 Then
        nCount = 0
        Call KillWait
        Call Icon(hIcon)
    End If
    If hMain <> 0 Then
        Call Errored(lpBox, lpDeadBox)
        bReseting = True
        GoTo lbl_Reset
    End If
    If hMsgbox <> 0 Then
        Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
        Call PostMessage(FindFwd25, WM_CLOSE, 0, 0)
        FwdNew25 = -1
        Exit Do
    End If
    If KillModal25 = True Then
        Call PostMessage(FindFwd25, WM_CLOSE, 0, 0)
        Call SetPreferences25
        FwdNew25 = 1
        Exit Do
    End If
    If IsWindow(hIcon) = 0 Then Exit Do
Loop
Call KeepAsNew25(nIndex)
End Function

Private Function CountOld25() As Integer
Dim hMain&, hTree&
Call RunMenuByString("Check Mail You've &Read")

Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Old Mail")
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0
Pause 0.5
CountOld25 = SendMessage(hTree, WM_USER + 12, 0, 0)

End Function

Private Sub ChatSend25(sText As String)
Dim hMain&, hBox&, nCount
hMain = Room25
hBox = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
Call ClearText(hBox)
Call SetText(hBox, sText)
Do
    Call SendMessage(hBox, WM_CHAR, 13, 0&)
    nCount = nCount + 1
Loop Until Len(GetText(hBox)) > 0 Or nCount > 5
End Sub

Private Function CountSent25() As Integer
Dim hMain&, hTree&
Call RunMenuByString("Check Mail You've &Sent")
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Outgoing Mail")
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0
Pause 0.5
CountSent25 = SendMessage(hTree, WM_USER + 12, 0, 0)
End Function

Private Function FindMail25() As Long
Dim hMain&, hStatic&, hIcon&, hButton&
Do
    DoEvents
    hMain = FindWindowEx(MDI, hMain, "AOL Child", vbNullString)
    hStatic = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
    hButton = FindWindowEx(hMain, 0, "_AOL_Button", vbNullString)
Loop Until (GetText(hIcon) = "Reply" And hIcon <> 0 And hButton <> 0) Or hMain = 0

FindMail25 = hMain
End Function

Private Function IM25(sPerson As String, sText As String) As Boolean
Dim hMain&, hEdit&
Dim hButton&, hMsgbox&
Dim hIcon&

Call RunMenuByString("Send an Instant Message")
Call KillModal25
Do
    DoEvents
    hMain& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    hEdit& = FindWindowEx(hMain&, 0&, "_AOL_Edit", vbNullString)
Loop Until hEdit <> 0
Call SetText(hEdit, sPerson)
    hEdit& = FindWindowEx(hMain&, hEdit&, "_AOL_Edit", vbNullString)
Call SetText(hEdit, sText)
    hIcon& = FindWindowEx(hMain&, hEdit&, "_AOL_Button", vbNullString)

Do
    Call SendMessage(hIcon, WM_KEYDOWN, VK_SPACE, 0)
    Call SendMessage(hIcon, WM_KEYUP, VK_SPACE, 0)
    hMsgbox = FindWindow("#32770", vbNullString)
Loop Until IsWindow(hIcon) = 0 Or hMsgbox <> 0

If hMsgbox <> 0 Then
    Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
    Call PostMessage(hMain, WM_CLOSE, 0, 0)
    IM25 = False
Else
    IM25 = True
End If

End Function

Private Sub Incoming²Box25(lpBox As ListBox)
Dim hMain&, hTree&
Dim nCount%, nLen%
Dim sBuffer As String

nCount = CountFlash25

Do
    DoEvents
    hMain& = FindWindowEx(MDI&, hMain&, "AOL Child", "Incoming FlashMail")
    hTree& = FindWindowEx(hMain&, 0&, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0

For nCount = 0 To nCount - 1
    nLen = SendMessage(hTree, (WM_USER + 11), nCount, 0)
    sBuffer = String(nLen + 1, 0)
    nLen = SendMessage(hTree, (WM_USER + 10), nCount, sBuffer)
    'MsgBox sBuffer
Next nCount

End Sub

Function Is25() As Boolean
Dim hIcon As Long, hTool As Long, hMain As Long

hMain& = FindWindow("AOL Frame25", vbNullString)
hTool& = FindWindowEx(hMain&, 0&, "AOL Toolbar", vbNullString)
hIcon& = FindWindowEx(hTool&, 0&, "_AOL_Icon", vbNullString)
hIcon& = FindWindowEx(hTool&, hIcon&, "_AOL_Icon", vbNullString)

If Len(GetText(hIcon)) = 0 And hIcon <> 0 Then
    Is25 = True
Else
    Is25 = False
End If

End Function

Private Sub KeepAsNew25(nIndex As Integer)
Dim hMain&, hTree&, hIcon&

hMain = FindWindowEx(MDI, 0, "AOL Child", "New Mail")
hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
hIcon = FindWindowEx(hMain, 0, "_AOL_Button", "Keep As New")

If hTree = 0 Then Exit Sub

Call PostMessage(hTree, (WM_USER + 7), nIndex, 0)
Call Button(hIcon): Call Icon(hIcon)

End Sub

Private Sub Ghost25(bGhosting As Boolean)
Dim hMain&, hIcon&
Call Keyword("buddy")
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", UserName & "'s Buddy Lists")
Loop Until hMain <> 0
Do
    hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
Loop Until GetText(hIcon) = "  Privacy Preferences  "
Call Icon(hIcon)
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Privacy Preferences")
    hIcon = FindWindowEx(hMain, 0, "_AOL_Static", vbNullString)
Loop Until hIcon <> 0
If bGhosting = True Then
        hIcon = FindWindowEx(hMain, hIcon, "_AOL_Button", "Block all AOL members and AOL Instant Messenger users")
Else
       hIcon = FindWindowEx(hMain, hIcon, "_AOL_Button", "Allow all AOL members and AOL Instant Messenger users")
End If
Call Button(hIcon)
Do
    DoEvents
    hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
Loop Until GetText(hIcon) = "Save"

Call Icon(hIcon)

Do
    DoEvents
    hMain = FindWindow("#32770", "America Online")
Loop Until hMain <> 0

Call PostMessage(hMain, WM_CLOSE, 0, 0)
Call PostMessage(FindWindowEx(MDI, 0, "AOL Child", UserName & "'s Buddy Lists"), WM_CLOSE, 0, 0)
End Sub

Private Sub Keyword25(sWord As String)
Dim hMain&, hEdit&, hIcon&

Call RunMenuByString("Keyword...")
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Keyword")
    hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Loop Until hEdit <> 0 And hIcon <> 0

Call SetText(hEdit, sWord)
Call Icon(hIcon)

End Sub

Private Function KillModal25() As Boolean
Dim hMain&
Dim nCount As Integer

Do
    DoEvents
    hMain = FindWindow("_AOL_Modal", vbNullString)
    If hMain <> 0 Then
        nCount = nCount + 1
        Call PostMessage(hMain, WM_CLOSE, 0, 0)
    End If
Loop Until hMain = 0

If nCount = 0 Then KillModal25 = False Else KillModal25 = True
End Function

Private Sub New²Box25(lpBox As ListBox)
Dim hMain&, hTree&
Dim sParse As String, nInstr%, nFor%, nCount%

Dim bOpen As Boolean
bOpen = OpenNew25
If bOpen = False Then Exit Sub
hMain = FindWindowEx(MDI, 0, "AOL Child", "New Mail")
hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
nCount = CountNew25

Call ShowWindow(hMain, SW_MINIMIZE)

For nFor = 0 To nCount - 1
    sParse = LBText16(hTree, nFor)
    nInstr = InStr(sParse, Chr(9))
    sParse = Mid(sParse, nInstr + 1)
    nInstr = InStr(sParse, Chr(9))
    sParse = Mid(sParse, nInstr + 1)
    sParse = Trim(sParse)
    lpBox.AddItem sParse
Next nFor
Call SendMessage(hMain, WM_CLOSE, 0, 0)
End Sub

Private Sub OpenFlash25()
Dim hMain&, hTree&

Call RunMenuByString("Read &Incoming Mail")
Do
    DoEvents
    hMain& = FindWindowEx(MDI&, hMain&, "AOL Child", "Incoming FlashMail")
    hTree& = FindWindowEx(hMain&, 0&, "_AOL_Tree", vbNullString)
Loop Until hTree <> 0
End Sub

Private Function OpenNew25() As Boolean
Dim hMain&, hTree&, hMsgbox&

Call RunMenuByString("Read &New Mail")
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "New Mail")
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
    hMsgbox = FindWindow("#32770", "America Online")
Loop Until hTree <> 0 Or hMsgbox <> 0

If hMsgbox <> 0 Then
    Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
    OpenNew25 = False
Else
    OpenNew25 = True
End If

End Function

Private Function OpenOld25() As Boolean
Dim hMain&, hTree&, hMsgbox&

Call RunMenuByString("Check Mail You've &Read")
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Old Mail")
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
    hMsgbox = FindWindow("#32770", "America Online")
Loop Until hTree <> 0 Or hMsgbox <> 0

If hMsgbox <> 0 Then
    Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
    OpenOld25 = False
Else
    OpenOld25 = True
End If

End Function

Private Function OpenSent25() As Boolean
Dim hMain&, hTree&, hMsgbox&

Call RunMenuByString("Check Mail You've &Sent")
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Sent Mail")
    hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
    hMsgbox = FindWindow("#32770", "America Online")
Loop Until hTree <> 0 Or hMsgbox <> 0

If hMsgbox <> 0 Then
    Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)
    OpenSent25 = False
Else
    OpenSent25 = True
End If

End Function

Private Function Room25() As Long
Dim hMain&, hView&, hEdit&, hListbox&, hIcon&

Do
    hMain = FindWindowEx(MDI, hMain, "AOL Child", vbNullString)
    hView = FindWindowEx(hMain, 0, "_AOL_View", vbNullString)
    hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    hListbox = FindWindowEx(hMain, 0, "_AOL_Listbox", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Loop Until (hIcon <> 0 And hListbox <> 0 And hEdit <> 0 And hView <> 0) Or hMain = 0
Room25 = hMain
End Function

Private Sub SendMail25(sWho As String, sSubject As String, sText As String, Optional sCarbonCopy As String)
Dim hMain&, hEdit&, hIcon&, hMsgbox&

Call RunMenuByString("&Compose Mail")
Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Compose Mail")
    hEdit = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Loop Until hMain <> 0 And hIcon <> 0

Call SetText(hEdit, sWho)
    hEdit = NextOfClass(hEdit)

If IsMissing(sCarbonCopy) = True Then
    Call SetText(hEdit, sCarbonCopy)
Else
    hEdit = NextOfClass(hEdit)
End If

Call SetText(hEdit, sSubject)
hEdit = NextOfClass(hEdit)
Call SetText(hEdit, sText)
Call Icon(hIcon)

Do
    DoEvents
    hMsgbox = FindWindow("#32770", "America Online")
Loop Until hMsgbox <> 0 Or IsWindow(hEdit) = 0
Call PostMessage(hMsgbox, WM_CLOSE, 0, 0)

End Sub

Private Sub SetPreferences25()
Dim hMain&, hIcon&, hPrefs, hBox&
Dim nNext%

Call RunMenuByString("Set Preferences")

Do
    DoEvents
    hMain& = FindWindowEx(MDI, 0&, "AOL Child", "Preferences")
    hIcon& = FindWindowEx(hMain, 0&, "_AOL_Icon", vbNullString)
    For nNext = 1 To 5
        hIcon& = FindWindowEx(hMain, hIcon&, "_AOL_Icon", vbNullString)
    Next nNext
Loop Until hIcon <> 0

Call Icon(hIcon)

Do
    DoEvents
    hPrefs = FindWindow("_AOL_Modal", vbNullString)
    hBox = FindWindowEx(hPrefs, 0, "_AOL_Button", vbNullString)
Loop Until hBox <> 0

Call PostMessage(hMain, WM_CLOSE, 0, 0)
Call CheckIt(hBox, 0)
hBox = NextOfClass(hBox)
Call CheckIt(hBox, 1)
hBox = NextOfClass(hBox)
Call CheckIt(hBox, 0)
hBox = NextOfClass(hBox)
Call CheckIt(hBox, 1)
hBox = NextOfClass(hBox)
hBox = NextOfClass(hBox)

Do
    DoEvents
    Call SendMessage(hBox, WM_KEYDOWN, VK_SPACE, 0)
    Call SendMessage(hBox, WM_KEYUP, VK_SPACE, 0)
Loop Until IsWindow(hBox) = 0


End Sub

Private Function UserName25() As String
Dim hChild As Long
Dim sText As String
Do
    DoEvents
    hChild = FindWindowEx(MDI, hChild, "AOL Child", vbNullString)
    sText = GetText(hChild)
    If (Left(sText, Len("Welcome,")) = "Welcome,") And (Right(sText, 1) = "!") Then
        sText = Mid(sText, Len("welcome, "), Len(sText) - Len("welcome, "))
        Exit Do
    Else
        GoTo lblNext
    End If
lblNext:
Loop Until hChild = 0
UserName25 = Trim(sText)
End Function

Private Function WriteMail25(sWho As String, sSubject As String, sMessage As String, Optional lpBox As ListBox, Optional lpDeadBox As ListBox) As Integer
Dim hMain&, hText&
Dim hIcon&, nCount%

hMain = FindWindowEx(MDI, 0, "AOL Child", "Compose Mail")
If hMain = 0 Then Call RunMenuByString("Compose Mail")
Call LockWindowUpdate(MDI)

Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Compose Mail")
    hText = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
    hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)
Loop Until hText <> 0

If IsMissing(lpBox) = True Then
    sWho = MailList(lpBox, True)
End If

Call SetText(hText, sWho)
hText = NextOfClass(hText)
hText = NextOfClass(hText)
Call SetText(hText, sSubject)
hText = NextOfClass(hText)
Call SetText(hText, sMessage)

Call Icon(hIcon)
Call LockWindowUpdate(0)

Do
    DoEvents
    hMain = FindWindowEx(MDI, 0, "AOL Child", "Error")
    nCount = nCount + 1
    If hMain <> 0 And IsMissing(lpBox) = True Then
        Call Errored(lpBox, lpDeadBox)
        hMain = FindWindowEx(MDI, 0, "AOL Child", "Compose Mail")
        hText = FindWindowEx(hMain, 0, "_AOL_Edit", vbNullString)
        Call SetText(hText, MailList(lpBox, True))
    Else
        WriteMail25 = -1
        Exit Function
    End If
Loop
    





End Function

Sub Button(hButton As Long)
Call SendMessage(hButton, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(hButton, WM_KEYUP, VK_SPACE, 0)
End Sub

'**********************************************************************************

Function UserName() As String
If Is25 Then
    UserName = UserName25
Else
    UserName = UserName40
End If

End Function

Function IM(sWho As String, sText As String) As Integer
If Is25 Then
    IM = IM25(sWho, sText)
Else
    IM = IM40(sWho, sText)
End If
End Function

Sub ChatSend(sText As String)
If Is25 Then
    ChatSend25 sText
Else
    ChatSend40 sText
End If
End Sub

Function CountFlash() As Integer
If Is25 Then
    CountFlash = CountFlash25
Else
    CountFlash = CountFlash40
End If
End Function

Function CountNew() As Integer
If Is25 Then
    CountNew = CountNew25
Else
    CountNew = CountNew40
End If
End Function

Function CountOld() As Integer
If Is25 Then
    CountOld = CountOld25
Else
    CountOld = CountOld40
End If
End Function

Function CountSent() As Integer
If Is25 Then
    CountSent = CountSent25
Else
    CountSent = CountSent40
End If
End Function

Function FindMail() As Long
If Is25 Then
    FindMail = FindMail25
Else
    FindMail = FindMail40
End If
End Function

Function FindFwd() As Long
If Is25 Then
    FindFwd = FindFwd25
Else
    FindFwd = FindFwd40
End If
End Function

Function FwdFlash(lpBox As ListBox, nIndex As Integer, sMessage As String, lpDeadBox As ListBox) As Integer
If Is25 Then
    FwdFlash = FwdFlash25(lpBox, nIndex, sMessage, lpDeadBox)
Else
    FwdFlash = FwdFlash40(lpBox, nIndex, sMessage, lpDeadBox)
End If
End Function

Function FwdNew(lpBox As ListBox, nIndex As Integer, sMessage As String, lpDeadBox As ListBox) As Integer
If Is25 Then
    FwdNew = FwdNew25(lpBox, nIndex, sMessage, lpDeadBox)
Else
    FwdNew = FwdNew40(lpBox, nIndex, sMessage, lpDeadBox)
End If
End Function

Sub Ghost(bGhosting As Boolean)
If Is25 Then
    Ghost25 bGhosting
Else
    Ghost40 bGhosting
End If
End Sub

Sub Incoming²Box(lpBox As ListBox)
If Is25 Then
    Call Incoming²Box25(lpBox)
Else
    Call Incoming²Box40(lpBox)
End If

End Sub

Sub New²Box(lpBox As ListBox)
If Is25 Then
    Call New²Box25(lpBox)
Else
    Call New²Box40(lpBox)
End If

End Sub

Sub OpenFlash()
If Is25 Then
    OpenFlash25
Else
    OpenFlash40
End If
End Sub

Sub OpenNew()
If Is25 Then
    OpenNew25
Else
    OpenNew40
End If
End Sub

Sub OpenOld()
If Is25 Then
    OpenOld25
Else
    OpenOld40
End If
End Sub

Sub OpenSent()
If Is25 Then
    OpenSent25
Else
    OpenSent40
End If
End Sub

Function Room() As Long
If Is25 Then
    Room = Room25
Else
    Room = Room40
End If
End Function

Sub SendMail(sWho As String, sSubject As String, sText As String, Optional CarbonCopy As String)
If Is25 Then
    If IsMissing(CarbonCopy) = True Then
        SendMail25 sWho, sSubject, sText, CarbonCopy
    Else
        SendMail40 sWho, sSubject, sText
    End If
Else
    If IsMissing(CarbonCopy) = True Then
        SendMail25 sWho, sSubject, sText, CarbonCopy
    Else
        SendMail40 sWho, sSubject, sText
    End If
End If
End Sub

Sub SetPreferences()
If Is25 Then
    SetPreferences25
Else
    SetPreferences40
End If
End Sub

Sub WriteMail(sWho As String, sSubject As String, sMessage As String, Optional lpBox As ListBox)
If Is25 Then
    If IsMissing(lpBox) = False Then
        WriteMail25 sWho, sSubject, sMessage, lpBox
    Else
        WriteMail40 sWho, sSubject, sMessage
    End If
Else
    If IsMissing(lpBox) = False Then
        WriteMail40 sWho, sSubject, sMessage, lpBox
    Else
        WriteMail40 sWho, sSubject, sMessage
    End If
End If
End Sub


Sub Keyword(sWord As String)
If Is25 Then
    Keyword25 sWord
Else
    Keyword40 sWord
End If
End Sub

Function KillModal() As Boolean
If Is25 Then
    KillModal = KillModal25
Else
    KillModal = KillModal40
End If
End Function

Sub KeepAsNew(nIndex As Integer)
If Is25 Then
    KeepAsNew25 nIndex
Else
    KeepAsNew40 nIndex
End If
End Sub

Sub DeleteFlash25(nIndex As Integer)
Dim hMain&, hTree&, hIcon&, hMsgbox&
Dim hCount&

hMain& = FindWindowEx(MDI&, hMain&, "AOL Child", "Incoming FlashMail")
hTree& = FindWindowEx(hMain&, 0&, "_AOL_Tree", vbNullString)
hIcon = FindWindowEx(hMain&, 0&, "_AOL_Button", "Delete")
If hTree = 0 Then Exit Sub

hCount = CountFlash
Call PostMessage(hTree, (WM_USER + 7), nIndex, 0)
Call Button(hIcon)

Do
    DoEvents
    hMsgbox = FindWindow("#32770", "America Online")
    hMsgbox = FindWindowEx(hMsgbox, 0, "Button", "&Yes")
    If hMsgbox <> 0 Then
        Do
            DoEvents
            Call Button(hMsgbox)
        Loop Until IsWindow(hMsgbox) = 0
        Exit Sub
    End If
Loop Until CountFlash <> hCount

End Sub

Sub DeleteFlash40(nIndex As Integer)
Dim hMain&, hTree&, hIcon&, hMsgbox&
Dim nCount As Integer

hMain = FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail")
hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)

If hTree = 0 Then Call OpenFlash

hMain = FindWindowEx(MDI, 0, "AOL Child", "Incoming/Saved Mail")
hTree = FindWindowEx(hMain, 0, "_AOL_Tree", vbNullString)
hIcon = FindWindowEx(hMain, 0, "_AOL_Icon", vbNullString)

Do
    DoEvents
    hIcon = FindWindowEx(hMain, hIcon, "_AOL_Icon", vbNullString)
Loop Until GetText(hIcon) = "     Delete     "

nCount = CountFlash
Call PostMessage(hTree, LB_SETCURSEL, nIndex, 0)
Call Icon(hIcon)

Do
    DoEvents
    hMsgbox = FindWindow("#32770", "America Online")
    hMsgbox = FindWindowEx(hMsgbox, 0, "Button", "&Yes")
    If hMsgbox <> 0 Then
        Do
            DoEvents
            Call Button(hMsgbox)
        Loop Until IsWindow(hMsgbox) = 0
        Exit Sub
    End If
Loop Until CountFlash <> nCount

End Sub

Sub DeleteFlash(nIndex As Integer)
If Is25 Then
    DeleteFlash25 nIndex
Else
    DeleteFlash40 nIndex
End If
End Sub

'________________________________________________________
Sub Loadlistbox(Directory As String, TheList As ListBox) '|
    Dim MyString As String                               '|
    On Error Resume Next                                 '|
    Open Directory$ For Input As #1                      '|
    While Not EOF(1)                                     '|
        Input #1, MyString$            'dos' material    '|
        DoEvents                                         '|
        TheList.AddItem MyString$                        '|
    Wend                                                 '|
    Close #1                                             '|
End Sub                                                  '|
                                                         '|
Sub SaveListBox(Directory As String, TheList As ListBox) '|
    Dim SaveList As Long                                 '|
    On Error Resume Next                                 '|
    Open Directory$ For Output As #1                     '|
    For SaveList& = 0 To TheList.ListCount - 1           '|
        Print #1, TheList.List(SaveList&)                '|
    Next SaveList&                                       '|
    Close #1                                             '|
End Sub                                                  '|
'________________________________________________________'|
Function MsgToLong() As Long
'returns that box that comes up whenever your text is to long
Dim hBox&, hText&, sText$
hBox& = FindWindow("#32770", vbNullString)
hText& = FindWindowEx(hBox&, 0&, "Static", vbNullString)
hText& = FindWindowEx(hBox&, hText, "Static", vbNullString)
If GetText(hText) = "Message is too long or too complex" Then
    MsgToLong = hBox
Else
    MsgToLong = 0
End If
End Function

Sub KillWait()
Call RunMenuByString("&About America Online")
Do
    DoEvents
Loop Until KillModal = True
End Sub
