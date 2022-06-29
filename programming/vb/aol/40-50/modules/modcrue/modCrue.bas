Attribute VB_Name = "modCrue"
'addon features were coded for AOL 4.0, AIM 2.0.813, mIRC 5.6, and Virtual Places v2.1.1 Exc
'people please, please learn something from this.
'
'Special Notes:
'   CharLower and CharUpper are faster than LCase and UCase
'   "If Len(AString$) = 0 Then..."  ·-is faster than-· "If AString$ = "" Then..."
'   for the most part, an API sub or function is just as fast, if not faster than
'       a sub or function built into or written with visual basic
'
'Rare Subs/Functions:
'   GetPixelColor, FileScanEMAIL, AddServerList
'   SortFlashmailBySubject, SortFlashmailByDate, SortFlashmailBySender
'   MixFlashmail, ReverseFlashmail, PictureToHTML
'   CountUniqueColors, SortArray, ***PLUS MANY MORE!!!***
'
'Contacting Crüe: (feel free to do so)
'   EMail:  CrueIzMe@Hotmail.com
'   AIM:    i am crue
'   VP:     crueizme
'   WWW:    http://come.to/cruelair
Option Explicit
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long


Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC
Public Const WM_LBUTTONDBLCLK = &H203

Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 6
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5

Public Const VK_SPACE = &H20
Public Const VK_RETURN = &HD
Public Const VK_DOWN = &H28
Public Const VK_RIGHT = &H27
Public Const VK_UP = &H26

Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_GETITEMDATA = &H199

Public Const EM_GETLINECOUNT = &HBA

Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0

Public Const SRCCOPY = &HCC0020
Public Const STRETCH_DELETESCANS = 3

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3

Public Const CB_GETCOUNT = &H146
Public Const CB_SETCURSEL = &H14E

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

Public Type COLORRGB
    Red As Long
    Green As Long
    Blue As Long
End Type
Public Sub AcceptAIM(ByVal AcceptIt As Boolean)
    'this sub will either accept or decline a message
    'sent to you by somebody using AIM
    'Example of use:
    '   Call AcceptAIM(True)
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Dim AOLChild As Long, AOLStatic As Long, AOLIcon As Long
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
        If InStr(1, GetText(AOLStatic&), "Would you like to accept?", 1) <> 0 Then
            AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
            If AcceptIt = False Then
                AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
            End If
            Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
            Exit Sub
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
End Sub
Public Sub AddDownloadsToList(ByVal TheList As ListBox)
    'this sub will add the downloads in your download manager to a listbox
    'Example of use:
    '   Call AddDownloadsToList(List1)
    If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
    Dim AOLTree As Long
    AOLTree& = FindDownloadManager&
    If AOLTree& = 0& Then
        Call LoadDownloadManager
        Do
            DoEvents
            AOLTree& = FindDownloadManager&
        Loop Until AOLTree& <> 0&
    End If
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_Tree", vbNullString)
    Dim TheCount As Long, TheDownload As String, TheLen As Long
    Dim LoopThrough As Long, TheINSTR As Integer
    TheCount& = CountListItems(AOLTree&)
    If TheCount& = 0& Then Exit Sub
    TheCount& = TheCount& - 1
    For LoopThrough& = 0& To TheCount&
        TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheDownload$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(AOLTree&, LB_GETTEXT, LoopThrough&, TheDownload$)
        TheINSTR% = InStr(1, TheDownload$, Chr(9), 1)
        TheDownload$ = Mid(TheDownload$, 1 + TheINSTR%, InStr(1 + TheINSTR%, TheDownload$, Chr(9), 1) - 1)
        TheList.AddItem TheDownload$
    Next LoopThrough&
End Sub

Public Sub AddServerList(ByVal TheListbox As ListBox)
    'this sub will add the items from a server list to a listbox
    'Example of use:
    '   List1.Visible = False
    '   Call AddServerList(List1)
    '   List1.Visible = True
    Dim TheWin As Long, NumLines As Long, theline As String
    Dim LoopThrough As Long, TheList As String
    Dim DummyText As String
    TheWin& = FindReadWindow&
    If TheWin& = 0& Then Exit Sub
    TheWin& = FindWindowEx(TheWin&, 0&, "RICHCNTL", vbNullString)
    NumLines& = LineCountWindow(TheWin&)
    TheList$ = GetText(TheWin&)
    If Len(TheList$) = 0 Then Exit Sub
    For LoopThrough& = 6 To (NumLines&)
        theline$ = LineFromString(TheList$, LoopThrough&)
        DummyText$ = Replace(theline$, " ", "")
        DummyText$ = Replace(DummyText$, "(", "")
        DummyText$ = Replace(DummyText$, "*", "")
        DummyText$ = Replace(DummyText$, ".", "")
        DummyText$ = Replace(DummyText$, "#", "")
        DummyText$ = Replace(DummyText$, "[", "")
        DummyText$ = Replace(DummyText$, "•", "")
        DummyText$ = Replace(DummyText$, "¤", "")
        DummyText$ = Replace(DummyText$, "-", "")
        If IsNumeric(Left$(DummyText$, 1)) = True Then
            TheListbox.AddItem theline$
        End If
    Next LoopThrough&
    DoEvents
End Sub




Public Sub ChatIgnore(ByVal IgnoreName As String, ByVal ExactSN As Boolean, ByVal IgnoreThem As Boolean)
    'this sub will ignore somebody in a AOL chatroom
    'if ExactSN is true, then the sub will search for the person
    'who has the EXACT same name as IgnoreName, case sensitive
    'if ExactSN is false, then the sub will remove all spaces
    'from the names it searches, and the search is not case sensitive
    'if IgnoreThem is true, they will be ignored.  if its false, they will be unignored
    'Example of use:
    '   Call ChatIgnore("Steve Case", True, True)
    '   Call ChatIgnore("stevecase", False, True)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String, MDIClient As Long
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long, TheINSTR As Integer
    Dim RoomList As Long, mThread As Long, InfoWin As Long, AOLCheckbox As Long
    Dim CheckVal As Long
    Room& = FindChatRoom&
    If Room& = 0& Then Exit Sub
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    If ExactSN = False Then IgnoreName = CharLower(Replace(IgnoreName, " ", ""))
    RoomList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call GetWindowThreadProcessId(RoomList&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(RoomList&, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, Chr(0))
            itmHold& = SendMessage(RoomList&, LB_GETITEMDATA, ByVal CLng(Index&), 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes&)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, Chr(0))
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            TheINSTR% = InStr(1, ScreenName$, Chr(0), 1)
            ScreenName$ = Left$(ScreenName$, TheINSTR% - 1)
            If (ExactSN = True And (ScreenName$ = IgnoreName$)) Or (ExactSN = False And (CharLower(Replace(ScreenName$, " ", ""))) = IgnoreName$) Then
                Call PostMessage(RoomList&, LB_SETCURSEL, Index&, 0&)
                DoEvents
                Call PostMessage(RoomList&, WM_LBUTTONDBLCLK, 0&, 0&)
                Do
                    DoEvents
                    InfoWin& = FindWindowEx(MDIClient&, 0&, "AOL Child", ScreenName$)
                    AOLCheckbox& = FindWindowEx(InfoWin&, 0&, "_AOL_Checkbox", vbNullString)
                Loop Until IsWindowVisible(AOLCheckbox&) = 1
                DoEvents
                CheckVal& = SendMessage(AOLCheckbox&, BM_GETCHECK, 0&, 0&)
                If (CheckVal& = 1 And IgnoreThem = True) Or (CheckVal& = 0& And IgnoreThem = False) Then Exit Sub
                Do
                    Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
                    Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
                    DoEvents
                    DoEvents
                    CheckVal& = SendMessage(AOLCheckbox&, BM_GETCHECK, 0&, 0&)
                Loop Until (CheckVal& = 1 And IgnoreThem = True) Or (CheckVal& = 0& And IgnoreThem = False)
                DoEvents
                Call PostMessage(InfoWin&, WM_CLOSE, 0&, 0&)
                DoEvents
                Exit Sub
            End If
        Next Index&
    End If
End Sub

Public Sub CloseAllAOLChilds()
    'this sub will close all of the AOL Child windows
    'Example of use:
    '   Call CloseAllAOLChilds
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Dim AOLChild As Long
FindTheChild:
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    If AOLChild& = 0& Then Exit Sub
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
    DoEvents
    GoTo FindTheChild
End Sub

Public Sub CloseAllMails()
    'this sub will close all of aol's email windows
    'Example of use:
    '   Call CloseAllMails
    Dim ReadWin As Long, WriteWin As Long, ForwardWin As Long
    ReadWin& = FindReadWindow&
    WriteWin& = FindWriteMail&
    ForwardWin& = FindForwardWindow&
    Do Until ReadWin& = 0& And WriteWin& = 0& And ForwardWin& = 0&
        Call PostMessage(ReadWin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(WriteWin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(ForwardWin&, WM_CLOSE, 0&, 0&)
        DoEvents
        ReadWin& = FindReadWindow&
        WriteWin& = FindWriteMail&
        ForwardWin& = FindForwardWindow&
    Loop
End Sub
Public Function CountChr(ByVal TheString As String, ByVal ChrToCount As String) As Long
    'this function counts the number of times a string appears in a string
    'Example of use:
    '   MsgBox CountChr("Hello, how are you today?", "o")
    '   MsgBox CountChr("Who What When Where Why?", "Wh")
    Dim Counter As Long, TheINSTR As Long, TheLen As Long
    TheINSTR& = InStr(1, TheString$, ChrToCount$, 1)
    If TheINSTR& = 0& Then
        CountChr& = 0&
        Exit Function
    End If
    Counter& = 1&
    TheLen& = Len(ChrToCount$)
    Do
        TheINSTR& = InStr(TheLen& + TheINSTR&, TheString$, ChrToCount$, 1)
        If TheINSTR& = 0& Then Exit Do
        Counter& = 1 + Counter&
    Loop
    CountChr& = Counter&
End Function

Public Function CountListItems(ByVal TheList As Long) As Long
    'this function counts the number of items in a listbox
    'Example of use:
    '   Dim AOLList As Long
    '   AOLList& = FindWindowEx(FindChatRoom&, 0&, "_AOL_Listbox", vbNullString)
    '   MsgBox CountListItems(AOLList&)
    CountListItems& = SendMessage(TheList&, LB_GETCOUNT, 0&, 0&)
End Function

Public Sub DeleteDuplicateFlashmails()
    'this sub will delete all duplicate flashmails
    'Example of use:
    '   Call DeleteDuplicateFlashmails
    Dim MBox As Long
    MBox& = FindFlashMailBox&
    If MBox& = 0& Then
        If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Call LoadFlashmail
        Do
            DoEvents
            MBox& = FindFlashMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    Dim AOLTree As Long, AOLIcon As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_Tree", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    Dim TheINSTR As Integer, TheINSTR2 As Integer, TheLen As Long
    Dim TheSender As String, TheSubject As String, MailCount As Long
    Dim TheSender2 As String, TheSubject2 As String
    MailCount& = SendMessage(AOLTree&, LB_GETCOUNT, 0&, 0&)
    If MailCount& < 2 Then Exit Sub
    Dim LoopThrough As Long, LoopThrough2 As Long
    MailCount& = MailCount& - 1
    For LoopThrough& = 0& To MailCount&
        TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheSender$ = String(1 + TheLen&, 0&)
        TheSubject$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(AOLTree&, LB_GETTEXT, LoopThrough&, TheSender$)
        TheINSTR% = InStr(1, TheSender$, Chr(9), 1)
        TheINSTR2% = InStr(1 + TheINSTR%, TheSender$, Chr(9), 1)
        TheSubject$ = Right(TheSender$, CLng(Len(TheSender$) - TheINSTR2%))
        TheSender$ = Mid(TheSender$, 1 + TheINSTR%, TheINSTR2% - TheINSTR% - 1)
        For LoopThrough2& = MailCount& To LoopThrough& Step -1
            DoEvents
            TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, LoopThrough2&, 0&)
            TheSender2$ = String(1 + TheLen&, 0&)
            TheSubject2$ = String(1 + TheLen&, 0&)
            Call SendMessageByString(AOLTree&, LB_GETTEXT, LoopThrough2&, TheSender2$)
            TheINSTR% = InStr(1, TheSender2$, Chr(9), 1)
            TheINSTR2% = InStr(1 + TheINSTR%, TheSender2$, Chr(9), 1)
            TheSubject2$ = Right(TheSender2$, CLng(Len(TheSender2$) - TheINSTR2%))
            TheSender2$ = Mid(TheSender2$, 1 + TheINSTR%, TheINSTR2% - TheINSTR% - 1)
            If (TheSubject$ = TheSubject2$) And (TheSender$ = TheSender2$) Then
                Call SendMessage(AOLTree&, LB_SETCURSEL, LoopThrough&, 0&)
                DoEvents
                Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                MailCount& = MailCount& - 1
            End If
        Next LoopThrough2&
    Next LoopThrough&
End Sub

Public Sub DeleteSentMail()
    'this sub will delete all the mail from your sent mail box
    'Example of use:
    '   Call DeleteSentMail
    If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
    Dim MBox As Long, AOLIcon As Long
    MBox& = FindMailBox&
    Dim CPos As POINTAPI, Mnu As Long
    AOLIcon& = GetAOLToolbarIcon(3)
    Call GetCursorPos(CPos)
    Call SetCursorPos(CPos.X, CPos.Y)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        Mnu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(Mnu&) = 1
    DoEvents
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
    DoEvents
    Call SetCursorPos(CPos.X, CPos.Y)
    Do
        DoEvents
        MBox& = FindMailBox&
    Loop Until MBox& <> 0&
    Call Sleep(2000)
    DoEvents
    Dim AOLTree As Long, TabP As Long, TabC As Long
    TabC& = FindWindowEx(MBox&, 0&, "_AOL_TabControl", vbNullString)
    TabP& = FindWindowEx(TabC&, 0&, "_AOL_TabPage", vbNullString)
    TabP& = FindWindowEx(TabC&, TabP&, "_AOL_TabPage", vbNullString)
    TabP& = FindWindowEx(TabC&, TabP&, "_AOL_TabPage", vbNullString)
    AOLTree& = FindWindowEx(TabP&, 0&, "_AOL_Tree", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    Dim LoopThrough As Long
DeleteTheMails:
    Call PostMessage(AOLTree&, LB_SETCURSEL, 0&, 0&)
    DoEvents
    For LoopThrough& = 0 To CountListItems(AOLTree&) - 1
        Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
    Next LoopThrough&
    If CountListItems(AOLTree&) <> 0& Then GoTo DeleteTheMails
End Sub

Public Sub EditProfile(ByVal MemberName As String, ByVal Location As String, ByVal Birthdate As String, ByVal MaritalStatus As String, ByVal Hobbies As String, ByVal ComputersUsed As String, ByVal Occupation As String, ByVal PersonalQuote As String)
    'this sub will edit your AOL profile (sex will be set to "no responce")
    'Example of use:
    '   Call EditProfile("<<u>font color="#0000FF" face="Tahoma">My Name", "My Location", "My Birthday", "My Marital Status", "My Hobbies", "My Computer", "My Job", "My Personal Quote")
    Dim AOLIcon As Long
    AOLIcon& = GetAOLToolbarIcon(6)
    If AOLIcon& = 0& Then Exit Sub
    Dim Mnu As Long, CPos As POINTAPI
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        Mnu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(Mnu&) = 1
    DoEvents
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
    DoEvents
    Call SetCursorPos(CPos.X, CPos.Y)
    Dim ProfileWin As Long, NameEdit As Long, LocEdit As Long, BDayEdit As Long
    Dim NoResponce As Long, MaritalEdit As Long, HobbiesEdit As Long
    Dim CpuEdit As Long, JobEdit As Long, QuoteEdit As Long
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Do
        DoEvents
        ProfileWin& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Edit Your Online Profile")
        NameEdit& = FindWindowEx(ProfileWin&, 0&, "_AOL_Edit", vbNullString)
        LocEdit& = FindWindowEx(ProfileWin&, NameEdit&, "_AOL_Edit", vbNullString)
        BDayEdit& = FindWindowEx(ProfileWin&, LocEdit&, "_AOL_Edit", vbNullString)
        MaritalEdit& = FindWindowEx(ProfileWin&, BDayEdit&, "_AOL_Edit", vbNullString)
        HobbiesEdit& = FindWindowEx(ProfileWin&, MaritalEdit&, "_AOL_Edit", vbNullString)
        CpuEdit& = FindWindowEx(ProfileWin&, HobbiesEdit&, "_AOL_Edit", vbNullString)
        JobEdit& = FindWindowEx(ProfileWin&, CpuEdit&, "_AOL_Edit", vbNullString)
        QuoteEdit& = FindWindowEx(ProfileWin&, JobEdit&, "_AOL_Edit", vbNullString)
        AOLIcon& = FindWindowEx(ProfileWin&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(ProfileWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        NoResponce& = FindWindowEx(ProfileWin&, 0&, "_AOL_Checkbox", vbNullString)
        NoResponce& = FindWindowEx(ProfileWin&, NoResponce&, "_AOL_Checkbox", vbNullString)
        NoResponce& = FindWindowEx(ProfileWin&, NoResponce&, "_AOL_Checkbox", vbNullString)
    Loop Until (IsWindowVisible(NameEdit&) = 1) And (IsWindowVisible(LocEdit&) = 1) And (IsWindowVisible(BDayEdit&) = 1) And (IsWindowVisible(MaritalEdit&) = 1) And (IsWindowVisible(HobbiesEdit&) = 1) And (IsWindowVisible(CpuEdit&) = 1) And (IsWindowVisible(JobEdit&) = 1) And (IsWindowVisible(QuoteEdit&) = 1) And (IsWindowVisible(NoResponce&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    DoEvents
    Call SetText(NameEdit&, MemberName$)
    Call SetText(LocEdit&, Location$)
    Call SetText(BDayEdit&, Birthdate$)
    Call SetText(MaritalEdit&, MaritalStatus$)
    Call SetText(HobbiesEdit&, Hobbies$)
    Call SetText(CpuEdit&, ComputersUsed$)
    Call SetText(JobEdit&, Occupation$)
    Call SetText(QuoteEdit&, PersonalQuote$)
    Call SendMessage(NoResponce&, BM_SETCHECK, 1, 0)
    DoEvents
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    ProfileWin& = FindWindow("_AOL_Modal", vbNullString)
    If ProfileWin& <> 0& Then
        AOLIcon& = FindWindowEx(ProfileWin&, 0&, "_AOL_Icon", vbNullString)
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do
        DoEvents
        ProfileWin& = FindWindow("#32770", "America Online")
        AOLIcon& = FindWindowEx(ProfileWin&, 0&, "Button", vbNullString)
    Loop Until IsWindowVisible(AOLIcon&) = 1
    Call PostMessage(AOLIcon&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub ExitAIM()
    'this sub will exit AIM
    'Example of use:
    '   Call ExitAIM
    Call RunMenuByString(FindAIM&, "E&xit")
End Sub

Public Sub FlashmailDatesToList(ByVal TheList As ListBox)
    'this sub will add the date of every mail in your flashmail box
    'to a listbox
    'Example of use:
    '   Call FlashmailDatesToList(List1)
    Dim TheTree As Long
    TheTree& = FindFlashMailBox&
    If TheTree& = 0& Then
        If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Call LoadFlashmail
        Do
            DoEvents
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
            TheTree& = FindFlashMailBox&
        Loop Until TheTree& <> 0&
        Call Sleep(2000)
    End If
    Dim MailCount As Long, TheLen As Long, TheINSTR As Integer
    Dim LoopThrough As Long, TheDate As String
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_Tree", vbNullString)
    MailCount& = SendMessage(TheTree&, LB_GETCOUNT, 0&, 0&)
    For LoopThrough& = 0 To MailCount& - 1
        TheLen& = SendMessage(TheTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheDate$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(TheTree&, LB_GETTEXT, LoopThrough&, TheDate$)
        TheINSTR% = InStr(1, TheDate$, Chr(9), 1)
        TheDate$ = Left(TheDate$, TheINSTR% - 1)
        TheList.AddItem Replace(TheDate$, Chr(0), "")
    Next LoopThrough&
End Sub

Public Sub FlashmailSendersToList(ByVal TheList As ListBox)
    'this sub will add the sender of every mail in your flashmail box
    'to a listbox
    'Example of use:
    '   Call FlashmailSubjectsToList(List1)
    Dim TheTree As Long
    TheTree& = FindFlashMailBox&
    If TheTree& = 0& Then
        If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Call LoadFlashmail
        Do
            DoEvents
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
            TheTree& = FindFlashMailBox&
        Loop Until TheTree& <> 0&
        Call Sleep(2000)
    End If
    Dim MailCount As Long, TheLen As Long, TheINSTR As Integer
    Dim LoopThrough As Long, TheSender As String, TheINSTR2 As Integer
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_Tree", vbNullString)
    MailCount& = SendMessage(TheTree&, LB_GETCOUNT, 0&, 0&)
    For LoopThrough& = 0 To MailCount& - 1
        TheLen& = SendMessage(TheTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheSender$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(TheTree&, LB_GETTEXT, LoopThrough&, TheSender$)
        TheINSTR% = InStr(1, TheSender$, Chr(9), 1)
        TheINSTR2% = InStr(1 + TheINSTR%, TheSender$, Chr(9), 1)
        TheSender$ = Mid(TheSender$, 1 + TheINSTR%, TheINSTR2% - TheINSTR% - 1)
        TheList.AddItem Replace(TheSender$, Chr(0), "")
    Next LoopThrough&
End Sub
Public Sub FlashmailSubjectsToList(ByVal TheList As ListBox)
    'this sub will add the subject of every mail in your flashmail box
    'to a listbox
    'Example of use:
    '   Call FlashmailSubjectsToList(List1)
    Dim TheTree As Long
    TheTree& = FindFlashMailBox&
    If TheTree& = 0& Then
        If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Call LoadFlashmail
        Do
            DoEvents
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
            TheTree& = FindFlashMailBox&
        Loop Until TheTree& <> 0&
        Call Sleep(2000)
    End If
    Dim MailCount As Long, TheLen As Long, TheINSTR As Integer
    Dim LoopThrough As Long, TheSubject As String
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_Tree", vbNullString)
    MailCount& = SendMessage(TheTree&, LB_GETCOUNT, 0&, 0&)
    For LoopThrough& = 0 To MailCount& - 1
        TheLen& = SendMessage(TheTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheSubject$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(TheTree&, LB_GETTEXT, LoopThrough&, TheSubject$)
        TheINSTR% = InStr(1, TheSubject$, Chr(9), 1)
        TheINSTR% = InStr(1 + TheINSTR%, TheSubject$, Chr(9), 1)
        TheSubject$ = Right(TheSubject$, Len(TheSubject$) - TheINSTR%)
        TheList.AddItem Replace(TheSubject$, Chr(0), "")
    Next LoopThrough&
End Sub
Public Sub ForwardFlashMail(ByVal TheIndex As Long, ByVal ToWho As String, ByVal TheMessage As String, ByVal RemoveFWD As Boolean)
    'this sub will open a email from your flashmail box and forward it to somebody
    'it will also remove "Fwd: " from the subject if you want
    'remember, indexes start at zero
    'Example of use:
    '   Call ForwardFlashMail(0, "Steve Case", "Check this out.", True)
    Dim AOLWin As Long
    AOLWin& = FindWindow("AOL Frame25", vbNullString)
    If AOLWin& = 0& Then Exit Sub
    Dim MDIClient As Long
    MDIClient& = FindWindowEx(AOLWin&, 0&, "MDIClient", vbNullString)
    Dim MBox As Long
    MBox& = FindFlashMailBox&
    If MBox& = 0& Then
        Call LoadFlashmail
        Do
            DoEvents
            MBox& = FindFlashMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    Dim AOLTree As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_Tree", vbNullString)
    If CountListItems(AOLTree&) <= TheIndex& Then Exit Sub
    Call SendMessage(AOLTree&, LB_SETCURSEL, TheIndex&, 0&)
    DoEvents
    Dim AOLIcon As Long
    AOLIcon& = FindWindowEx(MBox&, 0&, "_AOL_Icon", vbNullString)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Dim TheLen As Long, TheSubject As String, TheINSTR As Integer
    TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, TheIndex&, 0&)
    TheSubject$ = String(1 + TheLen&, 0&)
    Call SendMessageByString(AOLTree&, LB_GETTEXT, TheIndex&, TheSubject$)
    TheINSTR% = InStr(1, TheSubject$, Chr(9), 1)
    TheINSTR% = InStr(1 + TheINSTR%, TheSubject$, Chr(9), 1)
    TheSubject$ = Right(TheSubject$, Len(TheSubject$) - TheINSTR%)
    Dim AOLChild As Long, Rich As Long
    Do
        DoEvents
        AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", TheSubject$)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Loop Until (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    Call RunMenuByString(AOLWin&, "S&top Incoming Text")
    DoEvents
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Dim FwdWin As Long, TheCaption As String
    Dim ToEdit As Long, SubjEdit As Long
    Do
        DoEvents
        FwdWin& = FindForwardWindow&
        TheCaption$ = GetCaption(FwdWin&)
        If Len(TheCaption$) > 5 Then
            TheCaption$ = Right(TheCaption$, Len(TheCaption$) - 5)
        End If
        ToEdit& = FindWindowEx(FwdWin&, 0&, "_AOL_Edit", vbNullString)
        SubjEdit& = FindWindowEx(FwdWin&, ToEdit&, "_AOL_Edit", vbNullString)
        SubjEdit& = FindWindowEx(FwdWin&, SubjEdit&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(FwdWin&, 0&, "RICHCNTL", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
    Loop Until (InStr(1, TheSubject$, TheCaption$, 1) <> 0) And (IsWindowVisible(ToEdit&) = 1) And (IsWindowVisible(SubjEdit&) = 1) And (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    DoEvents
    If RemoveFWD = True Then
        Call SendMessageByString(SubjEdit&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(SubjEdit&, WM_SETTEXT, 0&, TheCaption$)
    End If
    Call SendMessageByString(ToEdit&, WM_SETTEXT, 0&, ToWho$)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, TheMessage$)
    DoEvents
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        FwdWin& = FindForwardWindow&
        If FwdWin& = 0& Then Exit Do
        TheCaption$ = GetCaption(FwdWin&)
        TheCaption$ = Right(TheCaption$, Len(TheCaption$) - 5)
    Loop Until InStr(1, TheSubject$, TheCaption$, 1) = 0
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub ForwardNewMail(ByVal TheIndex As Long, ByVal ToWho As String, ByVal TheMessage As String, ByVal RemoveFWD As Boolean, ByVal KeepAsNew As Boolean)
    'this sub will open a email from your mailbox and forward it to somebody
    'it will also remove "Fwd: " from the subject if you want, and keep the mail as new if you want
    'remember, indexes start at zero
    'Example of use:
    '   Call ForwardNewMail(0, "Steve Case", "Check this out.", True, True)
    Dim AOLWin As Long
    AOLWin& = FindWindow("AOL Frame25", vbNullString)
    If AOLWin& = 0& Then Exit Sub
    Dim MDIClient As Long
    MDIClient& = FindWindowEx(AOLWin&, 0&, "MDIClient", vbNullString)
    Dim MBox As Long
    MBox& = FindMailBox&
    If MBox& = 0& Then
        Call LoadMailBox
        Do
            DoEvents
            MBox& = FindMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    Dim AOLTree As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_TabControl", vbNullString)
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_TabPage", vbNullString)
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_Tree", vbNullString)
    If CountListItems(AOLTree&) <= TheIndex& Then Exit Sub
    Call SendMessage(AOLTree&, LB_SETCURSEL, TheIndex&, 0&)
    DoEvents
    Dim AOLIcon As Long
    AOLIcon& = FindWindowEx(MBox&, 0&, "_AOL_Icon", vbNullString)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Dim TheLen As Long, TheSubject As String, TheINSTR As Integer
    TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, TheIndex&, 0&)
    TheSubject$ = String(1 + TheLen&, 0&)
    Call SendMessageByString(AOLTree&, LB_GETTEXT, TheIndex&, TheSubject$)
    TheINSTR% = InStr(1, TheSubject$, Chr(9), 1)
    TheINSTR% = InStr(1 + TheINSTR%, TheSubject$, Chr(9), 1)
    TheSubject$ = Right(TheSubject$, Len(TheSubject$) - TheINSTR%)
    Dim AOLChild As Long, Rich As Long
    Do
        DoEvents
        AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", TheSubject$)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Loop Until (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    Call RunMenuByString(AOLWin&, "S&top Incoming Text")
    DoEvents
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Dim FwdWin As Long, TheCaption As String
    Dim ToEdit As Long, SubjEdit As Long
    Do
        DoEvents
        FwdWin& = FindForwardWindow&
        TheCaption$ = GetCaption(FwdWin&)
        If Len(TheCaption$) > 5 Then
            TheCaption$ = Right(TheCaption$, Len(TheCaption$) - 5)
        End If
        ToEdit& = FindWindowEx(FwdWin&, 0&, "_AOL_Edit", vbNullString)
        SubjEdit& = FindWindowEx(FwdWin&, ToEdit&, "_AOL_Edit", vbNullString)
        SubjEdit& = FindWindowEx(FwdWin&, SubjEdit&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(FwdWin&, 0&, "RICHCNTL", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(FwdWin&, AOLIcon&, "_AOL_Icon", vbNullString)
    Loop Until (InStr(1, TheSubject$, TheCaption$, 1) <> 0) And (IsWindowVisible(ToEdit&) = 1) And (IsWindowVisible(SubjEdit&) = 1) And (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    DoEvents
    If RemoveFWD = True Then
        Call SendMessageByString(SubjEdit&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(SubjEdit&, WM_SETTEXT, 0&, TheCaption$)
    End If
    Call SendMessageByString(ToEdit&, WM_SETTEXT, 0&, ToWho$)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, TheMessage$)
    DoEvents
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        FwdWin& = FindForwardWindow&
        If FwdWin& = 0& Then Exit Do
        TheCaption$ = GetCaption(FwdWin&)
        TheCaption$ = Right(TheCaption$, Len(TheCaption$) - 5)
    Loop Until InStr(1, TheSubject$, TheCaption$, 1) = 0
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
    DoEvents
    If KeepAsNew = True Then
        AOLIcon& = FindWindowEx(MBox&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    End If
End Sub



Public Function GetTheParent(ByVal Win As Long, ByVal ParIdx As Long) As Long
    'this function gets a specific parent from a window
    'Example of use:
    '   MsgBox GetTheParent(FindChatRoom&, 1)
    Dim LoopThrough As Long, AWin As Long
    AWin& = Win&
    For LoopThrough& = 1& To ParIdx&
        AWin& = GetParent(AWin&)
    Next LoopThrough&
    GetTheParent& = AWin&
End Function
Public Function GetTheSibling(ByVal Win As Long, ByVal SibIdx As Long) As Long
    'this function gets a specific sibling from a window
    'Example of use:
    
    Dim AWin As Long, LoopThrough As Long, aWinClass As String
    aWinClass$ = GetClass(Win&)
    AWin& = Win&
    LoopThrough& = 0&
    Do
        AWin& = GetWindow(AWin&, GW_HWNDPREV)
        If GetClass(AWin&) = aWinClass$ Then
            LoopThrough& = 1& + LoopThrough&
        End If
    Loop Until (LoopThrough& = SibIdx& And (LoopThrough& <> 0&)) Or (AWin& = 0&)
    GetTheSibling& = AWin&
End Function
Public Sub IMIgnore(ByVal ScreenName As String, ByVal IgnoreThem As Boolean)
    'this will ignore or unignore somebody from IMs on AOL
    'Example of use:
    '   Call IMIgnore("i am crue", False)
    If IgnoreThem = True Then
        Call SendIM("$IM_Off " & ScreenName$, "http://come.to/cruelair")
    Else
        Call SendIM("$IM_On " & ScreenName$, "http://come.to/cruelair")
    End If
End Sub

Public Function IsUserOnline() As Boolean
    'this checks to see if the user is online with AOL
    'Example of use:
    '   Call MsgBox(IsUserOnline)
    If FindWelcomeScreen& = 0& Then
        IsUserOnline = False
    Else
        IsUserOnline = True
    End If
End Function
Public Sub KeepNewMailAsNew()
    'this sub will keep all the mail in your mailbox as new
    'the mailbox must be already open, otherwise this sub is pointless
    'Example of use:
    '   Call KeepNewMailAsNew
    Dim MBox As Long
    MBox& = FindMailBox&
    If MBox& = 0& Then Exit Sub
    Dim AOLIcon As Long, AOLTree As Long, LoopThrough As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_TabControl", vbNullString)
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_TabPage", vbNullString)
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_Tree", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(MBox&, AOLIcon&, "_AOL_Icon", vbNullString)
    For LoopThrough& = 0 To CountListItems(AOLTree&) - 1
        Call PostMessage(AOLTree&, LB_SETCURSEL, LoopThrough&, 0&)
        DoEvents
        Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
    Next LoopThrough&
End Sub
Public Function ListToMailString(ByVal TheList As ListBox, ByVal BCC As Boolean) As String
    'this function will convert a list of names into a string you can use
    'for email
    'Example of use:
    '   Call SendMail(ListToMailString(List1, True), "subject", "message")
    Dim LoopThrough As Integer, TempString As String
    TempString$ = ""
    If BCC = True Then
        For LoopThrough& = 0& To CountListItems(TheList.hWnd) - 1
            TempString$ = TempString$ & "(" & TheList.List(LoopThrough&) & ")"
        Next LoopThrough&
    Else
        For LoopThrough& = 0& To CountListItems(TheList.hWnd) - 1
            TempString$ = TempString$ & TheList.List(LoopThrough&) & ", "
        Next LoopThrough&
        TempString$ = Left(TempString$, Len(TempString$) - 2)
    End If
    ListToMailString$ = TempString$
End Function

Public Sub MailSendersToList(ByVal TheList As ListBox)
    'this sub will add the sender of every mail in your mail box
    'to a listbox
    'Example of use:
    '   Call MailSendersToList(List1)
    Dim TheTree As Long
    TheTree& = FindMailBox&
    If TheTree& = 0& Then
        If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Call LoadMailBox
        Do
            DoEvents
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
            TheTree& = FindMailBox&
        Loop Until TheTree& <> 0&
        Call Sleep(2000)
    End If
    Dim MailCount As Long, TheLen As Long, TheINSTR As Integer
    Dim LoopThrough As Long, TheSender As String, TheINSTR2 As Integer
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_TabControl", vbNullString)
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_TabPage", vbNullString)
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_Tree", vbNullString)
    MailCount& = SendMessage(TheTree&, LB_GETCOUNT, 0&, 0&)
    For LoopThrough& = 0 To MailCount& - 1
        TheLen& = SendMessage(TheTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheSender$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(TheTree&, LB_GETTEXT, LoopThrough&, TheSender$)
        TheINSTR% = InStr(1, TheSender$, Chr(9), 1)
        TheINSTR2% = InStr(1 + TheINSTR%, TheSender$, Chr(9), 1)
        TheSender$ = Mid(TheSender$, 1 + TheINSTR%, TheINSTR2% - TheINSTR% - 1)
        TheList.AddItem TheSender$
    Next LoopThrough&
End Sub

Public Sub MailSubjecsToList(ByVal TheList As ListBox)
    'this sub will add the subject of every mail in your mail box
    'to a listbox
    'Example of use:
    '   Call MailSubjectsToList(List1)
    Dim TheTree As Long
    TheTree& = FindMailBox&
    If TheTree& = 0& Then
        If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Call LoadMailBox
        Do
            DoEvents
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
            TheTree& = FindMailBox&
        Loop Until TheTree& <> 0&
        Call Sleep(2000)
    End If
    Dim MailCount As Long, TheLen As Long, TheINSTR As Integer
    Dim LoopThrough As Long, TheSubject As String
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_TabControl", vbNullString)
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_TabPage", vbNullString)
    TheTree& = FindWindowEx(TheTree&, 0&, "_AOL_Tree", vbNullString)
    MailCount& = SendMessage(TheTree&, LB_GETCOUNT, 0&, 0&)
    For LoopThrough& = 0 To MailCount& - 1
        TheLen& = SendMessage(TheTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheSubject$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(TheTree&, LB_GETTEXT, LoopThrough&, TheSubject$)
        TheINSTR% = InStr(1, TheSubject$, Chr(9), 1)
        TheINSTR% = InStr(1 + TheINSTR%, TheSubject$, Chr(9), 1)
        TheSubject$ = Right(TheSubject$, Len(TheSubject$) - TheINSTR%)
        TheList.AddItem TheSubject$
    Next LoopThrough&
End Sub

Public Sub MixFlashmail()
    'this sub will randomly mix your flashmail
    'Example of use:
    '   Call MixFlashmail
    If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
    Dim MBox As Long
    MBox& = FindFlashMailBox&
    If MBox& = 0& Then
        Call LoadFlashmail
        Do
            DoEvents
            MBox& = FindFlashMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    Dim AOLTree As Long, TheCount As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_Tree", vbNullString)
    TheCount& = CountListItems(AOLTree&)
    If TheCount& < 3 Then Exit Sub
    Dim LoopThrough As Long
    For LoopThrough& = 1 To TheCount&
        DoEvents
        DoEvents
        Call PostMessage(AOLTree&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call PostMessage(AOLTree&, LB_SETCURSEL, RandomNumber(TheCount&, False), 0&)
        DoEvents
        Call PostMessage(AOLTree&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(AOLTree&, WM_KEYUP, VK_DOWN, 0&)
        DoEvents
        Call PostMessage(AOLTree&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
        DoEvents
    Next LoopThrough&
End Sub
Public Function NameFromIM() As String
    'this will get the name from the top AOL IM
    'Example of use:
    '   MsgBox NameFromIM$
    Dim TheIM As Long
    TheIM& = FindIM&
    If TheIM& = 0& Then
        NameFromIM$ = ""
        Exit Function
    End If
    Dim TheINSTR As Integer, TheCaption As String
    TheCaption$ = GetCaption(TheIM&)
    TheINSTR% = InStr(1, TheCaption$, ": ", 1)
    NameFromIM$ = Mid(TheCaption$, 2 + TheINSTR%)
End Function

Public Function NoNumbers(ByVal TheString As String) As String
    'this function removes numbers from a string
    'Example of use:
    '   MsgBox NoNumbers("325h83a;348tthat")
    NoNumbers$ = Replace(TheString$, "0", "")
    NoNumbers$ = Replace(NoNumbers$, "1", "")
    NoNumbers$ = Replace(NoNumbers$, "2", "")
    NoNumbers$ = Replace(NoNumbers$, "3", "")
    NoNumbers$ = Replace(NoNumbers$, "4", "")
    NoNumbers$ = Replace(NoNumbers$, "5", "")
    NoNumbers$ = Replace(NoNumbers$, "6", "")
    NoNumbers$ = Replace(NoNumbers$, "7", "")
    NoNumbers$ = Replace(NoNumbers$, "8", "")
    NoNumbers$ = Replace(NoNumbers$, "9", "")
End Function

Public Function ParentCount(ByVal TheWindow As Long) As Long
    'this function counts the number of parents a window has
    'Example of use:
    '   MsgBox ParentCount(FindChatRoom&)
    Dim TheParent As Long, ParentCounter As Long
    ParentCounter& = 1&
    TheParent& = GetParent(TheWindow&)
    If TheParent& = 0& Then
        ParentCount& = 0&
        Exit Function
    End If
    Do
        TheParent& = GetParent(TheParent&)
        If TheParent& = 0& Then Exit Do
        ParentCounter& = 1& + ParentCounter&
    Loop
    ParentCount& = ParentCounter&
End Function

Public Sub PlayWAV(ByVal TheFile As String)
    'this sub plays a .WAV file
    'Example of use:
    '   Call PlayWAV("C:\MyWav.WAV")
    Dim DummyString As String
    DummyString$ = Dir(TheFile$)
    If Len(DummyString$) = 0 Then Exit Sub
    Call sndPlaySound(DummyString$, snd_flags)
End Sub

Public Function ReplaceString(ByVal TheString As String, ByVal Find As String, ByVal ReplaceWith As String) As String
    'this function replaces a string in a string with a string, hehe
    'VB6 has a function called "Replace" built in that does the same thing
    'if your not using VB6, this is for you.
    'Example of use:
    '   MsgBox ReplaceString("http://1ome.to/1ruelair", "1", "c")
    Dim TheINSTR As Integer, TempString As String, FindLen As Integer
    Dim LeftVal As String, RightVal As String
    TheINSTR% = InStr(1, TheString$, Find$, 1)
    If TheINSTR% = 0 Then
        ReplaceString$ = TheString$
        Exit Function
    End If
    TempString$ = TheString$
    FindLen% = Len(Find$)
    Do
        LeftVal$ = Mid(TempString$, 1, TheINSTR% - 1)
        RightVal$ = Mid(TempString$, TheINSTR% + FindLen%)
        TempString$ = LeftVal$ & ReplaceWith$ & RightVal$
        TheINSTR% = InStr(1, TempString$, Find$, 1)
    Loop Until TheINSTR% = 0
    ReplaceString$ = TempString$
End Function
Public Sub AddFavoritePlace(ByVal TheLocation As String, ByVal TheDescription As String)
    'this sub adds a favorite place to AOL's favorite places window
    'Example of use:
    '   Call AddFavoritePlace("http://come.to/cruelair", "click me")
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Dim FPWin As Long, NewIco As Long, IsLoaded As Byte
    IsLoaded = 1
    FPWin& = FindFavoritePlaces&
    If FPWin& = 0& Then
        Call LoadFavoritePlaces
        IsLoaded = 0
        Do
            DoEvents
            FPWin& = FindFavoritePlaces&
        Loop Until FPWin& <> 0&
    End If
    NewIco& = FindWindowEx(FPWin&, 0&, "_AOL_Icon", vbNullString)
    NewIco& = FindWindowEx(FPWin&, NewIco&, "_AOL_Icon", vbNullString)
    Call PostMessage(NewIco&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(NewIco&, WM_LBUTTONUP, 0&, 0&)
    Dim AddWin As Long, LocEdit As Long, DesEdit As Long
    Dim OkIco As Long, AOLMsg As Long, MsgButton As Long
    Do
        DoEvents
        AddWin& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Add New Folder/Favorite Place")
        DesEdit& = FindWindowEx(AddWin&, 0&, "_AOL_Edit", vbNullString)
        DesEdit& = FindWindowEx(AddWin&, DesEdit&, "_AOL_Edit", vbNullString)
        LocEdit& = FindWindowEx(AddWin&, DesEdit&, "_AOL_Edit", vbNullString)
        OkIco& = FindWindowEx(AddWin&, 0&, "_AOL_Icon", vbNullString)
        OkIco& = FindWindowEx(AddWin&, OkIco&, "_AOL_Icon", vbNullString)
    Loop Until (IsWindowVisible(DesEdit&) = 1) And (IsWindowVisible(LocEdit&) = 1) And (IsWindowVisible(OkIco&) = 1)
    Call SendMessageByString(DesEdit&, WM_SETTEXT, 0&, TheDescription$)
    Call SendMessageByString(LocEdit&, WM_SETTEXT, 0&, TheLocation$)
    DoEvents
    Call PostMessage(OkIco&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OkIco&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        AOLMsg& = FindWindow("#32770", "America Online")
        MsgButton& = FindWindowEx(AOLMsg&, 0&, "Button", vbNullString)
        AddWin& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Add New Folder/Favorite Place")
    Loop Until (AddWin& = 0&) Or (IsWindowVisible(MsgButton&) = 1)
    If MsgButton& <> 0& Then
        Call PostMessage(MsgButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(MsgButton&, WM_KEYUP, VK_SPACE, 0&)
        DoEvents
        Call PostMessage(AddWin&, WM_CLOSE, 0&, 0&)
    End If
    If IsLoaded = 0 Then
        Call PostMessage(FPWin&, WM_CLOSE, 0&, 0&)
        DoEvents
    End If
End Sub

Public Sub AddFontsToCombo(ByVal TheCombo As ComboBox)
    'this sub adds your screen fonts to a combobox
    'Example of use:
    '   Call AddFontsToList(Combo1)
    Dim LoopThrough As Long
    For LoopThrough& = 0& To Screen.FontCount - 1
        TheCombo.AddItem Screen.Fonts(LoopThrough&)
    Next LoopThrough&
End Sub

Public Sub AddFontsToList(ByVal TheList As ListBox)
    'this sub adds your screen fonts to a listbox
    'Example of use:
    '   Call AddFontsToList(List1)
    Dim LoopThrough As Long
    For LoopThrough& = 0& To Screen.FontCount - 1
        TheList.AddItem Screen.Fonts(LoopThrough&)
    Next LoopThrough&
End Sub
Public Sub AddRoomToCombobox(ByVal TheComboBox As ComboBox, ByVal AddUser As Boolean)
    'this sub will add the names in AOL's chat room to a combo box
    'Example of use:
    '   Call AddRoomToComboBox(Combo1)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long, TheINSTR As Integer
    Dim RoomList As Long, mThread As Long, TheUser As String
    Room& = FindChatRoom&
    If Room& = 0& Then Exit Sub
    TheUser$ = GetUser$
    RoomList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call GetWindowThreadProcessId(RoomList&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(RoomList&, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, Chr(0))
            itmHold& = SendMessage(RoomList&, LB_GETITEMDATA, ByVal CLng(Index&), 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes&)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, Chr(0))
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            TheINSTR% = InStr(1, ScreenName$, Chr(0), 1)
            ScreenName$ = Left$(ScreenName$, TheINSTR% - 1)
            If ScreenName$ <> TheUser$ Or AddUser = True Then
                TheComboBox.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mThread&)
    End If
    If TheComboBox.ListCount > 0 Then
        TheComboBox.Text = TheComboBox.List(0)
    End If
End Sub
Public Sub AddRoomToListbox(ByVal TheListbox As ListBox, ByVal AddUser As Boolean)
    'this sub will add the names in AOL's chat room to a list box
    'Example of use:
    '   Call AddRoomToListBox(List1)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long, TheINSTR As Integer
    Dim RoomList As Long, mThread As Long, TheUser As String
    Room& = FindChatRoom&
    If Room& = 0& Then Exit Sub
    TheUser$ = GetUser$
    RoomList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    Call GetWindowThreadProcessId(RoomList&, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(RoomList&, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, Chr(0))
            itmHold& = SendMessage(RoomList&, LB_GETITEMDATA, ByVal CLng(Index&), 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes&)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, Chr(0))
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            TheINSTR% = InStr(1, ScreenName$, Chr(0), 1)
            ScreenName$ = Left$(ScreenName$, TheINSTR% - 1)
            If ScreenName$ <> TheUser$ Or AddUser = True Then
                TheListbox.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mThread&)
    End If
End Sub


Public Function ChatName() As String
    'this function returns the name of an aol chatroom if one is open
    'Example of use:
    '   MsgBox ChatName$
    ChatName$ = GetCaption(FindChatRoom&)
End Function
Public Function ChatNameAIM() As String
    'this function gets the name of AIM's chat room if its open
    'Example of use:
    '   MsgBox ChatNameAIM$
    Dim TheRoom As Long
    TheRoom& = FindAIMChat&
    If TheRoom& = 0& Then
        ChatNameAIM$ = ""
        Exit Function
    End If
    Dim TheCaption As String
    TheCaption$ = GetCaption(TheRoom&)
    ChatNameAIM$ = Mid(TheCaption$, 2 + InStr(1, TheCaption$, ":", 1))
End Function

Public Sub CloseCDRom()
    'this sub will close the CD Rom drive
    'Example of use:
    '   Call CloseCDRom
    Call mciSendString("set CDAudio door closed", vbNullString, 0&, 0&)
End Sub
Public Function CountFlashMail() As Long
    'this function will count the number of mails in AOL's flashmail box
    'Example of use:
    '   MsgBox CountFlashMail&
    If FindWindow("AOL Frame25", vbNullString) = 0& Then
        CountFlashMail& = 0&
        Exit Function
    End If
    Dim AOLTree As Long, LoadBox As Byte
    LoadBox = 0
    AOLTree& = FindFlashMailBox&
    If AOLTree& = 0& Then
        Call LoadFlashmail
        Do
            DoEvents
            AOLTree& = FindFlashMailBox&
        Loop Until AOLTree& <> 0&
        Call Sleep(2000)
        LoadBox = 1
        DoEvents
    End If
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_Tree", vbNullString)
    CountFlashMail& = SendMessage(AOLTree&, LB_GETCOUNT, 0&, 0&)
    If LoadBox = 1 Then
        AOLTree& = FindFlashMailBox&
        Call PostMessage(AOLTree&, WM_CLOSE, 0&, 0&)
        DoEvents
    End If
End Function

Public Function CountNewMail() As Long
    'this function will count the number of mails in AOL's mailbox
    'Example of use:
    '   MsgBox CountNewMail&
    If FindWindow("AOL Frame25", vbNullString) = 0& Then
        CountNewMail& = 0&
        Exit Function
    End If
    Dim AOLTree As Long, LoadBox As Byte
    LoadBox = 0
    AOLTree& = FindMailBox&
    If AOLTree& = 0& Then
        Call LoadMailBox
        Do
            DoEvents
            AOLTree& = FindMailBox&
        Loop Until AOLTree& <> 0&
        Call Sleep(2000)
        LoadBox = 1
        DoEvents
    End If
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_TabControl", vbNullString)
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_TabPage", vbNullString)
    AOLTree& = FindWindowEx(AOLTree&, 0&, "_AOL_Tree", vbNullString)
    CountNewMail& = SendMessage(AOLTree&, LB_GETCOUNT, 0&, 0&)
    If LoadBox = 1 Then
        AOLTree& = FindMailBox&
        Call PostMessage(AOLTree&, WM_CLOSE, 0&, 0&)
        DoEvents
    End If
End Function


Public Function CountUniqueColors(ByVal ThePicturebox As PictureBox, ByVal ColorList As ListBox) As Long
    'this function counts the number of unique colors in a picturebox
    'this function requires a listbox to keep track of colors
    'Example of use:
    '   MsgBox CountUniqueColors(Picture1, List1)
    Dim TheDC As Long, SearchList As Long, LastColor As Long, IsVis As Byte
    Dim LoopThroughV As Long, LoopThroughH As Long, TheColor As Long
    If ThePicturebox.Visible = False Then
        IsVis = 0
    Else
        IsVis = 1
    End If
    TheDC& = ThePicturebox.hdc
    ThePicturebox.Visible = True
    ThePicturebox.AutoRedraw = True
    ThePicturebox.ScaleMode = 3
    ColorList.Clear
    DoEvents
    LastColor& = -2
    For LoopThroughV& = 1 To ThePicturebox.ScaleHeight
        For LoopThroughH& = 1 To ThePicturebox.ScaleWidth
            TheColor& = GetPixel(TheDC&, LoopThroughH&, LoopThroughV&)
            If TheColor& <> LastColor& Then
                SearchList& = SendMessageByString(ColorList.hWnd, LB_FINDSTRINGEXACT, 0&, CStr(TheColor&))
                If SearchList& = -1 Then
                    ColorList.AddItem CStr(TheColor&)
                    DoEvents
                End If
                LastColor& = TheColor&
            End If
        Next LoopThroughH&
    Next LoopThroughV&
    If IsVis = 0 Then ThePicturebox.Visible = False
    CountUniqueColors& = ColorList.ListCount
End Function

Public Sub ExitAOL()
    'this sub will exit AOL
    'Example of use:
    '   Call ExitAOL
    Call RunMenuByString(FindWindow("AOL Frame25", vbNullString), "Exit")
End Sub
Public Function FindAIMChat() As Long
    'this function finds AIM's chat room
    'Example of use:
    '   MsgBox FindAIMChat&
    FindAIMChat& = FindWindow("AIM_ChatWnd", vbNullString)
End Function

Public Function FindAOLChild(ByVal NumberOfIcons As Byte, ByVal NumberOfRICHCNTLs As Byte, ByVal NumberOfEdits As Byte, ByVal NumberOfListboxes As Byte, ByVal NumberOfComboboxes As Byte, ByVal NumberOfStatics As Byte) As Long
    'this function finds an AOL Child depending on the number of
    '_AOL_Icon, RICHCNTL, _AOL_Edit, _AOL_Listbox, _AOL_Combobox, and _AOL_Static it has
    'Example of use:
    '   Dim TheChild As Long
    '   TheChild& = FindAOLChild(2, 1, 0, 1, 0, 5)
    '   MsgBox GetCaption(TheChild&)
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindAOLChild& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, Rich As Long
    Dim AOLEdit As Long, AOLList As Long, AOLCombo As Long
    Dim AOLStatic As Long, LoopThrough As Byte
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        If NumberOfIcons > 0 Then
            AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
            For LoopThrough = 2 To NumberOfIcons
                AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
            Next LoopThrough
        End If
        If NumberOfRICHCNTLs > 0 Then
            Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
            For LoopThrough = 2 To NumberOfRICHCNTLs
                Rich& = FindWindowEx(AOLChild&, Rich&, "RICHCNTL", vbNullString)
            Next LoopThrough
        End If
        If NumberOfEdits > 0 Then
            AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
            For LoopThrough = 2 To NumberOfEdits
                AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
            Next LoopThrough
        End If
        If NumberOfListboxes > 0 Then
            AOLList& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
            For LoopThrough = 2 To NumberOfListboxes
                AOLList& = FindWindowEx(AOLChild&, AOLList&, "_AOL_Listbox", vbNullString)
            Next LoopThrough
        End If
        If NumberOfComboboxes > 0 Then
            AOLCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
            For LoopThrough = 2 To NumberOfComboboxes
                AOLCombo& = FindWindowEx(AOLChild&, AOLCombo&, "_AOL_Combobox", vbNullString)
            Next LoopThrough
        End If
        If NumberOfStatics > 0 Then
            AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
            For LoopThrough = 2 To NumberOfStatics
                AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
            Next LoopThrough
        End If
        If (NumberOfIcons > 0 And AOLIcon& <> 0&) Or (NumberOfIcons = 0) Then
            If (NumberOfRICHCNTLs > 0 And Rich& <> 0&) Or (NumberOfRICHCNTLs = 0) Then
                If (NumberOfEdits > 0 And AOLEdit& <> 0&) Or (NumberOfEdits = 0) Then
                    If (NumberOfListboxes > 0 And AOLList& <> 0&) Or (NumberOfListboxes = 0) Then
                        If (NumberOfComboboxes > 0 And AOLCombo& <> 0&) Or (NumberOfComboboxes = 0) Then
                            If (NumberOfStatics > 0 And AOLStatic& <> 0&) Or (NumberOfStatics = 0) Then
                                FindAOLChild& = AOLChild&
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindAOLChild& = 0&
            
End Function

Public Function FindChildByClass(ByVal TheParent As Long, ByVal TheClass As String, ByVal SiblingsBefore As Long) As Long
    'this function will find the child of a parent by its class
    'Example of use:
    '   Dim TheWin As Long
    '   TheWin& = FindWindow("AOL Frame25", vbNullString)
    '   TheWin& = FindWindowEx(TheWin&, 0&, "AOL Toolbar", vbNullString)
    '   TheWin& = FindWindowEx(TheWin&, 0&, "_AOL_Toolbar", vbNullString)
    '   MsgBox GetText(FindChildByClass(TheWin&, "_AOL_Icon", 0&))
    '   MsgBox GetText(FindChildByClass(TheWin&, "_AOL_Icon", 2&))
    Dim DummyWin As Long
    DummyWin& = FindWindowEx(TheParent&, 0&, TheClass$, vbNullString)
    If SiblingsBefore& = 0& Then
        FindChildByClass& = DummyWin&
    Else
        Dim LoopThrough As Long
        For LoopThrough& = 1 To SiblingsBefore&
            DummyWin& = FindWindowEx(TheParent&, DummyWin&, TheClass$, vbNullString)
            If DummyWin& = 0& Then Exit For
        Next LoopThrough&
        FindChildByClass& = DummyWin&
    End If
End Function

Public Function FindChildByText(ByVal TheParent As Long, ByVal TheText As String) As Long
    'this function finds a windows child by its text
    'Example of use:
    '   Dim MDIClient As Long
    '   MDIClient& = FindWindow("AOL Frame25", vbNullString)
    '   MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    '   MsgBox FindChildByText(MDIClient&, "Buddy List Window")
    Dim DummyWin As Long
    DummyWin& = GetWindow(TheParent&, GW_CHILD)
    DummyWin& = GetWindow(DummyWin&, GW_HWNDFIRST)
    If GetText(DummyWin&) = TheText$ Then
        FindChildByText& = DummyWin&
        Exit Function
    End If
    Do
        DummyWin& = GetWindow(DummyWin&, GW_HWNDNEXT)
        If GetText(DummyWin&) = TheText$ Then
            FindChildByText& = DummyWin&
            Exit Function
        End If
    Loop Until DummyWin& = 0&
    FindChildByText& = 0&
End Function

Public Function FindDownloadManager() As Long
    'this function finds AOL's download manager window
    'Example of use:
    '   MsgBox FindDownloadManager&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindDownloadManager& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, AOLTree As Long
    Dim AOLIcon2 As Long, Rich As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        If (IsWindowVisible(AOLIcon&) = 1&) And (AOLIcon2& = 0&) And (IsWindowVisible(AOLTree&) = 1) And (Rich& = 0&) Then
            FindDownloadManager& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindDownloadManager& = 0&
End Function
Public Function FindFavoritePlaces() As Long
    'this will find AOL's favorite places window
    'Example of use:
    '   MsgBox FindFavoritePlaces&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindFavoritePlaces& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, AOLTree As Long
    Dim AOLIcon2 As Long, Rich As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        If (IsWindowVisible(AOLIcon&) = 1) And (AOLIcon2& = 0) And (IsWindowVisible(AOLTree&) = 1) And (Rich& = 0&) Then
            FindFavoritePlaces& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindFavoritePlaces& = 0&
End Function
Public Function FindForwardWindow() As Long
    'this function finds the window that forwards a mail on AOL
    'Example of use:
    '   MsgBox FindForwardWindow&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindForwardWindow& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLCombo As Long, AOLFCombo As Long
    Dim Rich As Long, AOLEdit As Long, AOLIcon As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
        AOLFCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_FontCombo", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        If (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLCombo&) = 1) And (IsWindowVisible(AOLFCombo&) = 1) And (IsWindowVisible(AOLEdit&) = 1) And (IsWindowVisible(AOLIcon&) = 1) Then
            FindForwardWindow& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindForwardWindow& = 0&
End Function
Public Function FindMailBox() As Long
    'this function finds AOL's regular mailbox
    'Example of use:
    '   MsgBox FindMailBox&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindMailBox& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, AOLImage As Long
    Dim AOLTab1 As Long, AOLTab2 As Long, AOLTab3 As Long, AOLTabControl As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
        AOLTabControl& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
        AOLTab1& = FindWindowEx(AOLTabControl&, 0&, "_AOL_TabPage", vbNullString)
        AOLTab2& = FindWindowEx(AOLTabControl&, AOLTab1&, "_AOL_TabPage", vbNullString)
        AOLTab3& = FindWindowEx(AOLTabControl&, AOLTab2&, "_AOL_TabPage", vbNullString)
        If (AOLIcon& <> 0&) And (AOLImage& <> 0&) And (IsWindowVisible(AOLTab1&) = 1 Or IsWindowVisible(AOLTab2&) = 1 Or IsWindowVisible(AOLTab3&) = 1) Then
            FindMailBox& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindMailBox& = 0&
End Function
Public Function FindMIRC() As Long
    'this function finds mIRC's main window
    'Example of use:
    '   MsgBox FindMIRC&
    FindMIRC& = FindWindow("mIRC32", vbNullString)
End Function

Public Function FindMIRCStatus() As Long
    'this function finds mIRC's status window
    'Example of use:
    '   MsgBox FindMIRCStatus&
    Dim mIRC As Long
    mIRC& = FindMIRC&
    mIRC& = FindWindowEx(mIRC&, 0&, "MDIClient", vbNullString)
    FindMIRCStatus& = FindWindowEx(mIRC&, 0&, "status", vbNullString)
End Function


Public Function FindWelcomeScreen() As Long
    'this function finds AOL's welcome screen
    'Example of use:
    '   MsgBox GetCaption(FindWelcomeScreen&)
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindWelcomeScreen& = 0&
        Exit Function
    End If
    Dim AOLChild As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        If InStr(1, GetCaption(AOLChild&), "Welcome, ", 1) = 1 Then
            FindWelcomeScreen& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindWelcomeScreen& = 0&
End Function
Public Sub FormNotOnTop(ByVal TheForm As Form)
    'this sub will undo the FormOnTop sub
    'Example of use:
    '   Call FormNotOnTop(Form1)
    Call SetWindowPos(TheForm.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub FormOnTop(ByVal TheForm As Form)
    'this sub will keep your form on top of all other windows
    'Example of use:
    '   Call FormOnTop(Form1)
    Call SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Function GetChatText() As String
    'this function gets the text from AOL's chat room
    'Example of use:
    '   Text1.Text = GetChatText$
    Dim Rich As Long
    Rich& = FindWindowEx(FindChatRoom&, 0&, "RICHCNTL", vbNullString)
    GetChatText$ = GetText(Rich&)
End Function

Public Function GetMailSender() As String
    'this function gets the name of the person who sent an email
    'Example of use:
    '   MsgBox GetMailSender$
    Dim TheMail As Long, TheText As String
    TheMail& = FindReadWindow&
    If TheMail& = 0& Then
        GetMailSender$ = ""
        Exit Function
    End If
    TheMail& = FindWindowEx(TheMail&, 0&, "RICHCNTL", vbNullString)
    TheText$ = GetText(TheMail&)
    TheText$ = LineFromString(TheText$, 3)
    GetMailSender$ = Mid(TheText$, 7)
End Function

Public Function GetMemberInfoAIM(ByVal TheMember As String) As String
    'this function will get a members info on AIM
    'Example of use:
    '   MsgBox GetMemberInfoAIM("i am crue")
    Dim AIMWin As Long
    AIMWin& = FindAIM&
    If AIMWin& = 0& Then
        GetMemberInfoAIM$ = ""
        Exit Function
    End If
    Call RunMenuByString(AIMWin&, "Get Member Inf&o")
    Dim LocateWin As Long, NameEdit As Long, OKButton As Long
    Do
        DoEvents
        LocateWin& = FindWindow("_Oscar_Locate", "Buddy Info: ")
        NameEdit& = FindWindowEx(LocateWin&, 0&, "_Oscar_PersistantCombo", vbNullString)
        NameEdit& = FindWindowEx(NameEdit&, 0&, "Edit", vbNullString)
        OKButton& = FindWindowEx(LocateWin&, 0&, "Button", vbNullString)
    Loop Until (IsWindowVisible(LocateWin&) = 1) And (IsWindowVisible(NameEdit&) = 1) And (IsWindowVisible(OKButton&) = 1)
    Call SendMessageByString(NameEdit&, WM_SETTEXT, 0&, TheMember$)
    DoEvents
    Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
    Dim TheInfoWin As Long, TheText As String
    Do
        DoEvents
        TheInfoWin& = FindWindowEx(LocateWin&, 0&, "WndAte32Class", vbNullString)
        TheText$ = GetText(TheInfoWin&)
    Loop Until (InStr(1, TheText$, "Please wait...", 1) <> 31)
    OKButton& = FindWindowEx(LocateWin&, OKButton&, "Button", vbNullString)
    GetMemberInfoAIM$ = GetText(TheInfoWin&)
    Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
End Function


Public Sub GetProfile(ByVal Who As String)
    'this sub gets an AOL members profile
    'Example of use:
    '   Call GetProfile("Steve Case")
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Dim AOLIcon As Long, CPos As POINTAPI, Mnu As Long
    AOLIcon& = GetAOLToolbarIcon(10)
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        Mnu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(Mnu&) = 1
    DoEvents
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CPos.X, CPos.Y)
    Dim AOLChild As Long, AOLEdit As Long
    Do
        DoEvents
        AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Get a Member's Profile")
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    Loop Until (IsWindowVisible(AOLEdit&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Who$)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Sub

Public Function GetUserAIM() As String
    'this function gets the users AIM name
    'Example of use:
    '   MsgBox GetUserAIM$
    Dim TheWin As Long, TheCaption As String
    TheWin& = FindAIM&
    If TheWin& = 0& Then
        GetUserAIM$ = ""
        Exit Function
    End If
    TheCaption$ = GetCaption(TheWin&)
    GetUserAIM$ = Left(TheCaption$, InStr(1, TheCaption$, "'s", 1) - 1)
End Function

Public Function LineFromString(ByVal TheString As String, ByVal Line As Long) As String
    'this function returns a specific line from a multi-line string
    'Example of use:
    '   Dim RICH As Long
    '   RICH& = FindWindowEx(FindChatRoom&, 0&, "RICHCNTL", vbNullString)
    '   MsgBox LineFromString(GetText(RICH&), 2)
    Dim theline As String
    Dim StartAt As Long, EndAt As Long, LoopThrough As Long
    If Line& = 1 Then
        theline$ = Left(TheString$, InStr(TheString$, Chr(13)) - 1)
        theline$ = Replace(theline$, Chr(10), "")
        theline$ = Replace(theline$, Chr(13), "")
        LineFromString$ = theline$
        Exit Function
    Else
        StartAt& = InStr(TheString$, Chr(13))
        For LoopThrough& = 1 To Line& - 1
            EndAt& = StartAt&
            StartAt& = InStr(1 + StartAt&, TheString$, Chr(13))
        Next LoopThrough
        If StartAt = 0 Then
            StartAt = Len(TheString$)
        End If
        theline$ = Mid(TheString$, EndAt&, 1 + (StartAt& - EndAt&))
        theline$ = Replace(theline$, Chr(13), "")
        theline$ = Replace(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function

Public Function CapitalizeFirstLetter(ByVal TheString As String) As String
    'this function will capitalize the first letter in every word
    'this was written in VB6, i dont know if it will work in lower
    'versions or not
    'Example of use:
    '   MsgBox CapitalizeFirstLetter("hey, i'm just testing this.  visit http://come.to/cruelair")
    CapitalizeFirstLetter$ = StrConv(TheString$, vbProperCase)
End Function
Public Sub CloseAllIMs()
    'this sub will close all of AOL's instant messages
    'Example of use:
    '   Call CloseAllIMs
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    Dim AOLChild As Long, TheINSTR As Long, TheCaption As String
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
FindTheIMs:
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do While AOLChild& <> 0&
        DoEvents
        TheCaption$ = GetCaption(AOLChild&)
        TheINSTR& = InStr(1, TheCaption$, "Instant Message", 1)
        If (TheINSTR& = 1) Or (TheINSTR& = 2) Or (TheINSTR& = 3) Then
            Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
            DoEvents
            GoTo FindTheIMs
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop
End Sub
Public Function FindReadWindow() As Long
    'this function finds a open mail the user is reading, in AOL
    'Example of use:
    '   MsgBox FindReadWindow&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindReadWindow& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, Rich As Long, Rich2 As Long
    Dim AOLCombo As Long, AOLFCombo As Long, AOLEdit As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        Rich2& = FindWindowEx(AOLChild&, Rich&, "RICHCNTL", vbNullString)
        AOLCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
        AOLFCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_FontCombo", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        If (IsWindowVisible(Rich&) = 1) And (Rich2& = 0&) And (AOLCombo& = 0&) And (AOLFCombo& = 0&) And (AOLEdit& = 0&) Then
            FindReadWindow& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindReadWindow& = 0&
End Function
Public Function FractionToDecimal(ByVal TheFraction As String) As String
    'this is just a function i wrote to help me with my homework
    'i thought somebody else might find it useful, so i put it here
    'Example of use:
    '   MsgBox FractionToDecimal("3/19")
    On Error Resume Next
    Dim Value1 As Long, Value2 As Long, TheINSTR As Long, TheLen As Long
    TheLen& = Len(TheFraction$)
    TheINSTR& = InStr(1, TheFraction$, "/", 1)
    If (TheINSTR& = 0&) Or (TheINSTR& = 1) Or (TheINSTR& = TheLen&) Then
        FractionToDecimal = 0#
        Exit Function
    End If
    Value1& = Left(TheFraction$, TheINSTR - 1)
    Value2& = Right(TheFraction$, TheLen& - TheINSTR&)
    FractionToDecimal$ = CStr(Value1& / Value2&)
End Function

Public Function LineCountWindow(ByVal TheWindow As Long) As Long
    'this will get the number of lines of text in a window
    'Example of use:
    '   MsgBox LineCountWindow(FindWindowEx(FindChatRoom&, 0&, "RICHCNTL", vbNullString))
    LineCountWindow& = SendMessage(TheWindow, EM_GETLINECOUNT, 0&, 0&)
End Function


Public Sub LoadBuddyList()
    'this sub will load AOL's buddy list
    'Example of use:
    '   Call LoadBuddyList
    Call Keyword("buddy view")
End Sub

Public Sub LoadDownloadManager()
    'this sub will load AOL's download manager window
    'Example of use:
    '   Call LoadDownloadManager
    Dim AOLIcon As Long
    AOLIcon& = GetAOLToolbarIcon(5)
    If AOLIcon& = 0& Then Exit Sub
    Dim Mnu As Long, CPos As POINTAPI
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        Mnu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(Mnu&) = 1
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
    DoEvents
    Call SetCursorPos(CPos.X, CPos.Y)
End Sub
Public Sub LoadFavoritePlaces()
    'this sub will load AOL's Favorite Places window
    'Example of use:
    '   Call LoadFavoritePlaces
    Dim AOLIcon As Long
    AOLIcon& = GetAOLToolbarIcon(7)
    If AOLIcon& = 0& Then Exit Sub
    Dim Mnu As Long, CPos As POINTAPI
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        Mnu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(Mnu&) = 1
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
    DoEvents
    Call SetCursorPos(CPos.X, CPos.Y)
End Sub

Public Sub LoadMailBox()
    'this sub loads AOL's regular mail box
    'Example of use:
    '   Call LoadMailBox
    Dim AOLIcon As Long
    AOLIcon& = GetAOLToolbarIcon(3)
    If AOLIcon& = 0& Then Exit Sub
    Dim TheMenu As Long, CPos As POINTAPI
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    DoEvents
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        TheMenu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(TheMenu&) = 1
    Call PostMessage(TheMenu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(TheMenu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(TheMenu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(TheMenu&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(TheMenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(TheMenu&, WM_KEYUP, VK_RETURN, 0&)
    DoEvents
    Call SetCursorPos(CPos.X, CPos.Y)
End Sub

Public Sub LoadOnlineClock()
    'this sub will load AOL's Online Clock
    'the online clock tells how long you've been online
    'Example of use:
    '   Call LoadOnlineClock
    Dim AOLIcon As Long
    AOLIcon& = GetAOLToolbarIcon(6)
    If AOLIcon& = 0& Then Exit Sub
    Dim Mnu As Long, CPos As POINTAPI
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        Mnu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(Mnu&) = 1
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
    DoEvents
    Call SetCursorPos(CPos.X, CPos.Y)
End Sub

Public Function MakeChr(ByVal TheString As String) As String
    'this function converts text into chr code
    'Example of use
    '   MsgBox MakeChr("http://come.to/cruelair")
    Dim LoopIt As Long, FinalString As String
    For LoopIt& = 1 To Len(TheString$)
        FinalString$ = FinalString$ & "Chr(" & Asc(Mid(TheString$, LoopIt&, 1)) & ") & "
    Next LoopIt&
    FinalString$ = Left(FinalString$, Len(FinalString$) - 2)
    MakeChr$ = FinalString$
End Function




Public Sub OpenAnyFile(ByVal FileName As String)
    'this sub will load any file
    'Example of use:
    ' Call OpenAnyFile("C:\Test.txt")
    Call ShellExecute(0&, "open", FileName$, "", "", SW_SHOW)
End Sub

Public Sub OpenCDRom()
    'this sub opens the CD Rom drive
    'Example of use
    '   Call OpenCDRom
    Call mciSendString("set CDAudio door open", vbNullString, 0&, 0&)
End Sub
Public Function PictureToHTML(ByVal ThePicture As PictureBox, ByVal TheCHR As String) As String
    'converts the picture in a picturebox to html
    'Example of use:
    '   Text1.Text = PictureToHTML(picture1, ";")
    Dim LoopThroughH As Long, LoopThroughV As Long, TheDC As Long
    Dim TempStr As String, TheMacro As String, ThePixel As Long
    Dim OldPixel As Long, OldStr As String
    OldPixel& = -1
    ThePicture.AutoRedraw = True
    ThePicture.ScaleMode = 3
    TheDC& = ThePicture.hdc
    TheMacro$ = ""
    For LoopThroughV& = 1 To ThePicture.ScaleHeight
        TempStr$ = ""
        For LoopThroughH& = 1 To ThePicture.ScaleWidth
            ThePixel& = GetPixel(TheDC&, LoopThroughH&, LoopThroughV&)
            If ThePixel& = OldPixel& Then
                TempStr$ = TempStr$ & TheCHR$
            Else
                OldStr$ = "</font><font color=" & Chr(34) & "#" & RGBtoHEX(ThePixel&) & Chr(34) & ">" & TheCHR$
                TempStr$ = TempStr$ & OldStr$
                OldPixel& = ThePixel&
            End If
        Next LoopThroughH&
        TheMacro$ = TheMacro$ & TempStr$ & "<br>" & vbCrLf
    Next LoopThroughV&
    PictureToHTML$ = TheMacro$
End Function

Public Function PictureToMacro(ByVal ThePicture As PictureBox) As String
    'converts the picture in a picturebox to a neat macro
    'Example of use:
    '   Text1.Text = PictureToMacro(picture1)
    Dim LoopThroughH As Long, LoopThroughV As Long, TheDC As Long
    Dim TempStr As String, TheMacro As String
    Dim ThePixelT As Long, ThePixelB As Long
    ThePicture.AutoRedraw = True
    ThePicture.ScaleMode = 3
    TheDC& = ThePicture.hdc
    TheMacro$ = ""
    For LoopThroughV& = 1 To ThePicture.ScaleHeight Step 2
        TempStr$ = ""
        For LoopThroughH& = 1 To ThePicture.ScaleWidth
            ThePixelT& = GetPixel(TheDC&, LoopThroughH&, LoopThroughV&)
            ThePixelB& = GetPixel(TheDC&, LoopThroughH&, 1 + LoopThroughV&)
            If ThePixelT& <> 16777215 And ThePixelB& <> 16777215 Then
                TempStr$ = TempStr$ & ";"
            ElseIf ThePixelT& = 16777215 And ThePixelB& <> 16777215 Then
                TempStr$ = TempStr$ & ","
            ElseIf ThePixelT& <> 16777215 And ThePixelB& = 16777215 Then
                TempStr$ = TempStr$ & "´"
            Else
                TempStr$ = TempStr$ & " "
            End If
        Next LoopThroughH&
        If Len(Trim(TempStr$)) > 0 Then
            TempStr$ = RTrim(TempStr$)
        End If
        TheMacro$ = TheMacro$ & TempStr$ & vbCrLf
    Next LoopThroughV&
    PictureToMacro$ = TheMacro$
End Function

Public Function PictureToMacroCustom(ByVal ThePicture As PictureBox, ByVal TheCHR As String) As String
    'converts the picture in a picturebox to a neat macro
    'using a custom chr
    'Example of use:
    '   Text1.Text = PictureToMacroCustom(picture1, ";")
    Dim LoopThroughH As Long, LoopThroughV As Long, TheDC As Long
    Dim TempStr As String, TheMacro As String, ThePixel As Long
    ThePicture.AutoRedraw = True
    ThePicture.ScaleMode = 3
    TheDC& = ThePicture.hdc
    TheMacro$ = ""
    For LoopThroughV& = 1 To ThePicture.ScaleHeight
        TempStr$ = ""
        For LoopThroughH& = 1 To ThePicture.ScaleWidth
            ThePixel& = GetPixel(TheDC&, LoopThroughH&, LoopThroughV&)
            Select Case ThePixel&
                Case 16777215
                    TempStr$ = TempStr$ & " "
                Case -1
                    TempStr$ = TempStr$ & " "
                Case Else
                    TempStr$ = TempStr$ & TheCHR$
            End Select
        Next LoopThroughH&
        If Len(Trim(TempStr$)) > 0 Then
            TempStr$ = RTrim(TempStr$)
        End If
        TheMacro$ = TheMacro$ & TempStr$ & vbCrLf
    Next LoopThroughV&
    PictureToMacroCustom$ = TheMacro$
End Function
Public Function PrivateRoom(ByVal TheRoom As String) As String
    'this function generates the keyword to get into a
    'private room on aol
    'Example of use:
    '   Call KeyWord(PrivateRoom("mp3"))
    PrivateRoom$ = "aol://2719:2-2-" & TheRoom$
End Function

Public Function RestrictedRoom(ByVal RoomName As String) As String
    'this function will generate the keyword to get you into a
    'restricted room on AOL
    'Example of use:
    '   Call KeyWord(RestrictedRoom$("warez"))
    Dim LoopThrough As Byte, TheKeyWord As String
    TheKeyWord$ = ""
    For LoopThrough = 1 To Len(RoomName$)
        TheKeyWord$ = TheKeyWord$ & Mid(RoomName$, LoopThrough, 1) & "%A0"
    Next LoopThrough
    RestrictedRoom$ = "aol://2719:2-2-" & TheKeyWord$
End Function
Public Function ReverseCase(ByVal TheString As String) As String
    'this function will reverse the casing in a string
    'Example of use:
    '   MsgBox ReverseCase("HeLlO")
    Dim LoopThrough As Long, TempString As String, DummyString As String
    Dim TheMid As String
    TempString$ = ""
    For LoopThrough& = 1 To Len(TheString$)
        TheMid$ = Mid(TheString$, LoopThrough&, 1)
        DummyString$ = TheMid$
        DummyString$ = CharLower(DummyString$)
        If DummyString$ = TheMid$ Then
            TempString$ = TempString$ & CharUpper(TheMid$)
        Else
            TempString$ = TempString$ & CharLower(TheMid$)
        End If
    Next LoopThrough
    ReverseCase$ = TempString$
End Function

Public Sub ReverseFlashmail()
    'this sub will reverse your flashmail
    'Example of use:
    '   Call ReverseFlashmail
    If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
    Dim MBox As Long
    MBox& = FindFlashMailBox&
    If MBox& = 0& Then
        Call LoadFlashmail
        Do
            DoEvents
            MBox& = FindFlashMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    Dim AOLTree As Long, TheCount As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_Tree", vbNullString)
    TheCount& = CountListItems(AOLTree&)
    If TheCount& < 2 Then Exit Sub
    Dim LoopThrough As Long
    For LoopThrough& = 0 To TheCount& - 2
        DoEvents
        DoEvents
        Call PostMessage(AOLTree&, WM_LBUTTONDOWN, 0&, 0&)
        DoEvents
        Call PostMessage(AOLTree&, LB_SETCURSEL, LoopThrough&, 0&)
        DoEvents
        Call PostMessage(AOLTree&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(AOLTree&, WM_KEYUP, VK_DOWN, 0&)
        DoEvents
        Call PostMessage(AOLTree&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
        DoEvents
    Next LoopThrough&
End Sub

Public Function ReverseString(ByVal TheString As String) As String
    'this function will reverse a string (backwards)
    'VB6 has a function built in called strReverse that does the same thing
    'Example of use:
    '   MsgBox ReverseString("Hello")
    Dim LoopThrough As Long, TempString As String
    TempString$ = ""
    For LoopThrough& = Len(TheString$) To 1 Step -1
        TempString$ = TempString$ & Mid(TheString$, LoopThrough&, 1)
    Next LoopThrough&
    ReverseString$ = TempString$
End Function
Public Function RGBtoHEX(ByVal RGB As Long) As String
    'this function converts a RGB value to a hex value
    'Example of use:
    '   MsgBox "<font color=" & Chr(34) & "#" & RGBtoHEX(vbYellow) & Chr(34) & ">"
    Dim Red As String, Green As String, Blue As String
    Dim Color1 As Long
    Color1& = RGB&
    Red$ = Hex(Color1& And 255)
    Green$ = Hex(Color1& \ 256 And 255)
    Blue$ = Hex(Color1& \ 65536 And 255)
    If Len(Red$) < 2 Then Red$ = "0" & Red$
    If Len(Green$) < 2 Then Green$ = "0" & Green$
    If Len(Blue$) < 2 Then Blue$ = "0" & Blue$
    RGBtoHEX = Red$ & Green$ & Blue$
End Function


Public Function RoomCount() As Long
    'this function counts the people in AOL's chatroom
    'Example of use
    '   MsgBox RoomCount&
    Dim TheList As Long
    TheList& = FindWindowEx(FindChatRoom&, 0&, "_AOL_Listbox", vbNullString)
    RoomCount& = SendMessage(TheList&, LB_GETCOUNT, 0&, 0&)
End Function

Public Sub ScrollMacro(ByVal TheMacro As String, ByVal PauseScroll As Boolean)
    'this sub will scroll a macro in an AOL chat room
    'Example of use: (assuming MyMacro$ = a macro)
    '   Call ScrollMacro(MyMacro$, True)
    Dim LoopThrough As Long, TheMacro2() As Variant, TheCount As Long
    TheCount& = LineCountString(TheMacro$)
    ReDim TheMacro2(TheCount&) As Variant
    For LoopThrough& = 1 To TheCount&
        TheMacro2(LoopThrough&) = CVar(LineFromString(TheMacro$, LoopThrough&))
    Next LoopThrough&
    If PauseScroll = True Then
        Call SendChatMulti(TheMacro2, True, 2)
    Else
        Call SendChatMulti(TheMacro2, False)
    End If
End Sub

Public Function SearchListBox(ByVal TheList As Long, ByVal SearchString As String) As Long
    'this function searches a list for a string (not case sensitive)
    'it will return -1 if the search was NOT found, if it was
    'found it will return the index of the item
    'remember, indexes in listboxes and comboboxes start at 0
    'Example of use
    '   Dim TheSearch As Long
    '   TheSearch& = SearchListBox(list1, "http://come.to/cruelair")
    '   If TheSearch& = -1 Then
    '       MsgBox "not found"
    '   Else
    '       MsgBox "the search was found in index #" & TheSearch&
    '   End If
    SearchListBox& = SendMessageByString(TheList&, LB_FINDSTRINGEXACT, 0&, SearchString$)
End Function



Public Sub SendAIM(ByVal ToWho As String, ByVal TheMessage As String)
    'this sub sends a IM using AIM
    'Example of use:
    '   Call SendAIM("i am crue", "hi")
    Dim AIMWin As Long
    AIMWin& = FindAIM&
    If AIMWin& = 0& Then Exit Sub
    Call RunMenuByString(AIMWin&, "Send &Instant Message")
    Dim IMWin As Long, WhoEdit As Long, MessageEdit As Long
    Dim TheCombo As Long, WndAte32Class As Long, SendButton As Long
    Do
        DoEvents
        IMWin& = FindWindow("AIM_IMessage", "Instant Message")
        TheCombo& = FindWindowEx(IMWin&, 0&, "_Oscar_PersistantCombo", vbNullString)
        WhoEdit& = FindWindowEx(TheCombo&, 0&, "Edit", vbNullString)
        WndAte32Class& = FindWindowEx(IMWin&, 0&, "WndAte32Class", vbNullString)
        MessageEdit& = FindWindowEx(IMWin&, WndAte32Class&, "WndAte32Class", vbNullString)
        SendButton& = FindWindowEx(IMWin&, 0&, "_Oscar_IconBtn", vbNullString)
    Loop Until (IsWindowVisible(WhoEdit&) = 1) And (IsWindowVisible(MessageEdit&) = 1) And (IsWindowVisible(SendButton&) = 1)
    Call SendMessageByString(WhoEdit&, WM_SETTEXT, 0&, ToWho$)
    Call SendMessageByString(MessageEdit&, WM_SETTEXT, 0&, TheMessage$)
    DoEvents
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub SendChatAIM(ByVal TheMessage As String)
    'this sub sends chat to AIM's chat room
    'Example of use:
    '   Call SendChatAIM("http://come.to/cruelair")
    Dim TheRoom As Long
    TheRoom& = FindAIMChat&
    If TheRoom& = 0& Then Exit Sub
    Dim WndAte As Long, Button As Long
    WndAte& = FindWindowEx(TheRoom&, 0&, "WndAte32Class", vbNullString)
    WndAte& = FindWindowEx(TheRoom&, WndAte&, "WndAte32Class", vbNullString)
    Button& = FindWindowEx(TheRoom&, 0&, "_Oscar_IconBtn", vbNullString)
    Button& = FindWindowEx(TheRoom&, Button&, "_Oscar_IconBtn", vbNullString)
    Button& = FindWindowEx(TheRoom&, Button&, "_Oscar_IconBtn", vbNullString)
    Button& = FindWindowEx(TheRoom&, Button&, "_Oscar_IconBtn", vbNullString)
    Call SendMessageByString(WndAte&, WM_SETTEXT, 0&, TheMessage$)
    Call PostMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(Button&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub SetChatPrefs(ByVal NotifyArrive As Boolean, ByVal NotifyLeave As Boolean, ByVal DoubleSpace As Boolean, ByVal Alphabetize As Boolean, ByVal Sounds As Boolean)
    'this sub will set your AOL chat prefrences
    'Example of use:
    '   Call SetChatPrefs(False, False, False, True, True)
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Dim Room As Long, AOLIcon As Long, AOLModal As Long
    Dim AOLCheckbox As Long
    Room& = FindChatRoom&
    If Room& <> 0& Then
        AOLIcon& = FindWindowEx(Room&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(Room&, AOLIcon&, "_AOL_Icon", vbNullString)
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
        Do
            DoEvents
            AOLModal& = FindWindow("_AOL_Modal", "Chat Preferences")
            AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
            AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
        Loop Until (IsWindowVisible(AOLCheckbox&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(NotifyArrive), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(NotifyLeave), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(DoubleSpace), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(Alphabetize), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(Sounds), 0&)
        DoEvents
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Else
        AOLIcon& = GetAOLToolbarIcon(6)
        Dim CPos As POINTAPI, Mnu As Long, AOLChild As Long
        Call GetCursorPos(CPos)
        Call SetCursorPos(Screen.Width, Screen.Height)
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
        Do
            Mnu& = FindWindow("#32768", vbNullString)
        Loop Until IsWindowVisible(Mnu&) = 1
        DoEvents
        Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
        Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
        Call PostMessage(Mnu&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(Mnu&, WM_KEYUP, VK_DOWN, 0&)
        Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
        DoEvents
        Call SetCursorPos(CPos.X, CPos.Y)
        Do
            DoEvents
            AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Preferences")
            AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
            AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
            AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
            AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        Loop Until (IsWindowVisible(AOLIcon&) = 1) And (GetText(AOLIcon&) = "Chat")
        DoEvents
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
        Do
            DoEvents
            AOLModal& = FindWindow("_AOL_Modal", "Chat Preferences")
            AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
            AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
        Loop Until (IsWindowVisible(AOLCheckbox&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(NotifyArrive), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(NotifyLeave), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(DoubleSpace), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(Alphabetize), 0&)
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(AOLCheckbox&, BM_SETCHECK, CLng(Sounds), 0&)
        DoEvents
        Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do
        DoEvents
        AOLIcon& = FindWindow("#32770", "America Online")
        AOLModal& = FindWindow("_AOL_Modal", "Chat Preferences")
    Loop Until (AOLModal& = 0&) Or (AOLIcon& <> 0&)
    If AOLIcon& <> 0& Then
        AOLIcon& = FindWindowEx(AOLIcon&, 0&, "Button", vbNullString)
        DoEvents
        Call PostMessage(AOLIcon&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
    End If
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub SetMailPrefs()
    'this sub will set AOL's mail prefrences
    'to the most commonly used settings
    'Example of use:
    '   Call SetMailPrefs
    Dim AOLIcon As Long
    AOLIcon& = GetAOLToolbarIcon(3)
    If AOLIcon& = 0& Then Exit Sub
    Dim Mnu As Long, CPos As POINTAPI
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        Mnu& = FindWindow("#32768", vbNullString)
    Loop Until IsWindowVisible(Mnu&) = 1
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(Mnu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(Mnu&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CPos.X, CPos.Y)
    Dim AOLModal As Long, AOLSpin As Long
    Dim AOLCheckbox1 As Long, AOLCheckbox2 As Long, AOLCheckbox3 As Long
    Do
        DoEvents
        AOLModal& = FindWindow("_AOL_Modal", "Mail Preferences")
        AOLCheckbox1& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
        AOLCheckbox2& = FindWindowEx(AOLModal&, AOLCheckbox1&, "_AOL_Checkbox", vbNullString)
        AOLCheckbox3& = FindWindowEx(AOLModal&, AOLCheckbox2&, "_AOL_Checkbox", vbNullString)
        AOLCheckbox3& = FindWindowEx(AOLModal&, AOLCheckbox3&, "_AOL_Checkbox", vbNullString)
        AOLCheckbox3& = FindWindowEx(AOLModal&, AOLCheckbox3&, "_AOL_Checkbox", vbNullString)
        AOLCheckbox3& = FindWindowEx(AOLModal&, AOLCheckbox3&, "_AOL_Checkbox", vbNullString)
        AOLSpin& = FindWindowEx(AOLModal&, 0&, "_AOL_Spin", vbNullString)
        AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until (IsWindowVisible(AOLCheckbox1&) = 1) And (IsWindowVisible(AOLCheckbox2&) = 1) And (IsWindowVisible(AOLCheckbox3&) = 1) And (AOLSpin& <> 0&) And (IsWindowVisible(AOLIcon&) = 1)
    Call PostMessage(AOLCheckbox1&, BM_SETCHECK, 0&, 0&)
    Call PostMessage(AOLCheckbox2&, BM_SETCHECK, 1&, 0&)
    Call PostMessage(AOLCheckbox3&, BM_SETCHECK, 0&, 0&)
    DoEvents
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub SetText(ByVal TheWindow As Long, ByVal TheText As String)
    'this sub will send text to any window you choose
    'Example of use:
    '   Call SetText(FindChatRoom&, "My Chat")
    Call SendMessageByString(TheWindow&, WM_SETTEXT, 0&, TheText$)
End Sub

Public Sub SignOffAIM()
    'this sub will sign a user off of AIM
    'Example of use:
    '   Call SignOffAIM
    Dim AIMWin As Long
    AIMWin& = FindAIM&
    If AIMWin& = 0& Then Exit Sub
    Call RunMenuByString(AIMWin&, "Sign O&ff")
End Sub

Public Sub SignOffAOL()
    'this sub will sign off of AOL
    'Example of use:
    '   Call SignOffAOL
    Call RunMenuByString(FindWindow("AOL Frame25", vbNullString), "&Sign Off")
End Sub

Public Sub SortArray(ByRef TheArray As Variant)
    'this sub will sort the contents of an Array alphabeticaly
    'Example of use:
    '   Dim MyArray(5) As Variant
    '   MyArray(1) = "c"
    '   MyArray(2) = "a"
    '   MyArray(3) = "e"
    '   MyArray(4) = "b"
    '   MyArray(5) = "d"
    '   Call SortArray(MyArray)
    '   MsgBox MyArray(1) & " - " & MyArray(2) & " - " & MyArray(3) & " - " & MyArray(4) & " - " & MyArray(5)
    Dim LoopThrough1 As Long, LoopThrough2 As Long
    For LoopThrough1& = UBound(TheArray) To LBound(TheArray) Step -1
        For LoopThrough2& = 1 + LBound(TheArray) To LoopThrough1&
            If TheArray(LoopThrough2& - 1) > TheArray(LoopThrough2&) Then
                Call SwitchValues(TheArray(LoopThrough2& - 1), TheArray(LoopThrough2&))
            End If
        Next LoopThrough2&
    Next LoopThrough1&
End Sub
Public Sub SortFlashmailByDate(ByVal AscendingOrder As Boolean, ByVal TheList As ListBox, ByVal TheList2 As ListBox)
    'this sub will sort your flashmail by date in Ascending or Descending order
    'this sub uses a function called InStrRev, wich is built into VB6
    'this sub uses a listbox to calculate the sort, an Array can also be used with the SortArray sub
    'if the mail is already sorted this will just leave the mail alone
    'i put a lot of DoEvents statements in the final few For...Next loops
    'they are just to slow the sorter down, so AOL dosent lagg behind on large mail lists
    '
    '***IMPORTANT: YOU MUST SET THE "SORTED" PROPERTY OF THE TheList TO TRUE!!!****
    '***IMPORTANT: YOU MUST SET THE "SORTED" PROPERTY OF TheList2 TO FALSE!!!****
    '
    'Example of use:
    '   Call SortFlashmailByDate(True, List1, List2)
    If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
    Dim MBox As Long
    MBox& = FindFlashMailBox&
    If MBox& = 0& Then
        Call LoadFlashmail
        Do
            DoEvents
            MBox& = FindFlashMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    TheList.Clear
    TheList2.Clear
    DoEvents
    Dim AOLTree As Long, TheCount As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_Tree", vbNullString)
    TheCount& = CountListItems(AOLTree&)
    If TheCount& < 2 Then Exit Sub
    TheCount& = TheCount& - 1
    Dim LoopThrough As Long
    Dim TheLen As Long, TheDate As String, TheINSTR As Integer
    For LoopThrough& = 0& To TheCount&
        TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheDate$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(AOLTree&, LB_GETTEXT, LoopThrough&, TheDate$)
        TheDate$ = Replace(TheDate$, Chr(0), "")
        TheList.AddItem TheDate$ & ":" & CStr(LoopThrough&)
        DoEvents
    Next LoopThrough&
    DoEvents
    If AscendingOrder = True Then
        For LoopThrough& = 0& To TheCount&
            TheINSTR% = InStrRev(TheList.List(LoopThrough&), ":")
            TheDate$ = Mid(TheList.List(LoopThrough&), 1 + TheINSTR%)
            TheList2.AddItem TheDate$, 0
        Next LoopThrough&
    Else
        For LoopThrough& = 0& To TheCount&
            TheINSTR% = InStrRev(TheList.List(LoopThrough&), ":")
            TheDate$ = Mid(TheList.List(LoopThrough&), 1 + TheINSTR%)
            TheList2.AddItem TheDate$
        Next LoopThrough&
    End If
    DoEvents
    Dim AlreadySorted As Boolean, CheckIt As Long
    AlreadySorted = True
    If AscendingOrder = True Then
        CheckIt& = TheCount&
        For LoopThrough& = 0& To TheCount&
            If CLng(TheList2.List(LoopThrough&)) <> CheckIt& Then
                AlreadySorted = False
                GoTo EndCheckSort
            End If
            CheckIt& = CheckIt& - 1
        Next LoopThrough&
    Else
        CheckIt& = 0&
        For LoopThrough& = TheCount& To 0& Step -1
            If CLng(TheList2.List(LoopThrough&)) <> CheckIt& Then
                AlreadySorted = False
                GoTo EndCheckSort
            End If
            CheckIt& = 1 + CheckIt&
        Next LoopThrough&
    End If
EndCheckSort:
    If AlreadySorted = True Then Exit Sub
    Call ShowWindow(MBox&, SW_HIDE)
    DoEvents
    Dim LoopThrough2 As Long, MoveFrom As Long, NextFrom As Long
    For LoopThrough& = 0& To TheCount&
        DoEvents
        DoEvents
        MoveFrom& = CLng(TheList2.List(LoopThrough&))
        If MoveFrom& <> 0& Then
            DoEvents
            DoEvents
            DoEvents
            Call PostMessage(AOLTree&, WM_LBUTTONDOWN, 0&, 0&)
            DoEvents
            Call PostMessage(AOLTree&, LB_SETCURSEL, MoveFrom& - 1, 0&)
            DoEvents
            Call PostMessage(AOLTree&, WM_KEYDOWN, VK_DOWN, 0&)
            Call PostMessage(AOLTree&, WM_KEYUP, VK_DOWN, 0&)
            DoEvents
            Call PostMessage(AOLTree&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        End If
        For LoopThrough2& = LoopThrough& To TheCount&
            DoEvents
            DoEvents
            NextFrom& = CLng(TheList2.List(LoopThrough2&))
            If MoveFrom& > NextFrom& Then
                DoEvents
                DoEvents
                DoEvents
                TheList2.List(LoopThrough2&) = CStr(1 + NextFrom&)
                DoEvents
                DoEvents
                DoEvents
            End If
            DoEvents
            DoEvents
        Next LoopThrough2&
        DoEvents
        DoEvents
        DoEvents
    Next LoopThrough&
    Call ShowWindow(MBox&, SW_SHOW)
    DoEvents
End Sub

Public Sub SortFlashmailBySender(ByVal AscendingOrder As Boolean, ByVal TheList As ListBox, ByVal TheList2 As ListBox)
    'this sub will sort your flashmail by sender in Ascending or Descending order
    'this sub uses a function called InStrRev, wich is built into VB6
    'this sub uses a listbox to calculate the sort, an Array can also be used with the SortArray sub
    'if the mail is already sorted this will just leave the mail alone
    'i put a lot of DoEvents statements in the final few For...Next loops
    'they are just to slow the sorter down, so AOL dosent lagg behind on large mail lists
    '
    '***IMPORTANT: YOU MUST SET THE "SORTED" PROPERTY OF THE TheList TO TRUE!!!****
    '***IMPORTANT: YOU MUST SET THE "SORTED" PROPERTY OF TheList2 TO FALSE!!!****
    '
    'Example of use:
    '   Call SortFlashmailBySender(True, List1, List2)
    If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
    Dim MBox As Long
    MBox& = FindFlashMailBox&
    If MBox& = 0& Then
        Call LoadFlashmail
        Do
            DoEvents
            MBox& = FindFlashMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    TheList.Clear
    TheList2.Clear
    DoEvents
    Dim AOLTree As Long, TheCount As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_Tree", vbNullString)
    TheCount& = CountListItems(AOLTree&)
    If TheCount& < 2 Then Exit Sub
    TheCount& = TheCount& - 1
    Dim LoopThrough As Long
    Dim TheLen As Long, TheSender As String, TheINSTR As Integer
    For LoopThrough& = 0& To TheCount&
        TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheSender$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(AOLTree&, LB_GETTEXT, LoopThrough&, TheSender$)
        TheINSTR% = InStr(1, TheSender$, Chr(9), 1)
        TheSender$ = Right(TheSender$, Len(TheSender$) - TheINSTR%)
        TheSender$ = Replace(TheSender$, Chr(0), "")
        TheList.AddItem TheSender$ & ":" & CStr(LoopThrough&)
        DoEvents
    Next LoopThrough&
    DoEvents
    If AscendingOrder = True Then
        For LoopThrough& = 0& To TheCount&
            TheINSTR% = InStrRev(TheList.List(LoopThrough&), ":")
            TheSender$ = Mid(TheList.List(LoopThrough&), 1 + TheINSTR%)
            TheList2.AddItem TheSender$, 0
        Next LoopThrough&
    Else
        For LoopThrough& = 0& To TheCount&
            TheINSTR% = InStrRev(TheList.List(LoopThrough&), ":")
            TheSender$ = Mid(TheList.List(LoopThrough&), 1 + TheINSTR%)
            TheList2.AddItem TheSender$
        Next LoopThrough&
    End If
    DoEvents
    Dim AlreadySorted As Boolean, CheckIt As Long
    AlreadySorted = True
    If AscendingOrder = True Then
        CheckIt& = TheCount&
        For LoopThrough& = 0& To TheCount&
            If CLng(TheList2.List(LoopThrough&)) <> CheckIt& Then
                AlreadySorted = False
                GoTo EndCheckSort
            End If
            CheckIt& = CheckIt& - 1
        Next LoopThrough&
    Else
        CheckIt& = 0&
        For LoopThrough& = TheCount& To 0& Step -1
            If CLng(TheList2.List(LoopThrough&)) <> CheckIt& Then
                AlreadySorted = False
                GoTo EndCheckSort
            End If
            CheckIt& = 1 + CheckIt&
        Next LoopThrough&
    End If
EndCheckSort:
    If AlreadySorted = True Then Exit Sub
    Call ShowWindow(MBox&, SW_HIDE)
    DoEvents
    Dim LoopThrough2 As Long, MoveFrom As Long, NextFrom As Long
    For LoopThrough& = 0& To TheCount&
        DoEvents
        DoEvents
        MoveFrom& = CLng(TheList2.List(LoopThrough&))
        If MoveFrom& <> 0& Then
            DoEvents
            DoEvents
            DoEvents
            Call PostMessage(AOLTree&, WM_LBUTTONDOWN, 0&, 0&)
            DoEvents
            Call PostMessage(AOLTree&, LB_SETCURSEL, MoveFrom& - 1, 0&)
            DoEvents
            Call PostMessage(AOLTree&, WM_KEYDOWN, VK_DOWN, 0&)
            Call PostMessage(AOLTree&, WM_KEYUP, VK_DOWN, 0&)
            DoEvents
            Call PostMessage(AOLTree&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        End If
        For LoopThrough2& = LoopThrough& To TheCount&
            DoEvents
            DoEvents
            NextFrom& = CLng(TheList2.List(LoopThrough2&))
            If MoveFrom& > NextFrom& Then
                DoEvents
                DoEvents
                DoEvents
                TheList2.List(LoopThrough2&) = CStr(1 + NextFrom&)
                DoEvents
                DoEvents
                DoEvents
            End If
            DoEvents
            DoEvents
        Next LoopThrough2&
        DoEvents
        DoEvents
        DoEvents
    Next LoopThrough&
    Call ShowWindow(MBox&, SW_SHOW)
    DoEvents
End Sub


Public Sub SortFlashmailBySubject(ByVal AscendingOrder As Boolean, ByVal TheList As ListBox, ByVal TheList2 As ListBox)
    'this sub will sort your flashmail by subject in Ascending or Descending order
    'this sub uses a function called InStrRev, wich is built into VB6
    'this sub uses a listbox to calculate the sort, an Array can also be used with the SortArray sub
    'if the mail is already sorted this will just leave the mail alone
    'i put a lot of DoEvents statements in the final few For...Next loops
    'they are just to slow the sorter down, so AOL dosent lagg behind on large mail lists
    '
    '***IMPORTANT: YOU MUST SET THE "SORTED" PROPERTY OF TheList TO TRUE!!!****
    '***IMPORTANT: YOU MUST SET THE "SORTED" PROPERTY OF TheList2 TO FALSE!!!****
    '
    'Example of use:
    '   Call SortFlashmailBySubject(True, List1, List2)
    If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
    Dim MBox As Long
    MBox& = FindFlashMailBox&
    If MBox& = 0& Then
        Call LoadFlashmail
        Do
            DoEvents
            MBox& = FindFlashMailBox&
            If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Loop Until MBox& <> 0&
        Call Sleep(2000)
        DoEvents
    End If
    TheList.Clear
    TheList2.Clear
    DoEvents
    Dim AOLTree As Long, TheCount As Long
    AOLTree& = FindWindowEx(MBox&, 0&, "_AOL_Tree", vbNullString)
    TheCount& = CountListItems(AOLTree&)
    If TheCount& < 2 Then Exit Sub
    TheCount& = TheCount& - 1
    Dim LoopThrough As Long
    Dim TheLen As Long, TheSubject As String, TheINSTR As Integer
    For LoopThrough& = 0& To TheCount&
        TheLen& = SendMessage(AOLTree&, LB_GETTEXTLEN, LoopThrough&, 0&)
        TheSubject$ = String(1 + TheLen&, 0&)
        Call SendMessageByString(AOLTree&, LB_GETTEXT, LoopThrough&, TheSubject$)
        TheINSTR% = InStr(1, TheSubject$, Chr(9), 1)
        TheINSTR% = InStr(1 + TheINSTR%, TheSubject$, Chr(9), 1)
        TheSubject$ = Right(TheSubject$, Len(TheSubject$) - TheINSTR%)
        TheSubject$ = Replace(TheSubject$, Chr(0), "")
        TheList.AddItem TheSubject$ & ":" & CStr(LoopThrough&)
        DoEvents
    Next LoopThrough&
    DoEvents
    If AscendingOrder = True Then
        For LoopThrough& = 0& To TheCount&
            TheINSTR% = InStrRev(TheList.List(LoopThrough&), ":")
            TheSubject$ = Mid(TheList.List(LoopThrough&), 1 + TheINSTR%)
            TheList2.AddItem TheSubject$, 0
        Next LoopThrough&
    Else
        For LoopThrough& = 0& To TheCount&
            TheINSTR% = InStrRev(TheList.List(LoopThrough&), ":")
            TheSubject$ = Mid(TheList.List(LoopThrough&), 1 + TheINSTR%)
            TheList2.AddItem TheSubject$
        Next LoopThrough&
    End If
    DoEvents
    Dim AlreadySorted As Boolean, CheckIt As Long
    AlreadySorted = True
    If AscendingOrder = True Then
        CheckIt& = TheCount&
        For LoopThrough& = 0& To TheCount&
            If CLng(TheList2.List(LoopThrough&)) <> CheckIt& Then
                AlreadySorted = False
                GoTo EndCheckSort
            End If
            CheckIt& = CheckIt& - 1
        Next LoopThrough&
    Else
        CheckIt& = 0&
        For LoopThrough& = TheCount& To 0& Step -1
            If CLng(TheList2.List(LoopThrough&)) <> CheckIt& Then
                AlreadySorted = False
                GoTo EndCheckSort
            End If
            CheckIt& = 1 + CheckIt&
        Next LoopThrough&
    End If
EndCheckSort:
    If AlreadySorted = True Then Exit Sub
    Call ShowWindow(MBox&, SW_HIDE)
    DoEvents
    Dim LoopThrough2 As Long, MoveFrom As Long, NextFrom As Long
    For LoopThrough& = 0& To TheCount&
        DoEvents
        DoEvents
        MoveFrom& = CLng(TheList2.List(LoopThrough&))
        If MoveFrom& <> 0& Then
            DoEvents
            DoEvents
            DoEvents
            Call PostMessage(AOLTree&, WM_LBUTTONDOWN, 0&, 0&)
            DoEvents
            Call PostMessage(AOLTree&, LB_SETCURSEL, MoveFrom& - 1, 0&)
            DoEvents
            Call PostMessage(AOLTree&, WM_KEYDOWN, VK_DOWN, 0&)
            Call PostMessage(AOLTree&, WM_KEYUP, VK_DOWN, 0&)
            DoEvents
            Call PostMessage(AOLTree&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        End If
        For LoopThrough2& = LoopThrough& To TheCount&
            DoEvents
            DoEvents
            NextFrom& = CLng(TheList2.List(LoopThrough2&))
            If MoveFrom& > NextFrom& Then
                DoEvents
                DoEvents
                DoEvents
                TheList2.List(LoopThrough2&) = CStr(1 + NextFrom&)
                DoEvents
                DoEvents
                DoEvents
            End If
            DoEvents
            DoEvents
        Next LoopThrough2&
        DoEvents
        DoEvents
        DoEvents
    Next LoopThrough&
    Call ShowWindow(MBox&, SW_SHOW)
    DoEvents
End Sub



Public Sub SwitchValues(ByRef Value1 As Variant, ByRef Value2 As Variant)
    'this sub will switch two values
    'Example of use:
    '   Dim Val1 As Variant, Val2 As Variant
    '   Val1 = 5
    '   Val2 = 10
    '   Call SwitchValues(Val1, Val2)
    '   MsgBox Val1 & vbCrLf & Val2
    Dim DummyValue As Variant
    DummyValue = Value1
    Value1 = Value2
    Value2 = DummyValue
End Sub








Public Sub VPSendIM(ByVal ToWho As String, ByVal Msg As String)
    'this sub will send an Instant Message using Virtual Places
    'Example of use:
    '   Call VPSendIM("crueizme", "hello")
    Dim VPWin As Long, IMWin As Long, ToEdit As Long, MsgEdit As Long
    Dim SendButton As Long, ClickSendTimer As Long
    VPWin& = FindWindow("VPFrame", vbNullString)
    If VPWin& = 0& Then Exit Sub
    Call RunMenuByString(VPWin&, "Send an &Instant Message")
    Do
        DoEvents
        IMWin& = FindWindow("#32770", "Send Instant Message")
        MsgEdit& = FindWindowEx(IMWin&, 0&, "Edit", vbNullString)
        ToEdit& = FindWindowEx(IMWin&, MsgEdit&, "Edit", vbNullString)
        SendButton& = FindWindowEx(IMWin&, 0&, "Button", vbNullString)
        SendButton& = FindWindowEx(IMWin&, SendButton&, "Button", vbNullString)
    Loop Until (IsWindowVisible(MsgEdit&) = 1) And (IsWindowVisible(ToEdit&) = 1) And (IsWindowVisible(SendButton&) = 1)
    DoEvents
    Call SendMessageByString(ToEdit&, WM_SETTEXT, 0&, ToWho$)
    Call SendMessageByString(MsgEdit&, WM_SETTEXT, 0&, Msg$)
    DoEvents
ClickSend:
    ClickSendTimer& = Timer
    Call SendMessage(SendButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(SendButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Do
        DoEvents
        IMWin& = FindWindow("#32770", "Send Instant Message")
        If Timer - ClickSendTimer& > 2 Then GoTo ClickSend
    Loop Until IMWin& = 0&
End Sub
Public Sub KillWait()
    'this sub kills the hour glass on AOL
    'Example of use:
    '   Call KillWait
    Dim aoModal As Long, Ico As Long, AOL As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    If AOL& = 0& Then Exit Sub
    aoModal& = FindWindow("_AOL_Modal", vbNullString)
    If aoModal& <> 0& Then Call KillAllModals
    DoEvents
    Call RunMenuByString(AOL&, "&About America Online")
    Do
        DoEvents
        aoModal& = FindWindow("_AOL_Modal", vbNullString)
        Ico& = FindWindowEx(aoModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until IsWindowVisible(Ico&) = 1
    DoEvents
    Call SendMessage(Ico&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Ico&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Call PostMessage(aoModal&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub ClearChat()
    'this sub will clear AOL's chat room
    'Example of use:
    '   Call ClearChat
    Dim Rich As Long
    Rich& = FindWindowEx(FindChatRoom&, 0&, "RICHCNTL", vbNullString)
    If Rich& = 0& Then Exit Sub
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "")
End Sub

Public Sub DeleteAllFlashmail()
    'this sub will delete all of your flashmail
    'Example of use:
    '   Call DeleteAllFlashmail
    Dim TheWin As Long
    TheWin& = FindFlashMailBox&
    If TheWin& = 0& Then
        If FindWindow("AOL Frame25", vbNullString) = 0& Then Exit Sub
        Call LoadFlashmail
        Do
            DoEvents
            TheWin& = FindFlashMailBox&
        Loop Until (TheWin& <> 0&)
        Call Sleep(2000)
        DoEvents
    End If
    Dim TheTree As Long, DeleteIcon As Long, OKMsg As Long
    Dim OKMsgBtn As Long, TreeCount As Long, OldTreeCount As Long
    TheTree& = FindWindowEx(TheWin&, 0&, "_AOL_Tree", vbNullString)
    DeleteIcon& = FindWindowEx(TheWin&, 0&, "_AOL_Icon", vbNullString)
    DeleteIcon& = FindWindowEx(TheWin&, DeleteIcon&, "_AOL_Icon", vbNullString)
    DeleteIcon& = FindWindowEx(TheWin&, DeleteIcon&, "_AOL_Icon", vbNullString)
    DeleteIcon& = FindWindowEx(TheWin&, DeleteIcon&, "_AOL_Icon", vbNullString)
    OldTreeCount& = SendMessage(TheTree&, LB_GETCOUNT, 0&, 0&)
    If OldTreeCount& = 0& Then Exit Sub
    TreeCount& = OldTreeCount&
    Do
        Call SendMessage(TheTree&, LB_SETCURSEL, 0&, 0&)
        DoEvents
        Call PostMessage(DeleteIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(DeleteIcon&, WM_LBUTTONUP, 0&, 0&)
        Do
            DoEvents
            OKMsg& = FindWindow("#32770", "America Online")
            OKMsgBtn& = FindWindowEx(OKMsg&, 0&, "Button", vbNullString)
        Loop Until (OKMsgBtn& <> 0& And IsWindowVisible(OKMsgBtn&) = 1)
        Call SendMessage(OKMsgBtn&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKMsgBtn&, WM_KEYUP, VK_SPACE, 0&)
        Do
            DoEvents
            TreeCount& = SendMessage(TheTree&, LB_GETCOUNT, 0&, 0&)
        Loop Until OldTreeCount& > TreeCount&
        OldTreeCount& = TreeCount&
    Loop Until OldTreeCount& = 0&
End Sub

Public Function FileScanEMAIL(ByVal TheFile As String) As String
    'this function will extract email addresses from a file
    'Example of use:
    '   MsgBox FileScanEMAIL("C:\Test.exe")
    Dim LOFF As Long, LoopThrough As Long, CurrentString As String
    Dim StringsFound As String, FileHandle As Integer, TheStep As Long
    Dim TheINSTR As Long, LoopAgain As Long, TheMid As String
    Dim TheASC As Integer, DummyString As String
    FileHandle% = FreeFile
    Open TheFile$ For Binary Access Read As #FileHandle%
    DoEvents
    LOFF& = LOF(FileHandle%)
    CurrentString$ = String(LOFF&, 0&)
    TheStep& = 32000 '32000, 64000, and 8420 are good.  the lower the number
                    'the more thurough the scan
    For LoopThrough& = 1 To LOFF& Step TheStep&
        Get #FileHandle%, LoopThrough&, CurrentString$
        DoEvents
        CurrentString$ = CharLower(CurrentString$)
        TheINSTR& = InStr(1, CurrentString$, ".com", 1)
        If TheINSTR& <> 0& Then
GetAddress:
            DummyString$ = ""
            For LoopAgain& = (3 + TheINSTR&) To (30 - LoopAgain&) Step -1
                TheMid$ = Mid(CurrentString$, LoopAgain&, 1)
                TheASC% = Asc(TheMid$)
                If TheMid$ = " " Or TheASC% < 45 Or TheASC% > 255 Then Exit For
                If TheMid$ <> Chr(0) Then
                    DummyString$ = TheMid$ & DummyString$
                End If
            Next LoopAgain&
            If InStr(1, DummyString$, "@", 1) <> 0 Then
                StringsFound$ = StringsFound$ & DummyString$ & vbCrLf
            End If
            TheINSTR& = InStr(1 + TheINSTR&, CurrentString$, ".com", 1)
            If TheINSTR& <> 0& Then GoTo GetAddress
        End If
    Next LoopThrough&
    Close #FileHandle%
    DoEvents
    FileScanEMAIL$ = StringsFound$
End Function
Public Sub RunMenuByString(ByVal TheWindow As Long, ByVal MenuString As String)
    'this sub searches a window for a menu item containing MenuString
    'if found, it will click the menu
    'Example of use:
    '   Call RunMenuByString(FindAIM&, "E&xit")
    Dim aMenu As Long, mCount As Long
    Dim LoopThrough As Long, sMenu As Long, sCount As Long
    Dim LoopThrough2 As Long, sID As Long, sString As String
    Dim TheSearch As String
    TheSearch$ = CharLower(MenuString$)
    aMenu& = GetMenu(TheWindow&)
    mCount& = GetMenuItemCount(aMenu&)
    For LoopThrough& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LoopThrough&)
        sCount& = GetMenuItemCount(sMenu&)
        For LoopThrough2& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LoopThrough2&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(1, CharLower(sString$), TheSearch$, 1) <> 0& Then
                Call SendMessageLong(TheWindow&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LoopThrough2&
    Next LoopThrough&
End Sub

Public Function FindAOLToolbar() As Long
    'this function will find AOL's Toolbar
    'Example of use:
    '   MsgBox FindAOLToolbar&
    Dim TlBar1 As Long
    TlBar1& = FindWindow("AOL Frame25", vbNullString)
    If TlBar1& = 0& Then
        FindAOLToolbar& = 0&
        Exit Function
    End If
    TlBar1& = FindWindowEx(TlBar1&, 0&, "AOL Toolbar", vbNullString)
    FindAOLToolbar& = FindWindowEx(TlBar1&, 0&, "_AOL_Toolbar", vbNullString)
End Function

Public Function FindIM() As Long
    'this function finds the top AOL IM
    'Example of use:
    '   MsgBox GetCaption(FindIM&)
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindIM& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, TheINSTR As Long, TheCaption As String
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        TheCaption$ = GetCaption(AOLChild&)
        TheINSTR& = InStr(1, TheCaption$, "Instant Message", 1)
        If (TheINSTR& = 1) Or (TheINSTR& = 2) Or (TheINSTR& = 3) Then
            FindIM& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindIM& = 0&
End Function

Public Function GetUser() As String
    'this function gets the users AOL screen name from the Welcome window
    'Example of use:
    '   MsgBox GetUser$
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        GetUser$ = ""
        Exit Function
    End If
    Dim AOLChild As Long, TheCaption As String, TheINSTR As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        TheCaption$ = GetCaption(AOLChild&)
        If Left(TheCaption$, 9) = "Welcome, " Then
            TheINSTR& = InStr(10, TheCaption$, "!", 1)
            GetUser$ = Mid(TheCaption$, 10, TheINSTR& - 10)
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    GetUser$ = ""
End Function


Public Sub IMsOff()
    'this sub turns your AOL IMs off
    'Example of use:
    '   Call IMsOff
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    Dim AOLChild As Long, AOLIcon As Long, AOLEdit As Long, Rich As Long
    Dim ACounter As Long, AOLMsg As Long, AOLMsgButton As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:$IM_off")
    Do
        DoEvents
        AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Loop Until (GetText(AOLEdit&) = "$IM_off") And (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "http://come.to/cruelair")
ClickSend:
    ACounter& = Timer
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        AOLMsg& = FindWindow("#32770", "America Online")
        AOLMsgButton& = FindWindowEx(AOLMsg&, 0&, "Button", vbNullString)
        If (Timer - ACounter&) > 3 Then GoTo ClickSend
    Loop Until (AOLMsg& <> 0&)
    Do
        DoEvents
        Call SendMessage(AOLMsgButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(AOLMsgButton&, WM_KEYUP, VK_SPACE, 0&)
        AOLMsg& = FindWindow("#32770", "America Online")
        AOLMsgButton& = FindWindowEx(AOLMsg&, 0&, "Button", vbNullString)
        DoEvents
    Loop Until AOLMsg& = 0&
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
    DoEvents

End Sub

Public Sub IMsOn()
    'this sub turns your AOL IMs on
    'Example of use:
    '   Call IMsOn
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    Dim AOLChild As Long, AOLIcon As Long, AOLEdit As Long, Rich As Long
    Dim ACounter As Long, AOLMsg As Long, AOLMsgButton As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:$IM_on")
    Do
        DoEvents
        AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Loop Until (GetText(AOLEdit&) = "$IM_on") And (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLIcon&) = 1)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "http://come.to/cruelair")
ClickSend:
    ACounter& = Timer
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        AOLMsg& = FindWindow("#32770", "America Online")
        AOLMsgButton& = FindWindowEx(AOLMsg&, 0&, "Button", vbNullString)
        If (Timer - ACounter&) > 3 Then GoTo ClickSend
    Loop Until (AOLMsg& <> 0&)
    Do
        DoEvents
        Call SendMessage(AOLMsgButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(AOLMsgButton&, WM_KEYUP, VK_SPACE, 0&)
        AOLMsg& = FindWindow("#32770", "America Online")
        AOLMsgButton& = FindWindowEx(AOLMsg&, 0&, "Button", vbNullString)
        DoEvents
    Loop Until AOLMsg& = 0&
    Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
    DoEvents
End Sub
Public Sub KillAllModals()
    'this sub will close any and all _AOL_Modal windows that are open
    'Example of use:
    '   Call KillAllModals
    Dim TheWin As Long
    TheWin& = FindWindow("_AOL_Modal", vbNullString)
    Do While TheWin& <> 0&
        Call PostMessage(TheWin&, WM_CLOSE, 0&, 0&)
        DoEvents
        TheWin& = FindWindow("_AOL_Modal", vbNullString)
    Loop
End Sub
Public Function GetRGB(ByVal CVal As Long) As COLORRGB
    'this function puts the Red, Green, and Blue values of a color
    'into a variable of the type COLORRGB
    'Example of use:
    '   Dim TheRGB As COLORRGB
    '   TheRGB = GetRGB(GetPixelColor)
    '   MsgBox TheRGB.Red & vbCrLf & TheRGB.Green & vbCrLf & TheRGB.Blue
    GetRGB.Blue = Int(CVal / 65536)
    GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
    GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function
Public Function GetClass(ByVal Win As Long) As String
    'this functil will return the class of a window
    'Example of use:
    '   MsgBox GetClass(FindChatRoom&)
    Dim TheClass As String, TheINSTR As Integer
    TheClass$ = String(256, 0&)
    Call GetClassName(Win&, TheClass$, 256)
    TheClass$ = Trim(TheClass$)
    TheINSTR% = InStr(1, TheClass$, Chr(0), 1)
    If TheINSTR% <> 0 Then TheClass$ = Mid(TheClass$, 1, TheINSTR% - 1)
    GetClass$ = TheClass$
End Function
Public Sub UpChat()
    'this sub will allow you to upload and do other things on AOL at the same time
    'Example of use:
    '   Call Upchat
    Dim AOMod As Long
    AOMod& = FindWindow("_AOL_Modal", vbNullString)
    If AOMod& = 0& Then Exit Sub
    If InStr(GetCaption(AOMod&), "File Transfer") = 0 Then Exit Sub
    Dim AOL As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    Call EnableWindow(AOMod&, 0)
    Call EnableWindow(AOL&, 1)
End Sub
Public Sub UnUpChat()
    'this sub will undo the Upchat sub
    'Example of use:
    '   Call UnUpchat
    Dim AOMod As Long
    AOMod& = FindWindow("_AOL_Modal", vbNullString)
    Dim AOL As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    If AOMod& = 0& Then
        Call EnableWindow(AOL&, 1&)
        Exit Sub
    End If
    If InStr(1, GetCaption(AOMod&), "File Transfer", 1) = 0 Then
        Call EnableWindow(AOL&, 1&)
        Exit Sub
    End If
    Call EnableWindow(AOL&, 0&)
    Call EnableWindow(AOMod&, 1&)
End Sub

Public Sub AIMKeyword(ByVal TheKW As String)
    'this sub uses a keyword via AIM
    'Example of use:
    '   Call AIMKeyword("http://come.to/cruelair")
    Dim Edit As Long
    Edit& = FindWindowEx(FindAIM&, 0&, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, TheKW$)
    DoEvents
    Call SendMessageLong(Edit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(Edit&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Function FindAIM() As Long
    'this function finds AIM's main window
    'Example of use:
    '   MsgBox GetCaption(FindAIM&)
    FindAIM& = FindWindow("_Oscar_BuddyListWin", vbNullString)
End Function
Public Function LineCountString(ByVal TheString As String) As Long
    'this function will count the number of lines in a string
    'Example of use:
    '   Dim AString As String
    '   AString$ = "test" & vbCrLf & "test"
    '   MsgBox LineCountString(AString$)
    Dim TheINSTR As Long, TheCount As Long
    TheCount& = 0&
    TheINSTR& = InStr(1, TheString$, Chr(13), 1)
    If TheINSTR& <> 0& Then
        Do
            TheINSTR& = InStr(1 + TheINSTR&, TheString$, Chr(13), 1)
            If TheINSTR& <> 0& Then
                TheCount& = 1 + TheCount&
            Else
                Exit Do
            End If
        Loop
    End If
    LineCountString& = 1 + TheCount&
End Function

Public Sub VPKeyword(ByVal strKeyword As String)
    'this sub calls a keyword in Virtual Places
    'Example of use:
    '   Call VPKeyword("http://come.to/cruelair")
    Dim VPWin As Long, Edit As Long
    VPWin& = FindWindow("VPFrame", vbNullString)
    If VPWin& = 0& Then Exit Sub
    Edit& = FindWindowEx(VPWin&, 0&, "AfxMDIFrame42s", vbNullString)
    Edit& = FindWindowEx(Edit&, 0&, "AfxMDIFrame42s", vbNullString)
    Edit& = FindWindowEx(Edit&, 0&, "AfxFrameOrView42s", vbNullString)
    Edit& = FindWindowEx(Edit&, 0&, "#32770", vbNullString)
    Edit& = FindWindowEx(Edit&, 0&, "ComboBox", vbNullString)
    Edit& = FindWindowEx(Edit&, 0&, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, strKeyword$)
    DoEvents
    Call SendMessageLong(Edit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(Edit&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub VPSendChat(ByVal Chat As String)
    'this sub sends chat to Virtual Places
    'Example of use:
    '   Call VPSendChat("http://come.to/cruelair")
    Dim VPMain As Long, VPMIDI As Long, VPAfxFrame As Long
    Dim Num32770 As Long, Edit As Long, Button As Long
    VPMain& = FindWindow("VPFrame", vbNullString)
    If VPMain& = 0& Then Exit Sub
    VPMIDI& = FindWindowEx(VPMain&, 0&, "AfxMDIFrame42s", vbNullString)
    VPMIDI& = FindWindowEx(VPMIDI&, 0&, "AfxMDIFrame42s", vbNullString)
    VPAfxFrame& = FindWindowEx(VPMIDI&, 0&, "AfxFrameOrView42s", vbNullString)
    VPAfxFrame& = FindWindowEx(VPMIDI&, VPAfxFrame&, "AfxFrameOrView42s", vbNullString)
    Num32770& = FindWindowEx(VPAfxFrame&, 0&, "#32770", vbNullString)
    Edit& = FindWindowEx(Num32770&, 0&, "Edit", vbNullString)
    Button& = FindWindowEx(Num32770&, 0&, "Button", vbNullString)
    If (Edit& = 0&) Or (Button& = 0&) Then Exit Sub
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, Chat$)
    DoEvents
    Call SendMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub PutImage(ByRef TakeFrom As PictureBox, ByRef GiveTo As PictureBox)
    'this sub resizes the image in one picturebox to fit another picturebox
    'this is VERY useful for skins
    'Example of use:
    '   Call PutImage(Picture1, Picture2)
    Dim Dest_hDC As Long, Src_hDC As Long, StretchBit As Long
    Dest_hDC = GiveTo.hdc
    Src_hDC = TakeFrom.hdc
    TakeFrom.AutoRedraw = True
    TakeFrom.ScaleMode = 3
    GiveTo.ScaleMode = 3
    Call SetStretchBltMode(Dest_hDC, STRETCH_DELETESCANS)
    StretchBit = StretchBlt(Dest_hDC, 0, 0, GiveTo.ScaleWidth, GiveTo.ScaleHeight, Src_hDC, 0, 0, TakeFrom.ScaleWidth, TakeFrom.ScaleHeight, SRCCOPY)
    TakeFrom.Refresh
    GiveTo.AutoRedraw = True
    Call PutImage2(TakeFrom, GiveTo)
End Sub

Public Sub PutImage2(ByRef TakeFrom As PictureBox, ByRef GiveTo As PictureBox)
    'this is just the second part of the PutImage sub
    'Example of use:
    '   [just use the PutImage sub]
    Dim Dest_hDC As Long, Src_hDC As Long, StretchBit As Long
    Dest_hDC = GiveTo.hdc
    Src_hDC = TakeFrom.hdc
    TakeFrom.AutoRedraw = True
    TakeFrom.ScaleMode = 3
    GiveTo.ScaleMode = 3
    Call SetStretchBltMode(Dest_hDC, STRETCH_DELETESCANS)
    StretchBit = StretchBlt(Dest_hDC, 0, 0, GiveTo.ScaleWidth, GiveTo.ScaleHeight, Src_hDC, 0, 0, TakeFrom.ScaleWidth, TakeFrom.ScaleHeight, SRCCOPY)
    TakeFrom.Refresh
    GiveTo.AutoRedraw = True
End Sub
Public Function Percent(ByVal Complete As Long, ByVal Total As Long, ByVal TotalOutput As Long) As Long
    'this function returns a precentage
    'Example of use:
    '   MsgBox Percent(5,10,100)
    On Error Resume Next
    Percent& = CLng((Complete& / Total&) * TotalOutput&)
End Function
Public Function RandomNumber(ByVal BiggestNum As Long, ByVal AllowZero As Boolean) As Long
    'this function returns a random number
    'Example of use:
    '   MsgBox RandomNumber(10, False)
    Randomize
    RandomNumber& = (CLng(BiggestNum&) * Rnd)
    If RandomNumber& = 0& And AllowZero = False Then
        RandomNumber& = 1&
    End If
End Function
Public Function RotateCase(ByVal TheString As String, ByVal SkipSpaces As Boolean) As String
    'this function will rotate the casing of a string
    'Example of use:
    '   MsgBox RotateCase("http://come.to/cruelair", False)
    Dim aVal As Byte, LoopThrough As Long, FinalString As String
    Dim CurLetter As String
    CurLetter$ = Mid(TheString$, 1, 1)
    If IsCharLower(Asc(CurLetter$)) = 1 Then
        aVal = 0
    Else
        aVal = 1
    End If
    FinalString$ = ""
    CurLetter$ = ""
    For LoopThrough& = 1 To Len(TheString$)
        CurLetter$ = Mid$(TheString$, LoopThrough&, 1)
        If SkipSpaces = True And CurLetter$ = " " Then GoTo SkipRotate
        If aVal = 0 Then
            CurLetter$ = CharLower(CurLetter$)
            aVal = 1
        Else
            CurLetter$ = CharUpper(CurLetter$)
            aVal = 0
        End If
        FinalString$ = FinalString$ & CurLetter$
SkipRotate:
    Next LoopThrough&
    RotateCase$ = FinalString$
End Function

Public Function FindBuddyList() As Long
    'this function will find AOL's buddy list
    'Example of use:
    '   MsgBox FindBuddyList&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindBuddyList& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, AOLListbox As Long
    Dim Rich As Long, AOLIcon2 As Long, AOLListbox2 As Long
    Dim AOLIcon3 As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
        AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
        AOLListbox2& = FindWindowEx(AOLChild&, AOLListbox&, "_AOL_Listbox", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        If (IsWindowVisible(AOLIcon&) = 1) And (IsWindowVisible(AOLIcon2&) = 0) And (AOLIcon3& = 0&) And (IsWindowVisible(AOLListbox&) = 1) And (AOLListbox2& = 0&) And (Rich& = 0&) Then
            FindBuddyList& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindBuddyList& = 0&
End Function

Public Function FindChatRoom() As Long
    'this function will find AOL's chat room
    'Example of use:
    '   MsgBox FindChatRoom&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindChatRoom& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, Rich As Long
    Dim AOLListbox As Long, AOLListbox2 As Long, AOLCombo As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        Rich& = FindWindowEx(AOLChild&, Rich&, "RICHCNTL", vbNullString)
        AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
        AOLListbox2& = FindWindowEx(AOLChild&, AOLListbox&, "_AOL_Listbox", vbNullString)
        AOLCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
        If (IsWindowVisible(AOLIcon&) = 1) And (IsWindowVisible(Rich&) = 1) And (IsWindowVisible(AOLListbox&) = 1) And (AOLListbox2& = 0&) And (IsWindowVisible(AOLCombo&) = 1) Then
            FindChatRoom& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindChatRoom& = 0&
End Function

Public Function FindFlashMailBox() As Long
    'this function finds AOL's flashmail box
    'Example of use:
    '   MsgBox FindFlashmailBox
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindFlashMailBox& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, AOLTree As Long
    Dim AOLIcon2 As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
        If (IsWindowVisible(AOLIcon&) = 1) And (IsWindowVisible(AOLTree&) = 1) And (AOLIcon2& = 0&) Then
            FindFlashMailBox& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "_AOL_Icon", vbNullString)
    Loop Until AOLChild& = 0&
    FindFlashMailBox& = 0&
End Function
Public Function FindSignOnWindowAIM() As Long
    'finds AIM's sign on window
    'one problem, this dosent work if the window is minimized.
    'if you want, just add something to restore the window, then check to
    'see if its the one your looking for
    'Example of use:
    '   MsgBox FindSignOnWindowAIM&
    Dim TheWin As Long
    TheWin& = FindWindow("#32770", vbNullString)
    If TheWin& = 0& Then
        FindSignOnWindowAIM& = 0&
        Exit Function
    End If
    Dim AIMStatic As Long, CBoxEdit As Long, PWEdit As Long, AIMButton As Long
    AIMStatic& = FindWindowEx(TheWin&, 0&, "Static", vbNullString)
    AIMStatic& = FindWindowEx(TheWin&, AIMStatic&, "Static", vbNullString)
    AIMStatic& = FindWindowEx(TheWin&, AIMStatic&, "Static", vbNullString)
    AIMStatic& = FindWindowEx(TheWin&, AIMStatic&, "Static", vbNullString)
    AIMStatic& = FindWindowEx(TheWin&, AIMStatic&, "Static", vbNullString)
    CBoxEdit& = FindWindowEx(TheWin&, 0&, "Combobox", vbNullString)
    CBoxEdit& = FindWindowEx(CBoxEdit&, 0&, "Edit", vbNullString)
    PWEdit& = FindWindowEx(TheWin&, 0&, "Edit", vbNullString)
    AIMButton& = FindWindowEx(TheWin&, 0&, "Button", vbNullString)
    AIMButton& = FindWindowEx(TheWin&, AIMButton&, "Button", vbNullString)
    AIMButton& = FindWindowEx(TheWin&, AIMButton&, "Button", vbNullString)
    If (IsWindowVisible(AIMStatic&) = 1) And (IsWindowVisible(CBoxEdit&) = 1) And (IsWindowVisible(PWEdit&) = 1) And (IsWindowVisible(AIMButton&) = 1) Then
        FindSignOnWindowAIM& = TheWin&
    Else
        FindSignOnWindowAIM& = 0&
    End If
End Function

Public Function FindWriteMail() As Long
    'this function finds the Write Mail window
    'Example of use:
    '   MsgBox FindWriteMail&
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindWriteMail& = 0&
        Exit Function
    End If
    Dim AOLChild As Long, AOLIcon As Long, AOLEdit As Long
    Dim Rich As Long, AOLEdit2 As Long, Rich2 As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
        AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        Rich2& = FindWindowEx(AOLChild&, Rich&, "RICHCNTL", vbNullString)
        If (IsWindowVisible(AOLIcon&) = 1) And (IsWindowVisible(AOLEdit&) = 1) And (AOLEdit2& = 0&) And (IsWindowVisible(Rich&) = 1) And (Rich2& = 0&) Then
            FindWriteMail& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindWriteMail& = 0&
End Function
Public Function GetAOLToolbarIcon(ByVal IconIndex As Byte) As Long
    'this function will get a _AOL_Icon from AOL's toolbar
    'Example of use:
    '   MsgBox GetAOLToolbarIcon(2)
    Dim TBar As Long
    TBar& = FindAOLToolbar&
    If TBar& = 0& Then
        GetAOLToolbarIcon& = 0&
        Exit Function
    End If
    Dim AOLIco As Long, LoopThrough As Byte
    AOLIco& = FindWindowEx(TBar&, 0&, "_AOL_Icon", vbNullString)
    For LoopThrough = 2 To IconIndex
        AOLIco& = FindWindowEx(TBar&, AOLIco&, "_AOL_Icon", vbNullString)
    Next LoopThrough
    GetAOLToolbarIcon& = AOLIco&
End Function

Public Function GetCaption(ByVal TheWin As Long) As String
    'use this for getting the captions of windows instead of the GetText function
    'Example of use:
    '   MsgBox GetCaption(FindChatRoom&)
    Dim TheLength As Long, TheCaption As String
    TheLength& = GetWindowTextLength(TheWin&)
    TheLength& = 1 + TheLength&
    TheCaption$ = String(TheLength&, 0&)
    Call GetWindowText(TheWin&, TheCaption$, TheLength&)
    GetCaption$ = TheCaption$
End Function

Public Sub GetChatRoom(ByRef TheWindow As Long, ByRef RichBox As Long, ByRef SendIcon As Long)
    'this sub will find the chat room, the richcntl, and the send icon
    'Example of use:
    '   Dim TheRoom As Long, TheRICH As Long, SendIcon As Long
    '   Call GetChatRoom(TheRoom&, TheRICH&, SendIcon&)
    '   MsgBox TheRoom & vbCrLf & TheRICH & vbCrLf & SendIcon
    Dim TheWin As Long
    TheWin& = FindChatRoom&
    If TheWin& = 0& Then
        TheWindow& = 0&
        RichBox& = 0&
        SendIcon& = 0&
        Exit Sub
    End If
    Dim DummyWin As Long
    DummyWin& = FindWindowEx(TheWin&, 0&, "RICHCNTL", vbNullString)
    RichBox& = FindWindowEx(TheWin&, DummyWin&, "RICHCNTL", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, 0&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    SendIcon& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    TheWindow& = TheWin&
End Sub

Public Function GetPixelColor() As Long
    'this function will get the color under the mouse, no matter
    'where the mouse is
    'Example of use: (maybe in a timer)
    '   Picture1.BackColor = GetPixelColor&
    Dim CPos As POINTAPI, TheRECT As RECT
    Dim TheDC As Long, TheWindow As Long
    Dim LeftVal As Long, TopVal As Long
    Call GetCursorPos(CPos)
    TheWindow& = WindowFromPoint(CPos.X, CPos.Y)
    Call GetWindowRect(TheWindow&, TheRECT)
    LeftVal& = CPos.X - TheRECT.Left
    TopVal& = CPos.Y - TheRECT.Top
    TheDC& = GetWindowDC(TheWindow&)
    GetPixelColor& = GetPixel(TheDC&, LeftVal&, TopVal&)
End Function

Public Sub GetSignOnWindow(ByRef TheWindow As Long, ByRef NamesCombo As Long, ByRef PasswordEdit As Long, ByRef SignOnIcon As Long)
    'this sub will find the sign on window and put the handles
    'of it, the name combobox, the password editbox, and the sign on icon into
    'four diffrent variables.
    'Example of use:
    '   Dim SOWin As Long, NCombo As Long, PWEdit As Long, SOIco As Long
    '   Call GetSignOnWindow(SOWin, NCombo, PWEdit, SOIco)
    '   MsgBox SOWin & vbCrLf & NCombo & vbCrLf & PWEdit & vbCrLf & SOIco
    Dim TheWin As Long
    TheWin& = FindSignOnWindow&
    If TheWin& = 0& Then
        TheWindow& = 0&
        NamesCombo& = 0&
        PasswordEdit& = 0&
        SignOnIcon& = 0&
        Exit Sub
    End If
    Dim DummyWin As Long, DummyWin2 As Long
    DummyWin& = FindWindowEx(TheWin&, 0&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin2& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    If DummyWin2& = 0& Then
        SignOnIcon& = DummyWin&
    Else
        SignOnIcon& = DummyWin2&
    End If
    DummyWin& = FindWindowEx(TheWin&, 0&, "_AOL_Combobox", vbNullString)
    NamesCombo& = DummyWin&
    DummyWin& = FindWindowEx(TheWin&, 0&, "_AOL_Edit", vbNullString)
    PasswordEdit& = DummyWin&
    TheWindow = TheWin&
End Sub

Public Sub GetSignOnWindowAIM(ByRef TheWindow As Long, ByRef NameCombo As Long, ByRef NameEdit As Long, ByRef PWEdit As Long, ByRef SignOnButton As Long)
    'this sub will find aim's sign on window and put the handles
    'of it, the name combobox, the name editbox, the password editbox, and the sign on button into
    'five diffrent variables.
    'Example of use:
    '   Dim SOWin As Long, NCombo As Long, PWEdit As Long, SOBtn As Long, NEdit As Long
    '   Call GetSignOnWindowAIM(SOWin, NCombo, NEdit, PWEdit, SOBtn)
    '   MsgBox SOWin & vbCrLf & NCombo & vbCrLf & PWEdit & vbCrLf & SOBtn & vbCrLf & NEdit
    Dim TheWin As Long
    TheWin& = FindSignOnWindowAIM&
    If TheWin& = 0& Then
        TheWindow& = 0&
        NameCombo& = 0&
        NameEdit& = 0&
        PWEdit& = 0&
        SignOnButton& = 0&
        Exit Sub
    End If
    Dim DummyWin As Long
    NameCombo& = FindWindowEx(TheWin&, 0&, "Combobox", vbNullString)
    NameEdit& = FindWindowEx(NameCombo&, 0&, "Edit", vbNullString)
    PWEdit& = FindWindowEx(TheWin&, 0&, "Edit", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, 0&, "Static", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "Static", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "Static", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "Static", vbNullString)
    SignOnButton& = FindWindowEx(TheWin&, DummyWin&, "Static", vbNullString)
    TheWindow& = TheWin&
End Sub

Public Function GetText(ByVal TheWin As Long) As String
    'this function will get the text from a window
    'Example of use:
    '   MsgBox GetText(GetAOLToolbarIcon(2))
    Dim TheString As String, TheLength As Long
    TheLength& = SendMessage(TheWin&, WM_GETTEXTLENGTH, 0&, 0&)
    TheString$ = String(TheLength&, 0&)
    Call SendMessageByString(TheWin&, WM_GETTEXT, 1 + TheLength&, TheString$)
    GetText$ = TheString$
End Function
Public Sub GetWriteMail(ByRef TheWindow As Long, ByRef ToWhoEdit As Long, ByRef SubjectEdit As Long, ByRef Rich As Long, ByRef SendIcon As Long)
    'this sub will return the handle of the Write Mail Window, the
    'To editbox, the subject editbox, the RICHCNTL, and the Send icon
    'Example of use:
    '   Dim MailWin As Long, ToEdit As Long, SubEdit As Long, SendIcon As Long, RICH As Long
    '   Call GetWriteMail(MailWin&, ToEdit&, SubEdit&, RICH&, SendIcon&)
    '   MsgBox MailWin & vbCrLf & ToEdit & vbCrLf & SubEdit & vbCrLf & RICH & vbCrLf & SendIcon
    Dim TheWin As Long
    TheWin& = FindWriteMail&
    If TheWin& = 0& Then
        TheWindow& = 0&
        ToWhoEdit& = 0&
        SubjectEdit& = 0&
        SendIcon& = 0&
        Rich& = 0&
        Exit Sub
    End If
    Dim DummyWin As Long
    ToWhoEdit& = FindWindowEx(TheWin&, 0&, "_AOL_Edit", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, ToWhoEdit&, "_AOL_Edit", vbNullString)
    SubjectEdit& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Edit", vbNullString)
    Rich& = FindWindowEx(TheWin&, 0&, "RICHCNTL", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, 0&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    DummyWin& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    SendIcon& = FindWindowEx(TheWin&, DummyWin&, "_AOL_Icon", vbNullString)
    TheWindow& = TheWin&
End Sub
Public Sub Keyword(ByVal TheKeyWord As String)
    'this will call a keyword (or web address) from aol
    'Example of use:
    '   Call Keyword("http://come.to/cruelair")
    Dim EditBox As Long
    EditBox& = FindAOLToolbar&
    EditBox& = FindWindowEx(EditBox&, 0&, "_AOL_Combobox", vbNullString)
    If IsWindowEnabled(EditBox&) = 0 Then Exit Sub
    EditBox& = FindWindowEx(EditBox&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditBox&, WM_SETTEXT, 0&, TheKeyWord$)
    Call SendMessageLong(EditBox&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditBox&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub LoadFlashmail()
    'this sub will load AOL's flashmail box
    'Example of use:
    '   Call LoadFlashmail
    Dim TBarIco As Long
    TBarIco& = GetAOLToolbarIcon(3)
    If TBarIco& = 0& Then Exit Sub
    Dim MenuWin As Long, CPos As POINTAPI
    Call GetCursorPos(CPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(TBarIco&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(TBarIco&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        MenuWin& = FindWindow("#32768", vbNullString)
    Loop Until (IsWindowVisible(MenuWin&) = 1)
    Call PostMessage(MenuWin&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(MenuWin&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(MenuWin&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call PostMessage(MenuWin&, WM_KEYUP, VK_RIGHT, 0&)
    Call PostMessage(MenuWin&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(MenuWin&, WM_KEYUP, VK_RETURN, 0&)
    DoEvents
    Call SetCursorPos(CPos.X, CPos.Y)
End Sub
Public Function RotateWord(ByVal TheWord As String) As String
    'this function will take the first letter of the word, and put it at the end
    'Example of use:
    '   MsgBox RotateWord("rhttp://come.to/cruelai")
    RotateWord$ = Right(TheWord$, Len(TheWord$) - 1) & Left(TheWord$, 1)
End Function

Public Sub SendChat(ByVal ChatString As String)
    'this sub will send text to AOL's chat room
    'Example of use:
    '   Call SendChat("http://come.to/cruelair")
    Dim ChatRoom As Long, SendIcon As Long, Rich As Long
    Dim TheLength As Long
    Call GetChatRoom(ChatRoom&, Rich&, SendIcon&)
    If ChatRoom& = 0& Then Exit Sub
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, ChatString$)
    Call SendMessageLong(Rich&, WM_CHAR, 13&, 0&)
    DoEvents
    TheLength& = SendMessage(Rich&, WM_GETTEXTLENGTH, 0&, 0&)
    If TheLength& <> 0& Then
        Call SendMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call SendMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
        DoEvents
        TheLength& = SendMessage(Rich&, WM_GETTEXTLENGTH, 0&, 0&)
        If TheLength& <> 0& Then
            Call SendMessageLong(Rich&, WM_CHAR, 13&, 0&)
            DoEvents
            Call SendMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            Call SendMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub

Public Sub SendChatMulti(ByVal TheStrings As Variant, ByVal PauseBetween As Boolean, Optional ByVal PauseDuration As Long)
    'use this for sending multiple lines of chat to the room
    'using this gives you less control over how much of a diffrence is between
    'each SendChat, but using this also means the program itself has to do
    'much less work.
    'Example of use:
    '    Dim ChatStrings(3) As Variant
    '    ChatStrings(1) = "http://come.to/cruelair"
    '    ChatStrings(2) = "crueizme@hotmail.com"
    '    ChatStrings(3) = "http://www.blaupunkt.com"
    '    Call SendChatMulti(ChatStrings, False)
    '    Call SendChatMulti(chatStrings, True, 2)
    Dim ChatRoom As Long, SendIcon As Long, Rich As Long
    Dim TheLength As Long, LoopThrough As Long
    Call GetChatRoom(ChatRoom&, Rich&, SendIcon&)
    If ChatRoom& = 0& Then Exit Sub
    For LoopThrough& = 1 To UBound(TheStrings)
        DoEvents
        Call SendMessageByString(Rich&, WM_SETTEXT, 0&, TheStrings(LoopThrough&))
        Call SendMessageLong(Rich&, WM_CHAR, 13&, 0&)
        DoEvents
        TheLength& = SendMessage(Rich&, WM_GETTEXTLENGTH, 0&, 0&)
        If TheLength& <> 0& Then
            Call SendMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
            TheLength& = SendMessage(Rich&, WM_GETTEXTLENGTH, 0&, 0&)
            If TheLength& <> 0& Then
                Call SendMessageLong(Rich&, WM_CHAR, 13&, 0&)
                DoEvents
                Call SendMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
                DoEvents
                Call SendMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
                Call SendMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
            End If
        End If
        If PauseBetween = True Then
            Call Sleep(PauseDuration * 1000)
            DoEvents
        End If
    Next LoopThrough&
End Sub




Public Sub SendIM(ByVal ToWho As String, ByVal Message As String)
    'this sub uses AOL to send an instant message (using a keyword)
    'Example of use:
    '   Call SendIM("i am crue", "hello")
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    Dim AOLChild As Long, AOLIcon As Long, AOLEdit As Long, Rich As Long
    Dim ACounter As Long, AOLMsg As Long
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & ToWho$)
    Do
        DoEvents
        AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        Rich& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Loop Until (AOLEdit& <> 0& And GetText(AOLEdit&) = ToWho$) And (Rich& <> 0& And IsWindowVisible(Rich&) = 1) And (AOLIcon& <> 0& And IsWindowVisible(AOLIcon&) = 1)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
ClickSend:
    ACounter& = Timer
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        AOLMsg& = FindWindow("#32770", "America Online")
        If (Timer - ACounter&) > 3 Then GoTo ClickSend
    Loop Until (AOLEdit& = 0&) Or (AOLMsg& <> 0&)
    If AOLMsg& <> 0& Then
        Dim AOLMsgButton As Long
        AOLMsgButton& = FindWindowEx(AOLMsg&, 0&, "Button", vbNullString)
        Call SendMessage(AOLMsgButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(AOLMsgButton&, WM_KEYUP, VK_SPACE, 0&)
        DoEvents
        Call PostMessage(AOLMsg&, WM_CLOSE, 0&, 0&)
        DoEvents
        Call PostMessage(AOLChild&, WM_CLOSE, 0&, 0&)
        DoEvents
    End If
    
End Sub
Public Sub SendMail(ByVal ToWho As String, ByVal TheSubject As String, ByVal TheMessage As String)
    'this sub sends a mail
    'Example of use:
    '   Call SendMail("CrueIzMe@Hotmail.com", "hey", "hello [and some other stuff] `,:^)")
    Dim MDIClient As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then Exit Sub
    Dim TheWin As Long, SendIcon As Long, ToEdit As Long
    Dim Rich As Long, SubjectEdit As Long, ACounter As Long
    Dim DummyWin As Long
    Call GetWriteMail(TheWin&, ToEdit&, SubjectEdit&, Rich&, SendIcon&)
    If (ToEdit& <> 0&) And (SubjectEdit& <> 0&) And (Rich& <> 0&) And (SendIcon& <> 0&) Then
        Call SendMessageByString(ToEdit&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(SubjectEdit&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "")
        GoTo EnterInformation
    End If
    ToEdit& = 0&
    SubjectEdit& = 0&
    Rich& = 0&
    SendIcon& = 0&
    Dim TBarIcon As Long
    TBarIcon& = GetAOLToolbarIcon(2)
ClickWriteMail:
    ACounter& = Timer
    Call PostMessage(TBarIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(TBarIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        Call GetWriteMail(TheWin&, ToEdit&, SubjectEdit&, Rich&, SendIcon&)
        If (Timer - ACounter&) > 4 Then GoTo ClickWriteMail
    Loop Until (ToEdit& <> 0&) And (SubjectEdit& <> 0&) And (Rich& <> 0&) And (SendIcon& <> 0&)
EnterInformation:
    DoEvents
    If IsWindowEnabled(SendIcon&) = 0 Then Exit Sub
    Call SendMessageByString(ToEdit&, WM_SETTEXT, 0&, ToWho$)
    Call SendMessageByString(SubjectEdit&, WM_SETTEXT, 0&, TheSubject$)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, TheMessage$)
ClickSendMail:
    ACounter& = Timer
    Call PostMessage(SendIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(SendIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        DummyWin& = FindWindowEx(TheWin&, 0&, "_AOL_Edit", vbNullString)
        If (Timer - ACounter&) > 2 Then GoTo ClickSendMail
    Loop Until DummyWin& = 0&
End Sub


Public Sub TimeOut(ByVal TheDuration As Long)
    'i prefer to use the Sleep API sub, but some of
    'you might like this better, so here.
    'Example of use:
    '   Call Timeout (2)
    Dim CurrentCount As Long
    CurrentCount& = Timer
    Do
        DoEvents
    Loop Until ((Timer - CurrentCount&) >= TheDuration&)
End Sub
Public Function HTMLEveryOther(ByVal Text As String, ByVal HTMLType As String) As String
    'this function rotates between the HTML and no HTML
    'Example of use:
    '   MsgBox HTMLEveryOther("http://come.to/cruelair", "u")
    Dim LoopThrough As Long, ACounter As Byte, HTML As String
    Dim TheMid As String
    HTMLEveryOther$ = ""
    HTML$ = ""
   ACounter = 0
    For LoopThrough& = 1 To Len(Text$)
        ACounter = 1 + ACounter
        If ACounter = 5 Then ACounter = 1
        If ACounter = 1 Then HTML$ = "<" & HTMLType$ & ">"
        If ACounter = 2 Then HTML$ = "</" & HTMLType$ & ">"
        If ACounter = 3 Then HTML$ = "<" & HTMLType$ & ">"
        If ACounter = 4 Then HTML$ = "</" & HTMLType$ & ">"
        TheMid$ = Mid$(Text$, LoopThrough&, 1)
        If TheMid$ = " " Then
            HTMLEveryOther$ = HTMLEveryOther$ & " "
            ACounter = ACounter - 1
            GoTo LoopIt
        End If
        HTMLEveryOther$ = HTMLEveryOther$ & HTML$ & TheMid$
LoopIt:
    Next LoopThrough&
End Function

Public Function HTMLFirstLetter(ByVal TheString As String, ByVal HTML_Type As String) As String
    'this function will use HTML on the first letter of each word
    'Example of use:
    '   MsgBox HTMLFirstLetter("please go to crües lair", "b")
    Dim HTML1 As String, HTML2 As String, LoopThrough As Integer, SpaceCheck As Byte
    If Mid(TheString$, 1, 1) = " " Then
        SpaceCheck = 0
    Else
        SpaceCheck = 1
    End If
    HTML1$ = "<" & HTML_Type & ">"
    HTML2$ = "</" & HTML_Type & ">"
    For LoopThrough% = 34 To 255
        If LoopThrough% <> 60 And LoopThrough% <> 62 Then
            TheString$ = Replace(TheString$, " " & Chr(LoopThrough%), " " & HTML1$ & Chr(LoopThrough%) & HTML2$)
        End If
    Next LoopThrough%
    If SpaceCheck = 1 Then
        TheString$ = HTML1$ & Mid(TheString$, 1, 1) & HTML2$ & Mid(TheString$, 2)
    End If
    HTMLFirstLetter$ = TheString$
End Function
Public Function HTMLRotate(ByVal Text As String, ByVal HTML_Type1 As String, ByVal HTML_Type2 As String) As String
    'this function will rotate between two HTML strings
    'Example of use:
    '   MsgBox HTMLRotate("http://come.to/cruelair", "b", "i")
    Dim LoopThrough As Long, ACounter As Byte, HTML As String
    Dim TheMid As String
    HTMLRotate$ = ""
    HTML$ = ""
    ACounter = 0
    For LoopThrough& = 1 To Len(Text$)
        ACounter = 1 + ACounter
        If ACounter = 5 Then ACounter = 1
        If ACounter = 1 Then HTML$ = "<" & HTML_Type1$ & ">"
        If ACounter = 2 Then HTML$ = "</" & HTML_Type1$ & ">"
        If ACounter = 3 Then HTML$ = "<" & HTML_Type2$ & ">"
        If ACounter = 4 Then HTML$ = "</" & HTML_Type2$ & ">"
        TheMid$ = Mid$(Text$, LoopThrough&, 1)
        If TheMid$ = " " Then
            HTMLRotate$ = HTMLRotate$ & " "
            ACounter = ACounter - 1
            GoTo LoopIt
        End If
        HTMLRotate$ = HTMLRotate$ & HTML$ & TheMid$
LoopIt:
    Next LoopThrough&
End Function

Public Function FindSignOnWindow() As Long
    'finds AOL 4.0's sign on window
    'Example of use:
    '   MsgBox FindSignOnWindow&
    Dim MDIClient As Long, AOLChild As Long, AOLIcon As Long
    Dim AOCombo As Long, AOEdit As Long
    MDIClient& = FindWindow("AOL Frame25", vbNullString)
    If MDIClient& = 0& Then
        FindSignOnWindow& = 0&
        Exit Function
    End If
    MDIClient& = FindWindowEx(MDIClient&, 0&, "MDIClient", vbNullString)
    AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
    If AOLChild& = 0& Then
        FindSignOnWindow& = 0&
        Exit Function
    End If
    Do
        AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
        AOCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
        AOCombo& = FindWindowEx(AOLChild&, AOCombo&, "_AOL_Combobox", vbNullString)
        AOEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
        If (IsWindowVisible(AOLIcon&) = 1) And (IsWindowVisible(AOCombo&) = 1) And (IsWindowVisible(AOEdit&) = 1) Then
            FindSignOnWindow& = AOLChild&
            Exit Function
        End If
        AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    Loop Until AOLChild& = 0&
    FindSignOnWindow& = 0&
End Function



Public Sub WindowClose(ByVal TheWindow As Long)
    'this sub will close a window
    'Example of use:
    '   Call WindowClose(FindChatRoom&)
    Call PostMessage(TheWindow&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub WindowDisable(ByVal TheWindow As Long)
    'this sub will disable a window if it is enabled
    'Example of use:
    '   Call WindowDisable(FindChatRoom&)
    Call EnableWindow(TheWindow&, 0&)
End Sub
Public Sub WindowEnable(ByVal TheWindow As Long)
    'this sub will enable a window if it has been disabled
    'Example of use:
    '   Call WindowEnable(FindChatRoom&)
    Call EnableWindow(TheWindow&, 1&)
End Sub
Public Sub WindowHide(ByVal TheWin As Long)
    'this sub will hide a window
    'Example of use:
    '   Call WindowHide (FindSignOnWindow&)
    Call ShowWindow(TheWin&, SW_HIDE)
End Sub


Public Sub WindowMinimize(ByVal TheWindow As Long)
    'this sub will minimize a window
    'Example of use:
    '   Call WindowMinimize(FindChatRoom&)
    Call ShowWindow(TheWindow&, SW_MINIMIZE)
End Sub
Public Sub WindowRestore(ByVal TheWindow As Long)
    'this sub will restore a window
    'Example of use:
    '   Call WindowRestore(FindChatRoom&)
    Call ShowWindow(TheWindow&, SW_RESTORE)
End Sub


Public Sub WindowShow(ByVal TheWin As Long)
    'this sub will show a window
    'Example of use:
    '   Call WindowShow (FindSignOnWindow&)
    Call ShowWindow(TheWin&, SW_SHOW)
End Sub

