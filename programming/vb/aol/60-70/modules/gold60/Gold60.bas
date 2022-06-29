Attribute VB_Name = "Gold60"
'First off let me say that I stopped programming about
'a year ago and just out of curiousity I checked knk's site
'to see what all programmers have done for AOL 6.0 and it
'made me sick that there wasn't a single IM code in the 3
'bas files there nor mail that didn't use pauses. Pauses
'are so elementry. This bas was NOT written entirely by
'me. This is a combo of many bas files and my subs with it.
'I gave credit to the subs/functions others wrote.
'If I felt that it was written as good as it could be
'then I didn't write one. This bas took me minutes to
'make and isn't thought out or well done, I used monkspy
'for about 10 minutes to make working aol 6 functions so
'atleast people have something half decent to program with.
'This is just the basics not alot of extras and something
'I whipped together in seconds or boredom. Shout outs to
'JMR wherever you are.
'------------ Gold --------------




'decs from dos32.bas added on by Gold
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
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
Public Const SW_RESTORE = 9
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


Sub Ghost_On()
'written by progee
a = FindWindow("aol frame25", vbNullString)
b = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(b, 0, "AOL Child", "Buddy List")
D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
For E = 1 To 6
D = GetWindow(D, 2)
Next E
Icon D
Do
f = FindWindowEx(b, 0, "AOL Child", "Buddy List Setup")
G = FindWindowEx(f, 0, "_AOL_Edit", vbNullString)
DoEvents
Loop Until G > 0
G = GetWindow(G, 3)
G = GetWindow(G, 3)
Icon G
Do
H = FindWindowEx(b, 0, "AOL Child", "Buddy List Preferences")
i = FindWindowEx(H, 0, "_AOL_TabControl", vbNullString)
j = FindWindowEx(i, 0, "_AOL_TabPage", vbNullString)
j2 = GetWindow(j, 2)
j3 = GetWindow(j2, 2)
k = FindWindowEx(j3, 0, "_AOL_RadioBox", vbNullString)
DoEvents
Loop Until k > 0
For L = 1 To 4
k = GetWindow(k, 2)
Next L
Icon k
For L2 = 1 To 3
k = GetWindow(k, 2)
Next L2
Icon k
M = FindWindowEx(H, 0, "_AOL_Icon", vbNullString)
Icon M
Pause 1
PostMessage f, WM_CLOSE, 0, 0
End Sub
Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub



Sub Ghost_Off()
'written by progee
a = FindWindow("aol frame25", vbNullString)
b = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(b, 0, "AOL Child", "Buddy List")
D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
For E = 1 To 6
D = GetWindow(D, 2)
Next E
Icon D
Do
f = FindWindowEx(b, 0, "AOL Child", "Buddy List Setup")
G = FindWindowEx(f, 0, "_AOL_Edit", vbNullString)
DoEvents
Loop Until G > 0
G = GetWindow(G, 3)
G = GetWindow(G, 3)
Icon G
Do
H = FindWindowEx(b, 0, "AOL Child", "Buddy List Preferences")
i = FindWindowEx(H, 0, "_AOL_TabControl", vbNullString)
j = FindWindowEx(i, 0, "_AOL_TabPage", vbNullString)
j2 = GetWindow(j, 2)
j3 = GetWindow(j2, 2)
k = FindWindowEx(j3, 0, "_AOL_RadioBox", vbNullString)
DoEvents
Loop Until k > 0
For L = 1 To 2
k = GetWindow(k, 2)
Next L
Icon k
For L2 = 1 To 5
k = GetWindow(k, 2)
Next L2
Icon k
M = FindWindowEx(H, 0, "_AOL_Icon", vbNullString)
Icon M
Pause 1
PostMessage f, WM_CLOSE, 0, 0
End Sub
Public Sub Loadlistbox(Path As String, List1 As ListBox)
'Loads list box
'Call LoadListBox(app.path, list1)
        Dim s60 As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, s60$
DoEvents
List1.AddItem s60$
Wend
Close #1
End Sub
Public Sub Keyword(KW As String)
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Sub IMsOff()
im "$im_off", " "
waitforok
a = FindWindow("aol frame25", vbNullString)
b = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(b, 0, "AOL Child", "Send Instant Message")
CloseWindow (c)
End Sub
Function SignedOn() As Boolean
If GetUser = "" Then
SignedOn = False
Else
SignedOn = True
End If
End Function
Sub WinForceShutDwn()
    Call ExitWindowsEx(EWX_FORCE, 0)
End Sub
Sub WinLogOff()
    Call ExitWindowsEx(EWX_LOGOFF, 0)
End Sub
Sub MsgSureExit(ProgName As String, TheFrm As Form)
    Select Case MsgBox("Are you sure you want to exit " & ProgName & "?", vbYesNo + vbQuestion + vbDefaultButton2, ProgName$ & " [Exit?]")
    Case vbYes
        Unload TheFrm
        End
        End
        Unload TheFrm
    Case vbNo
        Exit Sub
    End Select
End Sub

Function ListToString(TheList As ListBox) As String
    Dim DoList As Long, MailString As String
    If TheList.List(0) = "" Then Exit Function
    For DoList& = 0 To TheList.ListCount - 1
        MailString$ = MailString$ & TheList.List(DoList&) & ", "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToString$ = MailString$
End Function
Sub HideAOL()
ShowWindow AOL, SW_HIDE
End Sub
Sub ShowAOL()
ShowWindow AOL, SW_SHOW
End Sub

Sub ShowWelcome()
    AL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AL&, 0&, "MDIClient", vbNullString)
    X = FindWindowEx(MDI&, 0&, "AOL Child", "Welcome, " & GetUser & "!")
    Call ShowWindow(X, SW_SHOW)
End Sub
Sub HideWelcome()
    AL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AL&, 0&, "MDIClient", vbNullString)
    X = FindWindowEx(MDI&, 0&, "AOL Child", "Welcome, " & GetUser & "!")
    Call ShowWindow(X, SW_HIDE)
End Sub
Sub SetText(Window, What)
Call SendMessageByString(Window, WM_SETTEXT, 0&, What)
End Sub
Sub IMsOn()
im "$im_on", " "
waitforok
a = FindWindow("aol frame25", vbNullString)
b = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(b, 0, "AOL Child", "Send Instant Message")
CloseWindow (c)
End Sub

Sub HideWindow(Window)
werg = ShowWindow(Window, SW_HIDE)
End Sub


Sub im(who, What)
'only AOL 6 IM code without gay pauses =)
a = FindWindow("aol frame25", vbNullString)
b = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(b, 0, "AOL Child", "Buddy List")
If c = 0 Then Keyword "aol://9293:"
D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
Icon D
Do
eq = FindWindowEx(b, 0, "AOL Child", "Send Instant Message")
f = FindWindowEx(eq, 0, "_AOL_Edit", vbNullString)
G = FindWindowEx(eq, 0, "RICHCNTL", vbNullString)
OurParent& = eq
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Hand5& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Hand6& = FindWindowEx(OurParent&, Hand5&, "_AOL_Icon", vbNullString)
Hand7& = FindWindowEx(OurParent&, Hand6&, "_AOL_Icon", vbNullString)
Hand8& = FindWindowEx(OurParent&, Hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(OurParent&, Hand8&, "_AOL_Icon", vbNullString)
hq2& = FindWindowEx(OurParent&, Hand9&, "_AOL_Icon", vbNullString)
Loop Until eq <> 0 And f <> 0 And G <> 0 And hq2 <> 0
SetText f, who
SetText G, What
Icon hq2
End Sub


Sub Click(Icon)
    Call PostMessage(Icon, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(Icon, WM_LBUTTONUP, 0&, 0&)
End Sub

Sub KillGlyph()
'Kills that stupid spinning,glowing,blue AOL picture
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "AOL Toolbar", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "_AOL_Toolbar", vbNullString)
OurHandle& = FindWindowEx(OurParent&, 0, "_AOL_Glyph", vbNullString)
CloseWindow (OurHandle&)
End Sub


Sub KillModals()
Modal& = FindWindow("_AOL_Modal", vbNullString)
CloseWindow Modal&
End Sub
Sub HiddenIM(who, What)
a = FindWindow("aol frame25", vbNullString)
b = FindWindowEx(a, 0, "mdiclient", vbNullString)
c = FindWindowEx(b, 0, "AOL Child", "Buddy List")
If c = 0 Then Keyword "aol://9293:"
D = FindWindowEx(c, 0, "_AOL_Icon", vbNullString)
Icon D
Do
eq = FindWindowEx(b, 0, "AOL Child", "Send Instant Message")
f = FindWindowEx(eq, 0, "_AOL_Edit", vbNullString)
G = FindWindowEx(eq, 0, "RICHCNTL", vbNullString)
OurParent& = eq
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Hand5& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Hand6& = FindWindowEx(OurParent&, Hand5&, "_AOL_Icon", vbNullString)
Hand7& = FindWindowEx(OurParent&, Hand6&, "_AOL_Icon", vbNullString)
Hand8& = FindWindowEx(OurParent&, Hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(OurParent&, Hand8&, "_AOL_Icon", vbNullString)
hq2& = FindWindowEx(OurParent&, Hand9&, "_AOL_Icon", vbNullString)
Loop Until eq <> 0 And f <> 0 And G <> 0 And hq2 <> 0
HideWindow (G)
SetText f, who
SetText G, What
Icon hq2
End Sub
Public Function FindRoom() As Long
'by dos
    Dim AOL As Long, MDI As Long, Child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindRoom& = Child&
End Function
Public Function GetText(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, Buffer$)
    GetText$ = Buffer$
End Function
Public Sub ChatSend(strmessage As String)
'written by slove
On Error Resume Next
Dim lngchat As Long, lngrich As Long, strtext As String, WM_CLEAR As Data
    
    Let lngchat& = FindRoom&
    Let lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
    Let lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
    Let strtext$ = GetText(lngrich&)
    If strtext$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strmessage$)
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until GetText$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strtext$)
    End Sub
    
    Public Sub SendChat(strmessage As String)
On Error Resume Next
Dim lngchat As Long, lngrich As Long, strtext As String, WM_CLEAR As Data
    
    Let lngchat& = FindRoom&
    Let lngrich& = FindWindowEx(lngchat&, 0&, "richcntl", vbNullString)
    Let lngrich& = FindWindowEx(lngchat&, lngrich&, "richcntl", vbNullString)
    Let strtext$ = GetText(lngrich&)
    If strtext$ <> "" Then
        Call SendMessageLong(lngrich&, WM_CLEAR, 0&, 0&)
        Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "")
    End If
    X$ = FadeByColor3(&HFFFF00, &HFF0000, &H800000, strmessage$, False)
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, "<b></u></i><font face=""arial narrow""><font size = 10><Font Color=#00EDFF>•–[ " & X$ & " ]–•")
    Call SendMessageLong(lngrich&, WM_CHAR, ENTER_KEY, 0&)
    Do: DoEvents: Loop Until GetText$(lngrich&) = ""
    Call SendMessageByString(lngrich&, WM_SETTEXT, 0&, strtext$)
    End Sub

Public Function Icon(Ico)
'source60.bas
Call SendMessageLong(Ico, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Ico, WM_KEYUP, VK_SPACE, 0&)
End Function
Function ClickList(hwnd, index)
cliX0r& = SendMessage(hwnd, LB_SETCURSEL, ByVal CLng(index), ByVal 0&)
End Function

Public Sub CloseWindow(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub


Public Sub FormOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
 
Sub AcidTrip(Frm As Form)
' from Source60.bas
Dim cx, cy, Radius, Limit
    Frm.ScaleMode = 3
    cx = Frm.ScaleWidth / 2
    cy = Frm.ScaleHeight / 2
    If cx > cy Then Limit = cy Else Limit = cx
    For Radius = 0 To Limit
Frm.Circle (cx, cy), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   Pause 0.03
    Next Radius
End Sub
Sub OpenExe(Path)
Shell (Path)
End Sub
Sub mail(who, Subj, Mess)
'written using monkspy but works without
'gay pauses like the other aol6 bas files
Do
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "AOL Toolbar", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "_AOL_Toolbar", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Loop Until OurHandle& <> 0 And OurParent& <> 0
Icon OurHandle&
Do
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Write Mail")
whoto& = FindWindowEx(OurParent&, 0, "_AOL_Edit", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Edit", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Edit", vbNullString)
Subh& = FindWindowEx(OurParent&, Hand2&, "_AOL_Edit", vbNullString)
Messj& = FindWindowEx(OurParent&, 0, "RICHCNTL", vbNullString)
Loop Until whoto& <> 0 And Subh& <> 0 And Messj& <> 0
SetText whoto&, who
SetText Subh&, Subj
SetText Messj&, Mess
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Hand5& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Hand6& = FindWindowEx(OurParent&, Hand5&, "_AOL_Icon", vbNullString)
Hand7& = FindWindowEx(OurParent&, Hand6&, "_AOL_Icon", vbNullString)
Hand8& = FindWindowEx(OurParent&, Hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(OurParent&, Hand8&, "_AOL_Icon", vbNullString)
Hand10& = FindWindowEx(OurParent&, Hand9&, "_AOL_Icon", vbNullString)
Hand11& = FindWindowEx(OurParent&, Hand10&, "_AOL_Icon", vbNullString)
Hand12& = FindWindowEx(OurParent&, Hand11&, "_AOL_Icon", vbNullString)
Hand13& = FindWindowEx(OurParent&, Hand12&, "_AOL_Icon", vbNullString)
Hand14& = FindWindowEx(OurParent&, Hand13&, "_AOL_Icon", vbNullString)
Hand15& = FindWindowEx(OurParent&, Hand14&, "_AOL_Icon", vbNullString)
Hand16& = FindWindowEx(OurParent&, Hand15&, "_AOL_Icon", vbNullString)
Hand17& = FindWindowEx(OurParent&, Hand16&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand17&, "_AOL_Icon", vbNullString)
Icon (OurHandle&)
End Sub
Sub WinShowStartButton()
Shellwin& = FindWindow("Shell_TrayWnd", "")
Startbttn& = FindWindowEx(Shellwin&, 0, "Button", vbNullString)
ShowWindow Startbttn&, SW_SHOW
End Sub
Sub WinRestart()
    Call ExitWindowsEx(EWX_REBOOT, 0)
End Sub
Sub WinHideStartButton()
Shellwin& = FindWindow("Shell_TrayWnd", "")
Startbttn& = FindWindowEx(Shellwin&, 0, "Button", vbNullString)
HideWindow (Startbttn&)
End Sub

Sub DisableCAD()
'disables Alt+Ctrl+Del
Call SystemParametersInfo(97, True, 0&, 0)
End Sub
Public Sub drag(Frm As Form)
Call ReleaseCapture
    Call SendMessage(Frm.hwnd, WM_SYSCOMMAND, WM_MOVE, vbNullString)
End Sub

Sub EnableCAD()
'enables Alt+Ctrl+Del
Call SystemParametersInfo(97, False, 0&, 0)
End Sub
Public Function SetPW(Text As String)
'on the sign on screen
'ex: Call SetPw("PwWouldGoHere")
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Text$)
Call SendMessageLong(AOLEdit&, WM_CHAR, ENTER_KEY, 0&)
End Function
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, WavY)

End Function
Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    fstlen% = (Int(TextLen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, TextLen% - fstlen%)
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function
Function RGBtoHEX(RGB)
'heh, I didnt make this one...
    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function
Public Sub FormExitDown(TheForm As Form)
    On Error Resume Next
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) + 300))
    Loop Until TheForm.Top > 7200
End Sub

Public Function GetCaption(WindowHandle As Long) As String
'written by slove
Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Public Function AOL6_GetText(WindowHandle As Long) As String
 'written by slove
 'works on win2k
   Dim Buffer As String, TextLength As Long, txtlen As Long, lineNum As Long
    TextLength& = SendMessage(WindowHandle, EM_GETLINECOUNT, 0&, 0&)
    lineNum = TextLength&
   txtlen = sendmessagebynum(WindowHandle, EM_LINELENGTH, lineNum, 0&)
   Buffer = String(txtlen, 0&)
   Call SendMessage(WindowHandle, EM_GETLINE, lineNum, ByVal Buffer$)
    AOL6_GetText$ = Buffer$
End Function

Sub FormTrippyDance(s60 As Form)
'This makes a form move across the screen
'Example
'Call FormDance(form1)
s60.Left = 5
Pause (0.1)
s60.Left = 400
Pause (0.1)
s60.Left = 700
Pause (0.1)
s60.Left = 1000
Pause (0.1)
s60.Left = 2000
Pause (0.1)
s60.Left = 3000
Pause (0.1)
s60.Left = 4000
Pause (0.1)
s60.Left = 5000
Pause (0.1)
s60.Left = 4000
Pause (0.1)
s60.Left = 3000
Pause (0.1)
s60.Left = 2000
Pause (0.1)
s60.Left = 1000
Pause (0.1)
s60.Left = 700
Pause (0.1)
s60.Left = 400
Pause (0.1)
s60.Left = 5
Pause (0.1)
s60.Left = 400
Pause (0.1)
s60.Left = 700
Pause (0.1)
s60.Left = 1000
Pause (0.1)
s60.Left = 2000
End Sub



Sub ClearChat()

OurParent& = FindRoom&
OurHandle& = FindWindowEx(OurParent&, 0, "RICHCNTL", vbNullString)
SetText OurHandle&, ""
End Sub
Public Function aim_Chat_Clear()
'clears chatroom text on your computer
'ex: Call aim_chat_clear
Dim AteClass2 As Long
Dim WndAteClass As Long
Dim aimchatwnd As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
WndAteClass& = FindWindowEx(aimchatwnd&, 0&, "WndAte32Class", vbNullString)
AteClass2& = FindWindowEx(WndAteClass&, 0&, "Ate32Class", vbNullString)
Call SendMessageByString(WndAteClass&, WM_SETTEXT, 0&, "")
End Function
Sub CloseDecline()
'aim
OurParent& = FindWindowEx(ParHand1&, 0, "#32770", vbNullString)
OurHandle& = FindWindowEx(OurParent&, 0, "Button", vbNullString)
If OurHandle& <> 0 Then
Call SendMessage(OurHandle&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OurHandle&, WM_KEYUP, VK_SPACE, 0&)
    Call SendMessage(OurHandle&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OurHandle&, WM_KEYUP, VK_SPACE, 0&)
IMWin& = FindWindow("aim_imessage", vbNullString)
CloseWindow (IMWin&)
Pause 45
End If
End Sub
Sub ClickButton(Button As Long)
Ret& = SendMessage(Button&, &H100, &H20, 0&)
DoEvents
Ret& = SendMessage(Button&, &H100, &H20, 0&)
End Sub
Function AIMim(Person$, Message$, CloseIM As Boolean)
Call Icon(IMButton)
Dim AIMIMessage&
Do
DoEvents
IMWin& = FindWindow("aim_imessage", vbNullString)
Loop Until IMWin <> 0
Dim oscarpersistantcombo&
Dim edit&
AIMIMessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(AIMIMessage&, 0&, "_oscar_persistantcombo", vbNullString)
edit& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
Call SetText(edit&, Person$)
WndAteClass& = FindWindowEx(AIMIMessage&, 0&, "wndate32class", vbNullString)
AteClass& = FindWindowEx(WndAteClass&, 0&, "ate32class", vbNullString)
WndAteClass& = FindWindowEx(AIMIMessage&, WndAteClass&, "wndate32class", vbNullString)
Call SetText(WndAteClass&, Message$)
oscariconbtn& = FindWindowEx(AIMIMessage&, 0&, "_oscar_iconbtn", vbNullString)
Call Icon(oscariconbtn&)
If CloseIM = True Then
OurHandle& = FindWindow("AIM_IMessage", "" & LCase(Person$) & " - Instant Message")
CloseWindow (OurHandle&)
End If
End Function
Function IMButton()
'for aim
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
IMButton = oscariconbtn&
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
Public Function GetUser() As String
'written by slove
Dim AOL As Long, MDI As Long, welcome As Long
    Dim Child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(Child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(Child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    GetUser$ = ""
End Function

Public Function GetListIndex(oListBox As ListBox, sText As String) As Integer
Dim iIndex As Integer
With oListBox
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With
GetListIndex = -2
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
Sub KillAd()
'aim
Dim oscarbuddylistwin&
Dim WndAteClass&
Dim AteClass&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
WndAteClass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
AteClass& = FindWindowEx(WndAteClass&, 0&, "ate32class", vbNullString)

Call ShowWindow(AteClass&, SW_HIDE)
End Sub

Sub Awaymsgdefault()
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Buddy List")
If OurParent& = 0 Then Keyword "bv"
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurParent& <> 0 And OurHandle& <> 0
Icon OurHandle&
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Away Message")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurParent& <> 0 And OurHandle& <> 0
Icon OurHandle&
End Sub
Sub Awaymsgnew(title, Message)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Buddy List")
If OurParent& = 0 Then Keyword "bv"
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurParent& <> 0 And OurHandle& <> 0
Icon OurHandle&
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Away Message")
OurHandle& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Loop Until OurParent& <> 0 And OurHandle& <> 0
Icon OurHandle&
Do
DoEvents
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "New Away Message")
OurHandle& = FindWindowEx(OurParent&, 0, "_AOL_Edit", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "RICHCNTL", vbNullString)
OurHandl& = FindWindowEx(OurParent&, Hand1&, "RICHCNTL", vbNullString)
Loop Until OurHandle& <> 0 And OurParent& <> 0 And Hand1& <> 0 And OurHandl& <> 0
SetText OurHandle&, title
SetText OurHandl&, Message
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Hand5& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Hand6& = FindWindowEx(OurParent&, Hand5&, "_AOL_Icon", vbNullString)
Hand7& = FindWindowEx(OurParent&, Hand6&, "_AOL_Icon", vbNullString)
Hand8& = FindWindowEx(OurParent&, Hand7&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand8&, "_AOL_Icon", vbNullString)
Icon OurHandle&
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Away Message")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurParent& <> 0 And OurHandle& <> 0
Icon OurHandle&
End Sub
Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    Dim TextLen As String
    Dim faded As String
    WaveState = 0
    
    TextLen$ = Len(TheText)
    For i = 1 To TextLen$
        TextDone$ = Left(TheText, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen$ * i) + B1, ((G2 - G1) / TextLen$ * i) + G1, ((R2 - R1) / TextLen$ * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        faded$ = faded$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    FadeTwoColor = faded$
End Function
Function FadeByColor2(Colr1, Colr2, TheText$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, WavY)

End Function
Sub IMRespond(SN As String, Message)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "RICHCNTL", vbNullString)
ourhand& = FindWindowEx(OurParent&, Hand1&, "RICHCNTL", vbNullString)
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Hand5& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Hand6& = FindWindowEx(OurParent&, Hand5&, "_AOL_Icon", vbNullString)
Hand7& = FindWindowEx(OurParent&, Hand6&, "_AOL_Icon", vbNullString)
Hand8& = FindWindowEx(OurParent&, Hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(OurParent&, Hand8&, "_AOL_Icon", vbNullString)
ourhandle1& = FindWindowEx(OurParent&, Hand9&, "_AOL_Icon", vbNullString)
Call SendMessageLong(OurHandle&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(OurHandle&, WM_KEYUP, VK_SPACE, 0&)
Loop Until Hand1& <> 0 And ourhandle1& <> 0 And ourhand& <> 0
SetText ourhand&, Message
Call SendMessageLong(ourhandle1&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(ourhandle1&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Function AOLGetList(LBHandle, index, Buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
AOLThread = GetWindowThreadProcessId(LBHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or &HF0000, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(LBHandle, &H199, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 28
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)


Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function
Sub addroom(Lst As ListBox)
'heh this killed me that nobody could write this
'people told me "aol subclassed the listbox so you can't
'make a working addroom... :::shakes head:::

Room& = FindRoom
ListX& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
Count& = sendmessagebynum(ListX&, &H18B, 0, 0)
Buffer$ = Space(255)
For Counter% = 0 To Count& - 1
List& = AOLGetList(ListX&, Counter%, Buffer$)
Lst.AddItem LCase((Buffer$))
Next Counter%
For E = 0 To Lst.ListCount
    If Lst.List(E) = LCase(GetUser$) Then
Lst.RemoveItem E
End If
Next E
End Sub

Public Function FindIM() As Long
    Dim AOL As Long, MDI As Long, Child As Long, Caption As String
    ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
OurH& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, 0, "_AOL_Static", vbNullString)
c$ = GetText(OurHandle&)
If InStr(c$, "The Internet user") Then
Click (OurH)
End If
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(Child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        FindIM& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(Child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                FindIM& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindIM& = Child&
End Function

Public Sub waitforok()
     Do
        AOL2& = FindWindow("#32770", "America Online")
        Bottun& = FindWindowEx(AOL2&, 0, "Button", vbNullString)
        Loop Until AOL2& <> 0 And Bottun& <> 0
    Call SendMessage(Bottun&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Bottun&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub AOL5_GhostOn()
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Buddy List Window")
If OurParent& = 0 Then Keyword "bv"
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Loop Until OurParent& <> 0 And OurHandle& <> 0
Click (OurHandle&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", GetUser$ & "'s Buddy List")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Loop Until OurHandle& <> 0
Click (OurHandle&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Privacy Preferences")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Checkbox", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Checkbox", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Checkbox", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Checkbox", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand4&, "_AOL_Checkbox", vbNullString)
Loop Until OurHandle& <> 0
Click (OurHandle&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Privacy Preferences")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurHandle& <> 0
Click (OurHandle&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Privacy Preferences")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurHandle& <> 0
  Call SendMessage(OurHandle&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OurHandle&, WM_KEYUP, VK_SPACE, 0&)
    Our& = FindWindowEx(ParHand2&, 0, "AOL Child", GetUser$ & "'s Buddy List")
    CloseWindow (Our&)
     Do
        AOL2& = FindWindow("#32770", "America Online")
        Bottun& = FindWindowEx(AOL2&, 0, "Button", vbNullString)
        Loop Until AOL2& <> 0 And Bottun& <> 0
    Call SendMessage(Bottun&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Bottun&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub AOL5_GhostOff()
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Buddy List Window")
If OurParent& = 0 Then Keyword "bv"
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Loop Until OurParent& <> 0 And OurHandle& <> 0
Click (OurHandle&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", GetUser$ & "'s Buddy List")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Loop Until OurHandle& <> 0
Click (OurHandle&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Privacy Preferences")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Checkbox", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Checkbox", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Checkbox", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Checkbox", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand4&, "_AOL_Checkbox", vbNullString)
Loop Until OurHandle& <> 0
Pause 0.5
Click (Hand3&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Privacy Preferences")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurHandle& <> 0
Click (OurHandle&)
Do
DoEvents
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Privacy Preferences")
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Loop Until OurHandle& <> 0
    Our& = FindWindowEx(ParHand2&, 0, "AOL Child", GetUser$ & "'s Buddy List")
    CloseWindow (Our&)
    Do
        AOL2& = FindWindow("#32770", "America Online")
        Bottun& = FindWindowEx(AOL2&, 0, "Button", vbNullString)
        Loop Until AOL2& <> 0 And Bottun& <> 0
    Call SendMessage(Bottun&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Bottun&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub KillChannels()
ParHand1& = FindWindow("AOL Frame25", "America  Online")
OurParent& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "AOL Child", vbNullString)
OurHandle& = FindWindowEx(OurParent&, Hand1&, "AOL Child", "AOL Channels")
CloseWindow (OurHandle&)
End Sub
Sub UpchatOn()
Modal& = FindWindow("_AOL_Modal", vbNullString)
If InStr(1, GetCaption(Modal&), "File Transfer") <> 0 Then
DoEvents
Ret& = ShowWindow(Modal&, 0)
Ret& = SetFocusAPI(AOL)
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Write Mail")
ShowWindow Modal&, SW_MINIMIZE
ShowWindow OurParent&, SW_MINIMIZE
End If
End Sub
Sub UpChatOff()
Modal& = FindWindow("_AOL_Modal", vbNullString)
If InStr(1, GetCaption(Modal&), "File Transfer") <> 0 Then
DoEvents
Ret& = ShowWindow(Modal&, 1)
Ret& = SetFocusAPI(Modal&)
ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", "Write Mail")
ShowWindow Modal&, SW_RESTORE
ShowWindow OurParent&, SW_RESTORE
End If
End Sub

Public Sub UnloadAllForms()
    Dim OfTheseForms As Form
For Each OfTheseForms In Forms
Unload OfTheseForms
Set OfTheseForms = Nothing
Next OfTheseForms
End Sub

Public Function IMLastMsg() As String
If FindIM& = 0 Then Exit Function
IMText& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
IMLastMsg = GetText(IMText&)
End Function
Public Function IMSender() As String
    Dim im As Long, Caption As String
    If FindIM& = 0 Then Exit Function
    Caption$ = GetCaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function


Public Sub RunMenu(TopMenu As Long, SubMenu As Long)
    Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub
Sub killwait()
RunMenu 4, 10
Do
DoEvents
Modal& = FindWindow("_AOL_Modal", vbNullString)
k& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
Loop Until Modal& And k& <> 0
Click (k&)
End Sub
Public Sub ChatIgnoreByIndex(index As Long)
'dos
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
            If Version = 5 Then
            Call SendMessage(dButton&, WM_LBUTTONDOWN, 0&, 0&)
            Call SendMessage(dButton&, WM_LBUTTONUP, 0&, 0&)
            End If
            If Version = 6 Then
            Icon (dButton&)
            End If
            DoEvents
            SearchBox& = SearchBox& - 1
        End If
    Next SearchBox&
End Sub
Public Function MailSenderNew(index As Long) As String
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot1 As Long, Spot2 As Long, MyString As String
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
    Spot1& = InStr(MyString$, Chr(9))
    Spot2& = InStr(Spot1& + 1, MyString$, Chr(9))
    MyString$ = Mid(MyString$, Spot1& + 1, Spot2& - Spot1& - 1)
    MailSenderNew$ = MyString$
End Function
