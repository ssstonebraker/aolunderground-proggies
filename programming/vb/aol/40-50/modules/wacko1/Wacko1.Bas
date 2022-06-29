Attribute VB_Name = "WacKo"
'--------------------WacKo1.Bas-------------------
'----------------For Aol4 By WacKo----------------
'--------E-Mail me at xwacko64@hotmail.com--------
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
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
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
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
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
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
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
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
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global iniPath
Function TrimTime()
B$ = Left$(Time$, 5)
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(B$, 3) & " " & Ap$
End Function

Function TrimTime2()
B$ = Time$
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(B$, 5) & " " & Ap$
End Function

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Sub AOL4_EliteTalker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then Leet$ = "ë"
    If X = 2 Then Leet$ = "ê"
    If X = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then Leet$ = "ì"
    If X = 2 Then Leet$ = "ï"
    If X = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then Leet$ = "ô"
    If X = 2 Then Leet$ = "ð"
    If X = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then Leet$ = "ù"
    If X = 2 Then Leet$ = "û"
    If X = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then Leet$ = "Å"
    If X = 2 Then Leet$ = "Ä"
    If X = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then Leet$ = "Ï"
    If X = 2 Then Leet$ = "Î"
    If X = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If letter$ = "`" Then Leet$ = "´"
    If letter$ = "!" Then Leet$ = "¡"
    If letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
SendChat (Made$)
End Sub

Function List_CountList(lst As ListBox)
X = lst.ListCount
List_Count = X
End Function


Public Sub Midi_Stop(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub Midi_Play(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub List_Load(Directory As String, TheList As ListBox)
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

Public Sub List_Load2Lists(Directory As String, ListA As ListBox, ListB As ListBox)
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
Public Sub Combo_Load(ByVal Directory As String, Combo As ComboBox)
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

Function Play_Avi()
'Plays a AVI File Change the path Below to your
'AVI Path
lRet = mciSendString("play c:\windows\help\scroll.avi", 0&, 0, 0)
End Function

Sub List_KillDupes(lst As Control)
On Error Resume Next
For i = 0 To lst.ListCount - 1
For e = 0 To lst.ListCount - 1
If LCase(lst.List(i)) Like LCase(lst.List(e)) And i <> e Then
lst.RemoveItem (e)
End If
Next e
Next i
End Sub
Sub Combo_KillDupes(cmb As Control)
For i = 0 To cmb.ListCount - 1
For e = 0 To cmb.ListCount - 1
If LCase(cmb.List(i)) Like LCase(cmb.List(e)) And i <> e Then
cmb.RemoveItem (e)
End If
Next e
Next i
End Sub
Sub List_Clear(lst As ListBox)
lst.Clear
End Sub
Sub Combo_Clear(cmb As ComboBox)
cmb.Clear
End Sub
Sub WAV_Loop(File)

SoundName$ = File
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub WaitForOk()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub


Sub WAV_Play(File)

SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Sub WAV_Stop()

Call WAV_Play(" ")
End Sub
Sub Window_Kill(win)
' This is like killwin
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub
Sub Window_Hide(hwnd)

X = ShowWindow(hwnd, SW_HIDE)
End Sub
Sub Window_Maximize(win)

X = ShowWindow(win, SW_MAXIMIZE)
End Sub

Sub Window_Minimize(win)

X = ShowWindow(win, SW_MINIMIZE)
End Sub


Sub Window_Show(hwnd)

X = ShowWindow(hwnd, SW_SHOW)
End Sub

Sub UpchatOn()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Sub Center(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Public Sub AOL4_Mail_Send(Person As String, subject As String, message As String)
    Dim AOL As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
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
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub


Sub UnUpchatOff()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub
Function ChatSend_Bold(Txt As String)
ChatSend ("<b>" + Txt + "</b>")
End Function
Function ChatSend_Italic(Txt As String)
ChatSend ("<i>" + Txt + "</i>")
End Function
Function ChatSend_Underline(Txt As String)
ChatSend ("<u>" + Txt + "</u>")
End Function
Function AOL4_LastChatLineWithSN()
'duh this will get the text from
'the last chatline with the sn
' used in many bots and shit like that
chattext$ = GetchatText

For FindChar = 1 To Len(chattext$)

thechar$ = Mid(chattext$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(chattext$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function
Function AOL4_LastChatLine()
'duh this will get the text from
'the last chatline , used in many
'bots and shit like that
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function


Sub AOL4_Mail_AddNewToList(ListBox As ListBox)
' This will add your new mail to a listbox.
' It will not open the newmailbox so you will need to
' open it first.
' Call AOL4_Mail_AddNewToListBox(List1)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMD% = FindChildByClass(AOL%, "MDIClient")
themail% = FindChildByClass(AOLMD%, "AOL Child")
themail% = FindChildByClass(themail%, "_AOL_TabControl")
dsa% = FindChildByClass(themail%, "_AOL_TabPage")
thetree% = FindChildByClass(dsa%, "_AOL_Tree")
Do
NumMail% = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
TimeOut 1.5
nummails% = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
Loop Until NumMail% = nummails%
For X = 0 To nummails% - 1
Mails$ = String(256, " ")
Z = SendMessageByString(thetree%, LB_GETTEXT, X, Mails$)
k = Trim$(Mails$)
Where = InStr(Mails$, Chr$(9))
Mails$ = Mid$(Mails$, Where + 1)
Where = InStr(Mails$, Chr$(9))
SN$ = Trim$(Mid$(Mails$, 1, Where - 1))
sns$ = Len(SN$)
sns$ = sns$ + 2
last$ = Mid(Mails$, sns$, Len(Mails$))
ListBox.AddItem (last$)
Next X
End Sub
Public Sub AOL4_Mail_OpenNew()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, sMod As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub


Public Sub AOL4_Mail_AddSentToList(TheList As ListBox)
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
        TheList.AddItem MyString$
    Next AddMails&
End Sub
Public Sub AOL4_Mail_AddOldToList(TheList As ListBox)
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
        TheList.AddItem MyString$
    Next AddMails&
End Sub

Public Sub AOL4_Mail_AddFlashToList(TheList As ListBox)
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
        TheList.AddItem MyString$
    Next AddMails&
End Sub
Public Sub AOL4_Mail_Foward(SendTo As String, message As String, DeleteFwd As Boolean)
    Dim AOL As Long, MDI As Long, Error As Long
    Dim OpenForward As Long, OpenSend As Long, SendButton As Long
    Dim DoIt As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, Rich As Long, fCombo As Long
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
    If DeleteFwd = True Then
        TempSubject$ = GetText(EditSubject&)
        TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, TempSubject$)
        DoEvents
    End If
    Call SendMessageByString(EditTo&, WM_SETTEXT, 0, SendTo$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
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
Public Function AOL4_Mail_CountFlash() As Long
    Dim AOL As Long, MDI As Long, fMail As Long, fList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    fMail& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
    fList& = FindWindowEx(fMail&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MailCountFlash& = Count&
End Function

Public Function AOL4_CountNew() As Long
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
Public Function AOL4_Mail_CountOld() As Long
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
Public Function AOL4_Mail_CountSent() As Long
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
Public Function AOL4_LineCount(MyString As String) As Long
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
Public Function List_GetText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, buffer$)
    GetListText$ = buffer$
End Function


Public Function INI_Read(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub INI_Write(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub


Public Sub AOL4_Mail_OpenSent()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 5
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub



Function AOL4_Bot_Weapon()
' Put This In A Timer And Set Interval To 800
If P = AOL4_LastChatLineWithSN Then GoTo cd
P = AOL4_LastChatLineWithSN
Q = LCase(AOL4_LastChatLine)
r = AOL4_SNFromLastChatLine
  Randomize Timer
  wp% = (Rnd * 10)
   If wp% = 1 Then
    thing$ = "Mashedy"
   ElseIf wp% = 2 Then
    thing$ = "M-16"
   ElseIf wp% = 3 Then
    thing$ = "Hand Gernade"
   ElseIf wp% = 4 Then
    thing$ = "7 Inch Knife"
   ElseIf wp% = 5 Then
   thing$ = "Mother Fucking Tank"
   ElseIf wp% = 6 Then
   thing$ = "A Ruber Band"
   Else
    thing$ = "Rusty Hunk of Pipe"
   End If
  d = InStr(Q, "/weapon")
   If d Then Fade_RedBlue ("·•×® " + r + " you just got a " + thing$ + " for a weapon")
cd:
End Function


Sub FormFade(frm, Style)
'This fade makes a pipe like effect on the form
'Code example:
'FormFade Me, 1


Dim cx, cy, F, F1, F2, i    ' Declare variables
frm.AutoRedraw = True
frm.Cls
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2

DrawWidth = 2   ' Set DrawWidth.
    
For i = 255 To 0 Step -2

F = i / 255  ' Perform interim
F1 = 1 - F: F2 = 1 + F  ' calculations.

If Style = 1 Then frm.ForeColor = RGB(i, i, i) ' Black to white
If Style = 2 Then frm.ForeColor = RGB(0, i, i) ' Black to Cyan
If Style = 3 Then frm.ForeColor = RGB(i, 0, i) ' Black to Purple
If Style = 4 Then frm.ForeColor = RGB(i, i, 0) ' Black to Yellow
If Style = 5 Then frm.ForeColor = RGB(0, 0, i) ' Black to Blue
If Style = 6 Then frm.ForeColor = RGB(i, 0, 0) ' Black to Red
If Style = 7 Then frm.ForeColor = RGB(0, i, 0) ' Black to Green
If Style = 8 Then frm.ForeColor = RGB(0, i, 255) ' Blue to Green
If Style = 9 Then frm.ForeColor = RGB(i, i, 255) ' Blue to White
If Style = 11 Then frm.ForeColor = RGB(i, 0, 255) ' Blue to Purple
If Style = 12 Then frm.ForeColor = RGB(0, 0, 255 - i) ' Blue to Black
If Style = 13 Then frm.ForeColor = RGB(255, 0, i) ' Red to Purple
If Style = 14 Then frm.ForeColor = RGB(255, i, i) ' Red to White
If Style = 15 Then frm.ForeColor = RGB(255, i, 0) ' Red to Yellow
If Style = 16 Then frm.ForeColor = RGB(255 - i, 0, 0) ' Red to Black
If Style = 17 Then frm.ForeColor = RGB(i, 255, i) ' Green to White
If Style = 18 Then frm.ForeColor = RGB(0, 255, i) ' Green to Blue
If Style = 19 Then frm.ForeColor = RGB(i, 255, 0) ' Green to Yellow
If Style = 20 Then frm.ForeColor = RGB(0, 255 - i, 0) ' Green to Black
If Style = 21 Then frm.ForeColor = RGB(255 - i, 255 - i, 255 - i) ' White to Black
If Style = 22 Then frm.ForeColor = RGB(255, 255, 255 - i) 'White to Yellow
If Style = 23 Then frm.ForeColor = RGB(255, 255 - i, 255) 'White to Purple
If Style = 24 Then frm.ForeColor = RGB(255 - i, 255, 255) 'White to Cyan
If Style = 25 Then frm.ForeColor = RGB(255, 255, i) ' Yellow to White
If Style = 26 Then frm.ForeColor = RGB(255, i, 255) ' Purple to White
If Style = 27 Then frm.ForeColor = RGB(i, 255, 255) ' Cyan to White
If Style = 28 Then frm.ForeColor = RGB(255 - i, 255 - i, 0) ' Yellow to Black
If Style = 29 Then frm.ForeColor = RGB(255 - i, 0, 255 - i) ' Purple to Black
If Style = 30 Then frm.ForeColor = RGB(0, 255 - i, 255 - i) ' Cyan to Black
If Style = 31 Then frm.ForeColor = RGB(S1 - i, S2 - i, S3 - i) ' Selected color to black

frm.Line (cx * F1, 0)-(cx * F2, cy * 2), , BF
Next i
End Sub
Sub AOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Sub AOL4_AddRoomToListBox(ListBox As ListBox)
' this will add all the names from the chat room to a listbox
' Call AOL4_AddRoomToListBox(List1)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
ListBox.Clear
Room = AOL4_FindChatRoom()
If Room = 0 Then MsgBox "There must be a chat room open to use this sub", vbInformation, ""
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
ListBox.AddItem Person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub AOL4_AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
' this will add everyone in the chat room to a listbox, then to a combobox
' Call AOL4_AddRoomToComboBox(List1, Combo1)
Call AOL4_AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
Function AOL4_FindChatRoom()
' this will find the chat room and set the focus on the chat room
' Call AOL4_FindChatRoom
' or
' If AOL4_FindChatRoom = 0 then MsgBox "Please Enter A Chat Room Before using this function"
    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    listere% = FindChildByClass(firs%, "RICHCNTL")
    listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (L <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            listers% = FindChildByClass(firs%, "RICHCNTL")
            listere% = FindChildByClass(firs%, "RICHCNTL")
            listerb% = FindChildByClass(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            L = L + 1
    Loop
    If (L < 100) Then
        AOL4_FindChatRoom = firs%
        Exit Function
    End If
    AOL4_FindChatRoom = 0
End Function

Function AOL4_GetRoomName() As String
' this will get the Room Name
' tex$ = AOL40_GetRoomName
' MsgBox tex$
On Error Resume Next
AOL4_GetRoomName = GetApiText(AOL4_FindChatRoom())
End Function
Sub AOL4_Invite(Person)
'This will send an Invite to a person
'werks good for a pinter if u use a timer
FreeProcess
On Error GoTo ErrHandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
bud% = FindChildByTitle(MDI%, "Buddy List Window")
e = FindChildByClass(bud%, "_AOL_Icon")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
ClickIcon (e)
TimeOut (1#)
Chat% = FindChildByTitle(MDI%, "Buddy Chat")
aoledit% = FindChildByClass(Chat%, "_AOL_Edit")
If Chat% Then GoTo FILL
FILL:
Call AOL4_SetText(aoledit%, Person)
de = FindChildByClass(Chat%, "_AOL_Icon")
ClickIcon (de)
Killit% = FindChildByTitle(MDI%, "Invitation From:")
AOL4_KillWin (Killit%)
FreeProcess
ErrHandler:
Exit Sub
End Sub



Sub AOL4_Kill_Win(Windo)
'Closes a window....ex: AOL4Killwin (IM%)
CloseTheMofo = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub

Function AOL4_IM_MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMText% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(IMText%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

Sub AOL4_SetText(win, Txt)
'This is usually used for an _AOL_Edit or RICHCNTL
TheText% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub

Sub AOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_RESTORE)
Call AOL4SetFocus
End Sub


Function AOL4_UpChat()
'this is an upchat that minimizes the
'upload window
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call AOL4SetFocus
End Function

Sub AOL4_Load()
'This will load AOL4.0
X% = Shell("C:\aol40\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub

Function AOL4_Activate()
X = GetCap(AOLWindow)
AppActivate X
End Function
Function AOL4_IM_SNFromIM()

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function
Function AOL4_SNFromLastChatLine()
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function
Function AOL4_UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = user
End Function
Sub AOL4_Virus()
'This was takin outta nash40.bas
'Thanxz Nash
Printer.Print "WacKo Says Good Night"



End Sub
Function Fade_BlackBlueBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackGreenBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackGrey(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackGreyBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackPurpleBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_BlackRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_BlackRedBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlackYellowBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_BlueBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlueBlackBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_BlueGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlueGreenBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BluePurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
   ChatSend ("<b>" + Msg + "")
End Function

Function Fade_BluePurpleBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
  ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlueRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
 ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlueRedBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
   ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlueYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    ChatSend ("<b>" + Msg + "")
End Function
Function Fade_BlueYellowBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
 ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreenBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreenBlackGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreenBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
   ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreenBlueGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    ChatSend ("<b>" + Msg + "")
End Function

Function Fade_GreenPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
  ChatSend ("<b>" + Msg + "")
End Function

Function Fade_GreenPurpleGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreenRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    ChatSend ("<b>" + Msg + "")
End Function

Function Fade_GreenRedGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
 ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreenYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreenYellowGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & d
    Next B
    ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 220 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_GreyBlackGrey(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyBlueGrey(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_GreyGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyGreenGrey(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b> " + Msg + "")
End Function
Function Fade_GreyPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyPurpleGrey(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_GreyRedGrey(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_GreyYellowGrey(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 490 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_PurpleBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_PurpleBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_PurpleBluePurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_PurpleGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_PurpleGreenPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleGreenPurple = ("<b>" + Msg + "")
End Function
Function Fade_PurpleRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_PurpleRedPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
    PurpleRedPurple = ("<b>" + Msg + "")
End Function
Function Fade_PurpleYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_PurpleYellowPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_RedBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_RedBlackRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
 ChatSend ("<b>" + Msg + "")
End Function
Function Fade_RedBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
 ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowBlack(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowBlackYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowBlueYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowGreenYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowPurpleYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_YellowRedYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_RedBlueRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_RedGreen(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<b><Font Color=#" & h & "></b>" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_RedGreenRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Function Fade_RedPurple(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_RedYellow(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Function Fade_RedYellowRed(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function

Sub AOL4_Kill_Wait()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_Kill_Modal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub
Sub AOL4_Kill_Glyph()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub

Sub AOL4_Keyword(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub
Sub AOL4_IM_Send(SN, message)

Call AOL40_Keyword("aol://9293:" & SN)
Do: DoEvents
IMWin% = FindChildByTitle(AOLMDI(), "Send Instant Message")
Rich% = FindChildByClass(IMWin%, "RICHCNTL")
Icon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until Rich% <> 0 And Icon% <> 0
Call SendMessageByString(Rich%, WM_SETTEXT, 0, message)
For X = 1 To 9
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Next X
Call Pause(0.01)
AOLClickIcon (Icon%)
Do: DoEvents
IMWin% = FindChildByTitle(AOLMDI(), "Send Instant Message")
OK% = FindWindow("#32770", "America Online")
If OK% <> 0 Then Call SendMessage(OK%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Sub AOL4_IM_IgnoreSN()
' AOL40_IM_Send ($im_off & sn, "WacKo")
End Sub
Sub AOL4_IM_UnIgnoreSN()
' Call AOL4_IM_Send ($im_off & sn, "WacKo")
End Sub
Function Fade_BlackBlue(Text1)
a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 255 / a
        F = e * B
        G = RGB(F, 0, 0)
        h = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & h & ">" & d
    Next B
ChatSend ("<b>" + Msg + "")
End Function
Sub AOL4_BustInPR(URL As String)
' This will bust into a private room that you specify
busttin:
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
Call SendMessage(Room%, WM_CLOSE, 0, 0)
End If
tryagain:
Call AOL4_Keyword("aol://2719:2-2-" & URL$)
Call AOL4_FindChatRoom
If AOL4_FindChatRoom = 0 Then GoTo noroom:
GoTo FoundRoom:
noroom:
noroomagaiN:
full% = FindWindow("#32770", "America Online")
If AOL40_FindChatRoom <> 0 Then GoTo FoundRoom:
If full% = 0 Then GoTo noroomagaiN:
CloseWin (full%)
GoTo tryagain:
FoundRoom:
ChatSend "    •† RaVaGe Busta¹ Room Entered: " & AOL40_GetRoomName
End Sub


Sub AOL4_AntiIdle()
' this will make u stay online
' Call AOL4_AntiIdle
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AOL4_Anti45MinTimer()
' this will click on the window that pops up saying how long u have been online
' Call AOL4_Anti45MinTimer
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Public Sub AOL4_Button(but%)
Clicicon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
Clicicon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function AOL4_FindRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOL4FindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function

Function AOL4_hwnd() As Integer
'This sets focus on the AOL window
AOL = FindWindow("AOL Frame25", vbNullString)
End Function

Sub AOL4_Icon(Icon%)
Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Function AOL4_IM_Static(newcaption As String)
ANTI1% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
STS% = FindChildByClass(ANTI1%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call ChangeCap(st%, newcaption)
End Function


Sub ChangeCaption(HWD%, newcaption As String)
Call AOL4_SetText(HWD%, newcaption)
End Sub



Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub AOL4_RunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub


Sub ChatSend(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub AOL4_Linksend()
'Sends Link To Chat
ChatSend "<b>·•×® < a href=" + link + ">" + Txt + " ®×•·</a>"
End Sub
Public Sub AOL4_Mail_OpenOld()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    For DoThis& = 1 To 4
        Call PostMessage(sMod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(sMod&, WM_KEYUP, VK_DOWN, 0&)
    Next DoThis&
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub
Public Sub AOL4_Mail_OpenFalsh()
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim CurPos As POINTAPI, WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    Call GetCursorPos(CurPos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do
        sMod& = FindWindow("#32768", vbNullString)
        WinVis& = IsWindowVisible(sMod&)
    Loop Until WinVis& = 1
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RIGHT, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub


Public Sub Combo_Save(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub
Public Sub List_Save(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, GetString)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub



