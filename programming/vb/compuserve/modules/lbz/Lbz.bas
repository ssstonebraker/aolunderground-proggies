Attribute VB_Name = "Lbz"
'Thirty Three 1/3
'For Compuserve 2000 Ver. 5.o for Win 95 / 98
'Written by Lbz

Option Explicit
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
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
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

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

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        y As Long
End Type



Public Sub Mail_Send(who As String, Subject As String, message As String)
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
Toolbar& = FindWindowEx(cs&, 0, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
icon1& = FindWindowEx(toolbar2&, 0, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(toolbar2&, icon1&, "_AOL_Icon", vbNullString)
newmail& = FindWindowEx(toolbar2&, icon2&, "_AOL_Icon", vbNullString)
Click newmail&
Do
DoEvents
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
Messy& = FindWindowEx(aochild&, 0, "CS_RICHCNTL", vbNullString)
ToWho& = FindWindowEx(aochild&, 0, "_AOL_Edit", vbNullString)
CC& = FindWindowEx(aochild&, ToWho&, "_AOL_Edit", vbNullString)
Subj& = FindWindowEx(aochild&, CC&, "_AOL_Edit", vbNullString)
Loop Until aochild& <> 0& And ToWho& <> 0& And Messy& <> 0&
    Call SendMessageByString(ToWho&, WM_SETTEXT, 0, who$)
    DoEvents
    Call SendMessageByString(Messy&, WM_SETTEXT, 0, message$)
    DoEvents
    Call SendMessageByString(Subj&, WM_SETTEXT, 0, Subject$)
    DoEvents
    TimeOut 0.2

cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
hand1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
hand2& = FindWindowEx(aochild&, hand1&, "_AOL_Icon", vbNullString)
hand3& = FindWindowEx(aochild&, hand2&, "_AOL_Icon", vbNullString)
hand4& = FindWindowEx(aochild&, hand3&, "_AOL_Icon", vbNullString)
hand5& = FindWindowEx(aochild&, hand4&, "_AOL_Icon", vbNullString)
hand6& = FindWindowEx(aochild&, hand5&, "_AOL_Icon", vbNullString)
hand7& = FindWindowEx(aochild&, hand6&, "_AOL_Icon", vbNullString)
Hand8& = FindWindowEx(aochild&, hand7&, "_AOL_Icon", vbNullString)
Hand9& = FindWindowEx(aochild&, Hand8&, "_AOL_Icon", vbNullString)
Hand10& = FindWindowEx(aochild&, Hand9&, "_AOL_Icon", vbNullString)
Hand11& = FindWindowEx(aochild&, Hand10&, "_AOL_Icon", vbNullString)
Hand12& = FindWindowEx(aochild&, Hand11&, "_AOL_Icon", vbNullString)
Hand13& = FindWindowEx(aochild&, Hand12&, "_AOL_Icon", vbNullString)
Hand14& = FindWindowEx(aochild&, Hand13&, "_AOL_Icon", vbNullString)
Hand15& = FindWindowEx(aochild&, Hand14&, "_AOL_Icon", vbNullString)
TimeOut 0.2
Click Hand15&
End Sub

Public Function UserSN() As String
Dim cs&, Toolbar&, toolbar2&, icon1&, readmail&, mdi&, aochild&, TabControl&
Dim TabPage&, tree&, Count&, newmail&, grr$
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
Toolbar& = FindWindowEx(cs&, 0, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
icon1& = FindWindowEx(toolbar2&, 0, "_AOL_Icon", vbNullString)
readmail& = FindWindowEx(toolbar2&, icon1&, "_AOL_Icon", vbNullString)
Click readmail&
Do
DoEvents
Dim tabcont&
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
tabcont& = FindWindowEx(aochild&, 0, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(tabcont&, 0, "_AOL_TabPage", vbNullString)
tree& = FindWindowEx(TabPage&, 0, "_AOL_Tree", vbNullString)
Loop Until tabcont& <> 0& And tree& <> 0&
grr$ = GetCaption(aochild&)
Window_Close aochild&
If grr$ > "" Then
UserSN$ = Left(grr$, Len(grr$) - 17)
Exit Function
Else
UserSN$ = "not online"
End If
End Function
Public Function GetListText(Window As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(Window&, LB_GETTEXTLEN, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(Window&, LB_GETTEXT, TextLength& + 1, buffer$)
    GetListText$ = buffer$
End Function

Public Function GetCaption(Window As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(Window&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(Window&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function

Public Sub SetText(Window As Long, text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub

Public Sub Click(Whatever As Long)
    Call SendMessage(Whatever&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Whatever&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub RunMenuByString(something As String)
    Dim cs As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    cs& = FindWindow("AOL Frame25", "CompuServe 2000")
    aMenu& = GetMenu(cs&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(something$)) Then
                Call SendMessageLong(cs&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Public Function INI_Load(Section As String, Key As String, Directory As String) As String
   Dim strbuffer As String
   strbuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strbuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strbuffer, Len(strbuffer), Directory$))
End Function

Public Sub INI_Save(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub

Public Function GetText(Window As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(Window&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(Window&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetText$ = buffer$
End Function

Public Sub Window_Hide(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub

Public Sub Window_Show(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub

Public Sub GoWord(GW As String)
Dim cs&, Toolbar&, toolbar2&, Combo&, EditWin&
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
Toolbar& = FindWindowEx(cs&, 0, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
Combo& = FindWindowEx(toolbar2&, 0, "_AOL_Combobox", vbNullString)
EditWin& = FindWindowEx(Combo&, 0, "Edit", vbNullString)
SetText EditWin&, GW$
Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub ListBox_Save1(Directory As String, thelist As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub ListBox_Save2(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub

Public Sub ComboBox_Save(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub


Sub TextBox_Save(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

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

Public Function List_MailString(thelist As ListBox) As String
    Dim DoList As Long, MailString As String
    If thelist.List(0) = "" Then Exit Function
    For DoList& = 0 To thelist.ListCount - 1
        MailString$ = MailString$ & "(" & thelist.List(DoList&) & "), "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToMailString$ = MailString$
End Function

Public Sub ListBox_Load2(Directory As String, ListA As ListBox, ListB As ListBox)
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

Public Sub ComboBox_Load(ByVal Directory As String, Combo As ComboBox)
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

Public Sub ListBox_Load1(Directory As String, thelist As ListBox)
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

Sub TextBox_Load(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub

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

Public Sub TimeOut(Time As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Time
        DoEvents
    Loop
End Sub

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
    Dim tempstring As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
    tempstring$ = MyString$
    StringLength& = Len(tempstring$)
    Do While Count& <= StringLength&
        Count& = Count& + 1
        NextChr$ = Mid$(tempstring$, Count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
End Function

Public Function SwitchStrings(MyString As String, String1 As String, String2 As String) As String
    Dim tempstring As String, Spot1 As Long, Spot2 As Long
    Dim Spot As Long, ToFind As String, ReplaceWith As String
    Dim NewSpot As Long, LeftString As String, RightString As String
    Dim NewString As String
    If Len(String2) > Len(String1) Then
        tempstring$ = String1$
        String1$ = String2$
        String2$ = tempstring$
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
Public Sub Mid_Play(MidFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MidFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MidFile$, 0&, 0, 0)
    End If
End Sub

Public Sub Mid_Stop(MidFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MidFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MidFile$, 0&, 0, 0)
    End If
End Sub

Public Sub Wav_Play(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call SndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub

Public Function File_Exists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        File_Exists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        File_Exists = True
    Else
        File_Exists = False
    End If
End Function

Public Function File_GetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        File_GetAttributes% = GetAttr(TheFile$)
    End If
End Function

Public Sub File_Normal(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub File_ReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub File_Hidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
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

Public Sub Window_Close(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub

Public Function Mail_CountNew() As Long
Dim cs&, Toolbar&, toolbar2&, icon1&, readmail&, mdi&, aochild&, TabControl&
Dim TabPage&, tree&, Count&
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
Toolbar& = FindWindowEx(cs&, 0, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
icon1& = FindWindowEx(toolbar2&, 0, "_AOL_Icon", vbNullString)
readmail& = FindWindowEx(toolbar2&, icon1&, "_AOL_Icon", vbNullString)
Click readmail&
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
TabControl& = FindWindowEx(aochild&, 0, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0, "_AOL_TabPage", vbNullString)
tree& = FindWindowEx(TabPage&, 0, "_AOL_Tree", vbNullString)
If TabControl& <> 0& And TabPage& <> 0& Then
Count& = SendMessage(tree&, LB_GETCOUNT, 0&, 0&)
Mail_CountNew& = Count&
Else
End If
End Function

Public Sub IM_Send(Person As String, message As String)
Dim cs&, mdi&, aochild&, aoedit&, Rich&, icon1&, icon2&
Dim icon3&, icon4&, icon5&, icon6&, icon7&, icon8&, send&
GoWord "aol://9293:" & Person$
Do
DoEvents
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
aoedit& = FindWindowEx(aochild&, 0, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(aochild&, 0, "CS_RICHCNTL", vbNullString)
Loop Until aochild& <> 0& And aoedit& <> 0 And Rich& <> 0
SetText aoedit&, Person$
SetText Rich&, message$
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
icon1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(aochild&, icon1&, "_AOL_Icon", vbNullString)
icon3& = FindWindowEx(aochild&, icon2&, "_AOL_Icon", vbNullString)
icon4& = FindWindowEx(aochild&, icon3&, "_AOL_Icon", vbNullString)
icon5& = FindWindowEx(aochild&, icon4&, "_AOL_Icon", vbNullString)
icon6& = FindWindowEx(aochild&, icon5&, "_AOL_Icon", vbNullString)
icon7& = FindWindowEx(aochild&, icon6&, "_AOL_Icon", vbNullString)
icon8& = FindWindowEx(aochild&, icon7&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(aochild&, icon8&, "_AOL_Icon", vbNullString)
Click send&

End Sub
 
 Public Sub IM_On()
 IM_Send "$IM_ON", ";/>"
 End Sub
 
Public Sub IM_Off()
 IM_Send "$IM_OFF", ";/>"
 End Sub
 
Public Function IM_Win() As Long
Dim cs&, mdi&, aochild&, Rich&, Rich2&, Caption As String
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
Caption$ = GetCaption(aochild&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        IM_Win& = aochild&
        Exit Function
    Else
        Do
            aochild& = FindWindowEx(mdi&, aochild&, "AOL Child", vbNullString)
            Caption$ = GetCaption(aochild&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                IM_Win& = aochild&
                Exit Function
            End If
        Loop Until aochild& = 0&
End If
    IM_Win& = aochild&
End Function

Public Function IM_LastMsg() As String
    Dim Rich As Long, MsgString As String, Spot As Long
    Dim NewSpot As Long
    Rich& = FindWindowEx(IM_Win&, 0&, "CS_RICHCNTL", vbNullString)
    MsgString$ = GetText(Rich&)
    NewSpot& = InStr(MsgString$, Chr(9))
    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
    MsgString$ = Right(MsgString$, Len(MsgString$) - Spot&) '- 1)
    IM_LastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function

Public Sub IM_Respond(message As String)
    Dim IM As Long, Rich As Long, icon As Long
    IM& = IM_Win&
    If IM& = 0& Then Exit Sub
    Rich& = FindWindowEx(IM&, 0&, "CS_RICHCNTL", vbNullString)
    Rich& = FindWindowEx(IM&, Rich&, "CS_RICHCNTL", vbNullString)
    icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(IM&, icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
    DoEvents
    Call SendMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(icon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function IM_Sender() As String
    Dim IM As Long, Caption As String
    Caption$ = GetCaption(IM_Win&)
    If InStr(Caption$, ":") = 0& Then
        IM_Sender$ = ""
        Exit Function
    Else
        IM_Sender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function

Public Function IM_Text() As String
    Dim Rich As Long
    Rich& = FindWindowEx(IM_Win&, 0&, "CS_RICHCNTL", vbNullString)
    IM_Text$ = GetText(Rich&)
End Function

Public Sub IM_Ignore(who As String)
    Call IM_Send("$IM_OFF, " & who$, ";/>")
End Sub

Public Sub IM_UnIgnore(who As String)
    Call IM_Send("$IM_ON, " & who$, ";/>")
End Sub

Public Sub RunMenu(TopMenu As Long, SubMenu As Long)
    Dim cs As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    cs& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(cs&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(cs&, WM_COMMAND, mnID&, 0&)
End Sub

Public Function IM_Check(who As String) As Boolean
    Dim cs As Long, mdi As Long, IM As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    cs& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(cs&, 0&, "MDIClient", vbNullString)
    Call GoWord("aol://9293:" & who$)
    Do
        DoEvents
        IM& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "CS_RICHCNTL", vbNullString)
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
    Loop Until IM& <> 0& And Rich <> 0& And Available& <> 0& And Available& <> Available1& And Available& <> Available2& And Available& <> Available3&
    DoEvents
    Call SendMessage(Available&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Available&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        oWindow& = FindWindow("#32770", "Compuserve")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        IM_Check = True
    Else
        IM_Check = False
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End Function

Public Sub Form_OnTop(Form As Form)
Call SetWindowPos(Form.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub Form_NotOnTop(Form As Form)
Call SetWindowPos(Form.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub

Public Sub BuddyList_Block(who As String)
Dim cs&, mdi&, aochild&, icon1&, icon2&, hand1&, hand2&, hand3&, hand4&, hand5&, hand6&, hand7&
Dim edit&, icon3&
GoWord "contact view"
TimeOut 2
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
icon1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(aochild&, icon1&, "_AOL_Icon", vbNullString)
Click icon2&
Window_Close aochild&
TimeOut 2
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
hand1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
hand2& = FindWindowEx(aochild&, hand1&, "_AOL_Icon", vbNullString)
hand3& = FindWindowEx(aochild&, hand2&, "_AOL_Icon", vbNullString)
hand4& = FindWindowEx(aochild&, hand3&, "_AOL_Icon", vbNullString)
hand5& = FindWindowEx(aochild&, hand4&, "_AOL_Icon", vbNullString)
hand6& = FindWindowEx(aochild&, hand5&, "_AOL_Icon", vbNullString)
hand7& = FindWindowEx(aochild&, hand6&, "_AOL_Icon", vbNullString)
Click hand7&
Window_Close aochild&
TimeOut 2
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)


edit& = FindWindowEx(aochild&, 0, "_AOL_Edit", vbNullString)
SetText edit&, who$
icon1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(aochild&, icon1&, "_AOL_Icon", vbNullString)
icon3& = FindWindowEx(aochild&, icon2&, "_AOL_Icon", vbNullString)
Click icon1&
TimeOut 2
hand1& = FindWindowEx(aochild&, 0, "_AOL_Checkbox", vbNullString)
hand2& = FindWindowEx(aochild&, hand1&, "_AOL_Checkbox", vbNullString)
hand3& = FindWindowEx(aochild&, hand2&, "_AOL_Checkbox", vbNullString)
Click hand3&
Click icon3&
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
Dim okwin&, OKButton&
Do
  DoEvents
  okwin& = FindWindow("#32770", "Compuserve")
  OKButton& = FindWindowEx(okwin&, 0&, "Button", "OK")
  Loop Until okwin& <> 0& And OKButton& <> 0&
Click OKButton&
End Sub
Public Sub BuddyList_Add(group As String, who As String)
Dim cs&, mdi&, aochild&, icon1&, icon2&, hand1&, hand2&
Dim hand3&, hand4&, hand5&, hand6&, hand7&, edit&, icon3&, Combo&
GoWord "contact view"
TimeOut 2
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
icon1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(aochild&, icon1&, "_AOL_Icon", vbNullString)
Click icon2&
Window_Close aochild&
TimeOut 2
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
icon1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
Click icon1&
Window_Close aochild&
Dim edit1&, edit2&
Do
cs& = FindWindow("AOL Frame25", "CompuServe 2000")
mdi& = FindWindowEx(cs&, 0, "MDIClient", vbNullString)
aochild& = FindWindowEx(mdi&, 0, "AOL Child", vbNullString)
edit1& = FindWindowEx(aochild&, 0, "_AOL_Edit", vbNullString)
edit2& = FindWindowEx(aochild&, edit1&, "_AOL_Edit", vbNullString)
icon1& = FindWindowEx(aochild&, 0, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(aochild&, icon1&, "_AOL_Icon", vbNullString)
icon3& = FindWindowEx(aochild&, icon2&, "_AOL_Icon", vbNullString)
Loop Until edit1& <> 0& And edit2& <> 0&
SetText edit1&, group$
SetText edit2&, who$
Click icon1&
TimeOut 1
Click icon3&

End Sub
Public Sub KillWait()
RunMenuByString "A&bout Compuserve"
End Sub
Function File_Scan(TheFile As String, Searchstring As String) As Long
Free = FreeFile
Dim Where As Long
Open TheFile$ For Binary Access Read As #Free
For X = 1 To LOF(Free) Step 32000
    text$ = Space(32000)
    Get #Free, X, text$
    Debug.Print X
    If InStr(1, text$, Searchstring$, 1) Then
        Where = InStr(1, text$, Searchstring$, 1)
        ScanFile = (Where + X) - 1
        Close #Free
        Exit For
    End If
    Next X
Close #Free
End Function

Function Text_Decrypt(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If crapp% > 0 Then GoTo dustepp2
If NextChr$ = "~" Then Let NextChr$ = "A"
If NextChr$ = "`" Then Let NextChr$ = "a"
If NextChr$ = "!" Then Let NextChr$ = "B"
If NextChr$ = "@" Then Let NextChr$ = "c"
If NextChr$ = "#" Then Let NextChr$ = "c"
If NextChr$ = "$" Then Let NextChr$ = "D"
If NextChr$ = "%" Then Let NextChr$ = "d"
If NextChr$ = "^" Then Let NextChr$ = "E"
If NextChr$ = "&" Then Let NextChr$ = "e"
If NextChr$ = "*" Then Let NextChr$ = "f"
If NextChr$ = "(" Then Let NextChr$ = "H"
If NextChr$ = ")" Then Let NextChr$ = "I"
If NextChr$ = "-" Then Let NextChr$ = "i"
If NextChr$ = "_" Then Let NextChr$ = "k"
If NextChr$ = "+" Then Let NextChr$ = "L"
If NextChr$ = "=" Then Let NextChr$ = "M"
If NextChr$ = "[" Then Let NextChr$ = "m"
If NextChr$ = "]" Then Let NextChr$ = "N"
If NextChr$ = "{" Then Let NextChr$ = "n"
If NextChr$ = "O" Then Let NextChr$ = "}"
If NextChr$ = "\" Then Let NextChr$ = "o"
If NextChr$ = "|" Then Let NextChr$ = "P"
If NextChr$ = ";" Then Let NextChr$ = "p"
If NextChr$ = "'" Then Let NextChr$ = "r"
If NextChr$ = ":" Then Let NextChr$ = "S"
If NextChr$ = """" Then Let NextChr$ = "s"
If NextChr$ = "," Then Let NextChr$ = "t"
If NextChr$ = "." Then Let NextChr$ = "U"
If NextChr$ = "/" Then Let NextChr$ = "u"
If NextChr$ = "<" Then Let NextChr$ = "V"
If NextChr$ = ">" Then Let NextChr$ = "v"
If NextChr$ = "?" Then Let NextChr$ = "w"
If NextChr$ = "¥" Then Let NextChr$ = "x"
If NextChr$ = "Ä" Then Let NextChr$ = "X"
If NextChr$ = "ƒ" Then Let NextChr$ = "Y"
If NextChr$ = "Ü" Then Let NextChr$ = "y"
If NextChr$ = "¶" Then Let NextChr$ = "!"
If NextChr$ = "£" Then Let NextChr$ = "?"
If NextChr$ = "…" Then Let NextChr$ = "."
If NextChr$ = "æ" Then Let NextChr$ = ","
If NextChr$ = "q" Then Let NextChr$ = "1"
If NextChr$ = "w" Then Let NextChr$ = "%"
If NextChr$ = "e" Then Let NextChr$ = "2"
If NextChr$ = "r" Then Let NextChr$ = "3"
If NextChr$ = "t" Then Let NextChr$ = "_"
If NextChr$ = "y" Then Let NextChr$ = "-"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$
dustepp2:
If cra% > 0 Then Let cra% = cra% - 1
DoEvents
Loop
r_decrypt = newsent$
End Function

Function Text_Encrypt(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let NextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If crapp% > 0 Then GoTo dustepp2
If NextChr$ = "A" Then Let NextChr$ = "~"
If NextChr$ = "a" Then Let NextChr$ = "`"
If NextChr$ = "B" Then Let NextChr$ = "!"
If NextChr$ = "C" Then Let NextChr$ = "@"
If NextChr$ = "c" Then Let NextChr$ = "#"
If NextChr$ = "D" Then Let NextChr$ = "$"
If NextChr$ = "d" Then Let NextChr$ = "%"
If NextChr$ = "E" Then Let NextChr$ = "^"
If NextChr$ = "e" Then Let NextChr$ = "&"
If NextChr$ = "f" Then Let NextChr$ = "*"
If NextChr$ = "H" Then Let NextChr$ = "("
If NextChr$ = "I" Then Let NextChr$ = ")"
If NextChr$ = "i" Then Let NextChr$ = "-"
If NextChr$ = "k" Then Let NextChr$ = "_"
If NextChr$ = "L" Then Let NextChr$ = "+"
If NextChr$ = "M" Then Let NextChr$ = "="
If NextChr$ = "m" Then Let NextChr$ = "["
If NextChr$ = "N" Then Let NextChr$ = "]"
If NextChr$ = "n" Then Let NextChr$ = "{"
If NextChr$ = "O" Then Let NextChr$ = "}"
If NextChr$ = "o" Then Let NextChr$ = "\"
If NextChr$ = "P" Then Let NextChr$ = "|"
If NextChr$ = "p" Then Let NextChr$ = ";"
If NextChr$ = "r" Then Let NextChr$ = "'"
If NextChr$ = "S" Then Let NextChr$ = ":"
If NextChr$ = "s" Then Let NextChr$ = """"
If NextChr$ = "t" Then Let NextChr$ = ","
If NextChr$ = "U" Then Let NextChr$ = "."
If NextChr$ = "u" Then Let NextChr$ = "/"
If NextChr$ = "V" Then Let NextChr$ = "<"
If NextChr$ = "W" Then Let NextChr$ = ">"
If NextChr$ = "w" Then Let NextChr$ = "?"
If NextChr$ = "X" Then Let NextChr$ = "¥"
If NextChr$ = "x" Then Let NextChr$ = "Ä"
If NextChr$ = "Y" Then Let NextChr$ = "ƒ"
If NextChr$ = "y" Then Let NextChr$ = "Ü"
If NextChr$ = "!" Then Let NextChr$ = "¶"
If NextChr$ = "?" Then Let NextChr$ = "£"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "æ"
If NextChr$ = "1" Then Let NextChr$ = "q"
If NextChr$ = "%" Then Let NextChr$ = "w"
If NextChr$ = "2" Then Let NextChr$ = "e"
If NextChr$ = "3" Then Let NextChr$ = "r"
If NextChr$ = "_" Then Let NextChr$ = "t"
If NextChr$ = "-" Then Let NextChr$ = "y"
If NextChr$ = " " Then Let NextChr$ = " "
Let newsent$ = newsent$ + NextChr$
dustepp2:
If Crap% > 0 Then Let Crap% = Crap% - 1
DoEvents
Loop
r_encrypt = newsent$
End Function

Sub Form_FadeBlink(theform As Form)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
Dim a, b, i, y
theform.BackColor = &H0&
theform.DrawStyle = 6
theform.DrawMode = 13

theform.DrawWidth = 2
theform.ScaleMode = 3
theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, b)-(theform.Width, b + 2), RGB(a + 3, a, a * 3), BF

b = b + 2
Next a

For i = 255 To 0 Step -1
theform.Line (0, 0)-(theform.Width, y + 2), RGB(i + 3, i, i * 3), BF
y = y + 2
Next i

End Sub
Sub Form_FadeBlack(theform As Form)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
theform.BackColor = &H0&
theform.DrawStyle = 6
theform.DrawMode = 13

theform.DrawWidth = 2
theform.ScaleMode = 3
theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, b)-(theform.Width, b + 2), RGB(a + 3, a, a * 3), BF

b = b + 2
Next a

For i = 255 To 0 Step -1
theform.Line (0, 0)-(theform.Width, y + 2), RGB(i + 3, i, i * 3), BF
y = y + 2
Next i
theform.BackColor = &H0&
theform.DrawStyle = 6
theform.DrawMode = 13

theform.DrawWidth = 2
theform.ScaleMode = 3
theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, b)-(theform.Width, b + 2), RGB(a + 3, a, a * 3), BF

b = b + 2
Next a

For i = 255 To 0 Step -1
theform.Line (0, 0)-(theform.Width, y + 2), RGB(i + 3, i, i * 3), BF
y = y + 2
Next i

End Sub
Sub Form_FadeBW(theform)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
theform.ScaleHeight = (256 * 2)
For a = 255 To 0 Step -1
theform.Line (0, b)-(theform.Width, b + 1), RGB(a + 1, a, a * 1), BF
b = b + 2
Next a

End Sub

Sub Form_FadeBlue(vForm As Form)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub Form_FadeFire(vForm As Object)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255

        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B
        Next intLoop
End Sub
Sub Form_FadeGreen(vForm As Form)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub Form_FadeGrey(vForm As Form)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub

Sub Form_FadeIce(vForm As Object)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
   
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B
        Next intLoop
End Sub

Sub Form_FadePlatinum(vForm As Object)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
                vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub
Sub Form_FadePurple(vForm As Form)
'This function cannot be done in Form_Load ()
'You may use Form_Resize() as a replacement
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub File_Destroy(sFileName As String)
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
        Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1
        For iLoop = 1 To Blocks
            offset = Seek(hFileHandle)
            Put hFileHandle, , Block1
            Put hFileHandle, offset, Block2
        Next iLoop
    Close hFileHandle
        Kill sFileName
End Sub

