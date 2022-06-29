Attribute VB_Name = "Chronx3"
'Chronx.bas v3.0 by Chron__x
'                *32 BIT*
'Coded in Visual Basic 6.0 Pro. For use with AOL 4.0
'Aim- iichronx
'Email- Chron__x@hotmail.com
'Webpage- Http://chronx.cjb.net

'                                                        **Chat Ocx Coming Soon**
'I need help with the name for it, if you can think of one please tell me and you will be put in the greets

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Public Declare Function CharNext Lib "user32" Alias "CharNextA" (ByVal lpsz As String) As String
Public Declare Function CharPrev Lib "user32" Alias "CharPrevA" (ByVal lpszStart As String, ByVal lpszCurrent As String) As String
Public Declare Function CharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetProfileString& Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long)
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
Public Declare Function GetTextColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer

Public Declare Function IsCharAlpha Lib "user32" Alias "IsCharAlphaA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharAlphaNumeric Lib "user32" Alias "IsCharAlphaNumericA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharLower Lib "user32" Alias "IsCharLowerA" (ByVal cChar As Byte) As Long
Public Declare Function IsCharUpper Lib "user32" Alias "IsCharUpperA" (ByVal cChar As Byte) As Long
Public Declare Function IsChild Lib "user32" (ByVal hWndParent As Long, ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function MciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Public Declare Function VkKeyScanEx Lib "user32" Alias "VkKeyScanExA" (ByVal Ch As Byte, ByVal dwhkl As Long) As Integer

Public Declare Function WindowFromPointXy Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
 ByVal bRedraw As Boolean) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const Cb_AddString& = &H143
Public Const Cb_DeleteString& = &H144
Public Const Cb_FindStringExact& = &H158
Public Const Cb_GetCount& = &H146
Public Const Cb_GetItemData = &H150
Public Const Cb_GetLbText& = &H148
Public Const Cb_ResetContent& = &H14B
Public Const Cb_SetCursel& = &H14E

Public Const Em_GetLineCount& = &HBA
Public Const ENTER_KEY = 13

Public Const Hwnd_NotTopMost = -2
Public Const HWND_TOPMOST = -1

Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_INSERTSTRING = &H181
Public Const Lb_ResetContent& = &H184
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const MF_BYPOSITION = &H400&

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const Snd_Flags = SND_ASYNC Or SND_NODEFAULT
Public Const Snd_Flags2 = SND_ASYNC Or SND_LOOP

Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE& = 3
Public Const SW_MINIMIZE& = 6
Public Const SW_RESTORE& = 9
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Swp_Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Const Sys_Add = &H0
Public Const Sys_Delete = &H2
Public Const Sys_Message = &H1
Public Const Sys_Icon = &H2
Public Const Sys_Tip = &H4

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &HF012
Public Const Wm_RButtonDblClk& = &H206
Public Const Wm_RButtonDown& = &H204
Public Const Wm_RButtonUp& = &H205
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_USER& = &H400

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const Op_Flags = PROCESS_READ Or RIGHTS_REQUIRED

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

Public SysTray As NOTIFYICONDATA

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type


Public Sub WinHide(WinHandle As Long)
    Call ShowWindow(WinHandle&, SW_HIDE)
End Sub

Public Sub WinShow(WinHandle As Long)
    Call ShowWindow(WinHandle&, SW_SHOW)
End Sub

Public Sub WinRestore(WinHandle As Long)
    Call ShowWindow(WinHandle&, SW_RESTORE)
End Sub

Public Sub WinMinimize(WinHandle As Long)
    Call ShowWindow(WinHandle&, SW_MINIMIZE)
End Sub

Public Sub WinMaximize(WinHandle As Long)
    Call ShowWindow(WinHandle&, SW_MAXIMIZE)
End Sub

Public Sub WinEnable(WinHandle As Long)
    Call EnableWindow(WinHandle&, 1&)
End Sub

Public Sub WinDisable(WinHandle As Long)
    Call EnableWindow(WinHandle&, 0&)
End Sub

Public Sub WinClose(WinHandle As Long)
    Call PostMessage(WinHandle&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub WinBringToTop(WinHandle As Long)
    Call BringWindowToTop(WinHandle&)
End Sub

Public Function WinMaximized(WinHandle As Long) As Boolean
    Dim MaxVal As Long
    MaxVal& = IsZoomed(WinHandle&)
    If MaxVal& > 0& Then
        WinMaximized = True
       ElseIf MaxVal& = 0& Then
        WinMaximized = False
    End If
End Function
Function MailString(list As ListBox) As String
'Use this with Mail to Bcc it
Dim X As Integer, prepstring As String
For X = 0 To list.ListCount - 1
    prepstring = prepstring & "((" & list.list(X) & ")),"
Next X
MailString = prepstring
End Function
Public Function WinMinimized(WinHandle As Long) As Boolean
    Dim MinVal As Long
    MinVal& = IsIconic(WinHandle&)
    If MinVal& > 0& Then
        WinMinimized = True
       ElseIf MinVal& = 0& Then
        WinMinimized = False
    End If
End Function

Public Sub FormOntop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Swp_Flags)
End Sub

Public Sub FormNotTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, Hwnd_NotTopMost, 0&, 0&, 0&, 0&, Swp_Flags)
End Sub

Public Sub FormDrag(FormName As Form)
    Call ReleaseCapture
    Call PostMessage(FormName.hwnd, WM_SYSCOMMAND, WM_MOVE, 0&)
End Sub

Public Sub FormCenter(CenterMe As Form)
    CenterMe.Top = (Screen.Height * 0.85) / 2 - CenterMe.Height / 2
    CenterMe.Left = Screen.Width / 2 - CenterMe.Width / 2
End Sub

Public Function GetCaption(WinHandle As Long) As String
    Dim Buffer As String, TextLen As Long
    TextLen& = GetWindowTextLength(WinHandle&)
    Buffer$ = String(TextLen&, 0&)
    Call GetWindowText(WinHandle&, Buffer$, TextLen& + 1)
    GetCaption$ = Buffer$
End Function

Public Function GetText(WinHandle As Long) As String
    Dim Buffer As String, TextLen As Long
    TextLen& = SendMessageByNum(WinHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(TextLen&, 0&)
    Call SendMessageByString(WinHandle&, WM_GETTEXT, TextLen& + 1, Buffer$)
    GetText$ = Buffer$
End Function

Public Sub ClickIcon(IconHandle As Long)
    Call PostMessage(IconHandle&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(IconHandle&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub ClickButton(ButtonHandle As Long)
    Call PostMessage(ButtonHandle&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(ButtonHandle, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Function LineCountHwnd(WinHandle As Long) As Long
    Dim CurrentCount As Long
    On Local Error Resume Next
    CurrentCount& = SendMessageLong(WinHandle&, Em_GetLineCount, 0&, 0&)
    LineCountHwnd& = Format$(CurrentCount&, "##,###")
End Function

Public Function LineCharacter(InThisString As String, CharacterPos As Long) As String
    If CharacterPos& > Len(InThisString$) Then Exit Function
    LineCharacter$ = Right(Left(InThisString$, CharacterPos&), 1)
End Function

Public Function LineCount(StringToCount As String) As Long
    If Len(StringToCount$) = 0& Then
        LineCount& = 0
        Exit Function
    End If
    LineCount& = StringCount(StringToCount$, vbCr) + 1
End Function

Public Function ListCount(ListBox As Long) As Long
    ListCount& = SendMessageLong(ListBox&, LB_GETCOUNT, 0&, 0&)
End Function

Public Function ComboCount(ComboBox As Long) As Long
    ComboCount& = SendMessageLong(ComboBox&, Cb_GetCount, 0&, 0&)
End Function

Public Sub ListSetFocus(ListBox As Long, ListIndex As Long)
    Call SendMessageLong(ListBox&, LB_SETCURSEL, ListIndex&, 0&)
End Sub

Public Sub ComboSetFocus(ComboBox As Long, ListIndex As Long)
    Call SendMessageLong(ComboBox&, Cb_SetCursel, ListIndex&, 0&)
End Sub

Public Sub ListMouseMoveTip(Form As Form, ListBox As ListBox, Y As Single)
    Dim YPos As Long, OldFontSize As Integer
    OldFontSize = Form.Font.Size
    Form.Font.Size = ListBox.Font.Size
    YPos& = Y \ Form.TextHeight("Xyz") + ListBox.TopIndex
    Form.Font.Size = OldFontSize
    If YPos& < ListBox.ListCount Then
        ListBox.ToolTipText = ListBox.list(YPos&)
       ElseIf YPos& > ListBox.ListCount Then
        ListBox.ToolTipText = ""
    End If
End Sub

Public Sub SystemFonts(ListOrComboBox As Control)
    Dim CurrentFontNumber As Long
    For CurrentFontNumber& = 0 To Screen.FontCount - 1
        ListOrComboBox.AddItem Screen.Fonts(CurrentFontNumber&)
    Next CurrentFontNumber&
End Sub

Public Sub SetText(WinHandle As Long, StringToSet As String, Optional ClearBefore As Boolean = True)
    If ClearBefore = True Then Call SendMessageByString(WinHandle&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(WinHandle&, WM_SETTEXT, 0&, StringToSet$)
End Sub

Public Sub Yield(StopDuration As Long)
    Dim InitialTime As Long
    InitialTime& = Timer
    Do Until Timer - InitialTime& >= StopDuration&
        DoEvents
    Loop
End Sub
Public Sub AutoListerPr(RoomList As Control, SnList As Control, Optional BustIfFull As Boolean = True, Optional LimitTriesOnBust As Long = "2")
    Dim ListIndex As Long
    If FindRoom& <> 0& Then AddRoomToList SnList, False
    For ListIndex& = 0 To RoomList.ListCount - 1
       If BustIfFull = False Then
           Call KeyWord("aol://2719:2-2-" & RoomList.list(ListIndex&))
           WaitForOkorRoom RoomList.list(ListIndex&)
           timeout 2
           AddRoomToList SnList, False
          ElseIf BustIfFull = True Then
           Call RoomForceEnter("aol://2719:2-2-", RoomList.list(ListIndex&), False, 0.2, LimitTriesOnBust&)
           AddRoomToList SnList, False
       End If
     
       timeout 2
    Next ListIndex&
End Sub
Public Sub MailSendBCC(Person As String, BCC As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
    Dim AolFrame As Long, AolToolbar As Long, Toolbar As Long
    Dim AolIcon1 As Long, AolEdit1 As Long
    Dim AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim AolIcon2 As Long, AolModal As Long, AolIcon3 As Long
    Dim CheckBox As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon1& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
    Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
        RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
        AolIcon2& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
    Loop Until FindSendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0& And AolIcon2& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, Person$)
    timeout (0.5)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, BCC$)
    timeout (0.2)
    Call SendMessageByString(AolEdit3&, WM_SETTEXT, 0&, Subject$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If CheckReturnReceipts = True Then
        CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
    End If
    timeout (0.2)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolIcon3& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolIcon3& <> 0&
    If AolModal& <> 0& And FindWindowEx(AolMdi&, 0&, "AOL Child", "Write Mail") = 0& Then
        Call PostMessage(AolIcon3&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon3&, WM_LBUTTONUP, 0&, 0&)
        Exit Sub
       ElseIf FindWindowEx(AolMdi&, 0&, "AOL Child", "Write Mail") = 0& And AolModal& = 0& Then
        Exit Sub
    End If
    
End Sub


Public Function ExtractNumeric(TheString As String) As String
    Dim Instance As Long, FoundNumeric As String, NewString As String
    For Instance& = 1 To Len(TheString$)
       FoundNumeric$ = Mid(TheString$, Instance&, 1)
       If IsNumeric(FoundNumeric$) = True Then
           NewString$ = NewString$ & FoundNumeric$
       End If
    Next Instance&
    ExtractNumeric$ = NewString$
End Function

Public Function ExtractAlpha(TheString As String) As String
    Dim Instance As Long, FoundAlpha As String, NewString As String
    For Instance& = 1 To Len(TheString$)
    FoundAlpha$ = Mid(TheString$, Instance&, 1)
       If IsNumeric(FoundAlpha$) = False Then
           NewString$ = NewString$ & FoundAlpha$
       End If
    Next Instance&
    ExtractAlpha$ = NewString$
End Function

Public Function ChrValues(ConvertString As String) As String
    Dim StringLen As Long, StringOutPut As String, handler
    On Error GoTo handler
    For StringLen& = 1 To Len(ConvertString$)
        StringOutPut$ = StringOutPut$ & "Chr(" & CStr(Asc(Mid(ConvertString$, StringLen&, 1))) & ") & "
    Next StringLen&
    StringOutPut$ = Trim(StringOutPut$)
    StringOutPut$ = Mid(StringOutPut$, 1, Len(StringOutPut$) - 2)
    ChrValues$ = StringOutPut$
handler:
End Function

Public Function ReplaceCharacters(MainString As String, ToLookFor As String, ToReplaceWith As String) As String
    Dim NewMain As String, Instance As Long
    NewMain$ = MainString$
    Do While InStr(1, NewMain$, ToLookFor$)
       DoEvents
       Instance& = InStr(1, NewMain$, ToLookFor$)
       NewMain$ = Left(NewMain$, (Instance& - 1)) & ToReplaceWith$ & Right(NewMain$, Len(NewMain$) - (Instance& + Len(ToLookFor$) - 1))
    Loop
    ReplaceCharacters$ = NewMain$
End Function

Public Sub ListRemoveSelected(ListBox As ListBox)
    Dim ListCount As Long
    ListCount& = ListBox.ListCount
    Do While ListCount& > 0&
       ListCount& = ListCount& - 1
       If ListBox.Selected(ListCount&) = True Then
          ListBox.RemoveItem (ListCount&)
       End If
    Loop
End Sub

Public Function FindChildByClass(ParentWindow As Long, ClassWindow As String) As Long
'For those who dont do 32 bit api i converted these oldschool 16 bit methods for you
    FindChildByClass& = FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString)
End Function

Public Function FindChildByTitleEx(ParentWindow As Long, WindowText As String) As Long
'For those who dont do 32 bit api i converted these oldschool 16 bit methods for you, with a twist
    FindChildByTitleEx& = FindWindowEx(ParentWindow&, 0&, vbNullString, WindowText$)
End Function

Public Function ClassInstance(ParentWindow As Long, ClassWindow As String) As Long
     Dim OnInstance As Long, CurrentCount As Long
     If FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString) = 0& Then Exit Function
     ClassInstance& = 0&
     Do: DoEvents
         OnInstance& = FindWindowEx(ParentWindow&, OnInstance&, ClassWindow$, vbNullString)
         If OnInstance& <> 0& Then
             CurrentCount& = CurrentCount& + 1
            Else
             Exit Do
         End If
     Loop
     ClassInstance& = CurrentCount&
End Function

Public Function ChildInstance(ParentWindow As Long) As Long
     Dim OnInstance As Long, CurrentCount As Long
     If IsWindow(ParentWindow&) = 0& Then Exit Function
     OnInstance& = GetWindow(ParentWindow&, 5)
     If OnInstance& <> 0& Then ChildInstance& = 1
     Do: DoEvents
         OnInstance& = GetWindow(OnInstance&, 2)
         If OnInstance& <> 0& Then CurrentCount& = CurrentCount& + 1
     Loop Until OnInstance& = 0&
     ChildInstance& = CurrentCount& + 1
End Function

Public Function NextOfClassByCount(ParentWindow As Long, ClassWindow As String, ByCount As Long) As Long
    Dim NextOfClass As Long, NextWindow As Long
    If ByCount& > ClassInstance(ParentWindow&, ClassWindow$) Then Exit Function
    If FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString) = 0& Then Exit Function
    For NextOfClass& = 1 To ByCount&
        NextWindow& = FindWindowEx(ParentWindow&, NextWindow&, ClassWindow$, vbNullString)
    Next NextOfClass&
    NextOfClassByCount& = NextWindow&
End Function

Public Function ReverseString(TextToReverse As String) As String
    Dim Step As Long, NewString As String
    For Step& = 1 To Len(TextToReverse$)
       NewString$ = Mid(TextToReverse$, Step&, 1) & NewString$
    Next Step&
    ReverseString$ = NewString$
End Function

Public Function InsertCharacters(InThisString As String, CharactersToInsert As String) As String
    Dim NewInString As String, StringLength As Long, NextLen As Long
    Dim DownString As String, prepstring As String
    NewInString$ = InThisString$
    StringLength& = Len(NewInString$)
    Do While NextLen& <= StringLength&
        NextLen& = NextLen& + 1
        DownString$ = Mid(NewInString$, NextLen&, 1)
        DownString$ = DownString$ & CharactersToInsert$
        prepstring$ = prepstring$ & DownString$
    Loop
    InsertCharacters$ = Left(prepstring$, Len(prepstring$) - 2)
End Function

Public Function ListToTextString(ListBox As ListBox, InsertSeparator As String) As String
    Dim CurrentCount As Long, prepstring As String
    For CurrentCount& = 0 To ListBox.ListCount - 1
        prepstring$ = prepstring$ & ListBox.list(CurrentCount&) & InsertSeparator$
    Next CurrentCount&
    ListToTextString$ = Left(prepstring$, Len(prepstring$) - 2)
End Function

Public Function ComboToTextString(ComboBox As ListBox, InsertSeparator As String) As String
    Dim CurrentCount As Long, prepstring As String
    For CurrentCount& = 0 To ComboBox.ListCount - 1
        prepstring$ = prepstring$ & ComboBox.list(CurrentCount&) & InsertSeparator$
    Next CurrentCount&
    ComboToTextString$ = Left(prepstring$, Len(prepstring$) - 2)
End Function

Public Sub ListSearchScroll(ListBox As ListBox, SearchString As String, Optional Delay As Single = "0.6")
     Dim Search  As Long
     For Search& = 0 To ListBox.ListCount - 1
         If InStr(LCase(ListBox.list(Search&)), LCase(SearchString$)) <> 0& Then Call RoomSend(ListBox.list(Search&))
         Call Yield(Val(Delay))
     Next Search&
End Sub

Public Function Findchildbytitle(ParentWindow As Long, WindowText As String) As Long
'NOT RECOMENDED FOR USE, THE 32 BIT API IN THIS MODULE IS MORE ACCURATE AND FASTER
    Dim GetChild As Long, GetNextChild As Long, PrepLong As Long
    GetChild& = GetWindow(ParentWindow&, 5)
    If UCase(GetCaption(GetChild&)) Like UCase(WindowText$) Then Findchildbytitle& = GetChild&
    GetChild& = GetWindow(ParentWindow&, 5)
    While ParentWindow&
        GetNextChild& = GetWindow(ParentWindow&, 5)
        If UCase(GetCaption(GetNextChild&)) Like UCase(WindowText$) & "*" Then Findchildbytitle& = GetChild&
           GetChild& = GetWindow(ParentWindow&, 5)
        If UCase(GetCaption(GetChild&)) Like UCase(WindowText$) & "*" Then Findchildbytitle& = GetChild&
    Wend
    Findchildbytitle& = 0&
End Function

Public Function GetMessageText(MessageWindow As Long) As String
    Dim StaticWindow1 As Long, StaticWindow2 As Long, AolStatic1 As Long
    Dim AolStatic2 As Long
    StaticWindow1& = FindWindowEx(MessageWindow&, 0&, "Static", vbNullString)
    StaticWindow2& = FindWindowEx(MessageWindow&, StaticWindow1&, "Static", vbNullString)
    AolStatic1& = FindWindowEx(MessageWindow&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(MessageWindow&, AolStatic1&, "_AOL_Static", vbNullString)
    If StaticWindow2& <> 0& Then
        GetMessageText$ = GetText(StaticWindow2&)
       ElseIf AolStatic2& <> 0& Then
        GetMessageText$ = GetText(AolStatic2&)
    End If
End Function

Public Function GetInstance(InThisString As String, CharacterInstance As String, InstanceNumber As Long) As String
    Dim Instance As Long, FindInstance As Long, NewInstance As Long
    If InstanceNumber& < 1 Then
        GetInstance$ = ""
        Exit Function
    End If
    Instance& = 0&
    For FindInstance& = 1 To InstanceNumber&
        NewInstance& = Instance&
        Instance& = InStr(NewInstance& + 1, InThisString$, CharacterInstance$)
        If Instance& = 0& Then
            If FindInstance& = InstanceNumber& Then
                 GetInstance$ = Mid(InThisString$, NewInstance& + 1, Len(InThisString$) - NewInstance&)
                Else
                 GetInstance$ = ""
            End If
            Exit Function
        End If
    Next FindInstance&
    GetInstance$ = Mid(InThisString$, NewInstance& + 1, Instance& - NewInstance& - 1)
End Function

Public Sub CheckBoxSetValue(CheckBox As Long, CheckValue As Boolean)
    Call PostMessage(CheckBox&, BM_SETCHECK, CheckValue, 0&)
End Sub

Public Function CheckBoxGetValue(CheckBox As Long) As Boolean
    Dim CheckValue As Long
    CheckValue& = SendMessageLong(CheckBox&, BM_GETCHECK, 0&, 0&)
    If CheckValue& = 0& Then
        CheckBoxGetValue = False
       ElseIf CheckValue& <> 0& Then
        CheckBoxGetValue = True
    End If
End Function



Public Function TrimNull(MainString As String) As String
    Dim NewMain As String, Instance As Long
    NewMain$ = MainString$
    Do While InStr(1, NewMain$, vbNullChar)
        DoEvents
        Instance& = InStr(1, NewMain$, vbNullChar)
        NewMain$ = Left(NewMain$, (Instance& - 1)) & "" & Right(NewMain$, Len(NewMain$) - Instance&)
    Loop
    TrimNull$ = NewMain$
End Function

Public Function TrimSpaces(MainString As String) As String
    Dim NewMain As String, Instance As Long
    NewMain$ = MainString$
    Do While InStr(1, NewMain$, " ")
        DoEvents
        Instance& = InStr(1, NewMain$, " ")
        NewMain$ = Left(NewMain$, (Instance& - 1)) & "" & Right(NewMain$, Len(NewMain$) - Instance&)
    Loop
    TrimSpaces$ = NewMain$
End Function

Public Sub ComboCopy(SourceCombo As Long, DestinationCombo As Long)
    Dim SourceCount As Long, OfCountForIndex As Long, FixedString As String
    SourceCount& = SendMessageLong(SourceCombo&, Cb_GetCount, 0&, 0&)
    Call SendMessageLong(DestinationCombo&, Cb_ResetContent, 0&, 0&)
    If SourceCount& = 0& Then Exit Sub
    For OfCountForIndex& = 0 To SourceCount& - 1
        FixedString$ = String(250, 0)
        Call SendMessageByString(SourceCombo&, Cb_GetLbText, OfCountForIndex&, FixedString$)
        Call SendMessageByString(DestinationCombo&, Cb_AddString, 0&, FixedString$)
    Next OfCountForIndex&
End Sub

Public Sub ListCopy(SourceList As Long, DestinationList As Long)
    Dim SourceCount As Long, OfCountForIndex As Long, FixedString As String
    SourceCount& = SendMessageLong(SourceList&, LB_GETCOUNT, 0&, 0&)
    Call SendMessageLong(DestinationList&, Lb_ResetContent, 0&, 0&)
    If SourceCount& = 0& Then Exit Sub
    For OfCountForIndex& = 0 To SourceCount& - 1
        FixedString$ = String(250, 0)
        Call SendMessageByString(SourceList&, LB_GETTEXT, OfCountForIndex&, FixedString$)
        Call SendMessageByString(DestinationList&, LB_ADDSTRING, 0&, FixedString$)
    Next OfCountForIndex&
End Sub

Public Sub ListKillDuplicates(ListBox As ListBox)
'The higher the number of items in a list the more time it will take and the more likely your program will crash, i havent found a way to fix this problem, and i dont think anyone ever will
    Dim FirstCount As Long, SecondCount As Long
    On Error Resume Next
    For FirstCount& = 0& To ListBox.ListCount - 1
        For SecondCount& = 0& To ListBox.ListCount - 1
            If LCase(ListBox.list(FirstCount&)) Like LCase(ListBox.list(SecondCount&)) And FirstCount& <> SecondCount& Then
                ListBox.RemoveItem SecondCount&
            End If
        Next SecondCount&
    Next FirstCount&
End Sub

Public Sub ComboKillDuplicates(ComboBox As ComboBox)
    Dim FirstCount As Long, SecondCount As Long
    On Error Resume Next
    For FirstCount& = 0& To ComboBox.ListCount - 1
        For SecondCount& = 0& To ComboBox.ListCount - 1
            If LCase(ComboBox.list(FirstCount&)) Like LCase(ComboBox.list(SecondCount&)) And FirstCount& <> SecondCount& Then
                ComboBox.RemoveItem SecondCount&
            End If
        Next SecondCount&
    Next FirstCount&
End Sub

Public Function WindowIsChild(ParentWindow As Long, ChildWindow As Long) As Boolean
    Dim ChildValue As Long
    ChildValue& = IsChild(ParentWindow&, ChildWindow&)
    If ChildValue& <> 0& Then
        WindowIsChild = True
       ElseIf ChildValue& = 0& Then
        WindowIsChild = False
    End If
End Function

Public Function RGBtoHEX(RGB As Long) As String
   Dim HexValue As String, LenHexValue As Long
   HexValue$ = Hex(RGB&)
   LenHexValue& = Len(HexValue$)
   If LenHexValue& = 1 Then HexValue$ = "00000" & HexValue$
   If LenHexValue& = 2 Then HexValue$ = "0000" & HexValue$
   If LenHexValue& = 3 Then HexValue$ = "000" & HexValue$
   If LenHexValue& = 4 Then HexValue$ = "00" & HexValue$
   If LenHexValue& = 5 Then HexValue$ = "0" & HexValue$
   RGBtoHEX$ = "#" & HexValue$
End Function

Public Function RgbToFontColor(RedValue As Long, GreenValue As Long, BlueValue As Long) As String
    RgbToFontColor$ = "<Font Color=#" & Hex(RGB(RedValue&, GreenValue&, BlueValue&)) & ">"
End Function

Public Function RgbToBodyColor(RedValue As Long, GreenValue As Long, BlueValue As Long) As String
    RgbToBodyColor$ = "<Body BgColor=#" & Hex(RGB(RedValue&, GreenValue&, BlueValue&)) & ">"
End Function

Public Function GetClass(WinHandle As Long) As String
    Dim FixedString As String
    FixedString$ = String(250, 0)
    Call GetClassName(WinHandle&, FixedString$, 250)
    GetClass$ = FixedString$
End Function

Public Function SpyHandle() As Long
    Dim CursorPos As POINTAPI
    Call GetCursorPos(CursorPos)
    SpyHandle& = WindowFromPointXy(CursorPos.X, CursorPos.Y)
End Function

Public Function SpyClass() As String
    SpyClass$ = GetClass(SpyHandle&)
End Function

Public Function SpyText() As String
    SpyText$ = GetText(SpyHandle&)
End Function

Public Function SpyStyle() As String
    SpyStyle$ = GetWindowLong(SpyHandle&, (-16))
End Function

Public Function SpyId() As String
    SpyId$ = GetWindowLong(SpyHandle&, (-12))
End Function

Public Function SpyParent() As Long
    SpyParent& = GetParent(SpyHandle&)
End Function

Public Function SpyParentClass() As String
    SpyParentClass$ = GetClass(SpyParent&)
End Function

Public Function SpyParentText() As String
    SpyParentText$ = GetText(SpyParent&)
End Function

Public Function SpyParentStyle() As String
    SpyParentStyle$ = GetWindowLong(SpyParent&, (-16))
End Function

Public Function SpyParentId() As String
    SpyParentId$ = GetWindowLong(SpyParent&, (-12))
End Function

Public Sub RunMenuByString(ParentWindow As Long, StringToGet As String)
    Dim MenuHandle As Long, MenuItemCount As Long, NextItem As Long
    Dim SubMenu As Long, NextMenuItemCount As Long, MenuItemId As Long
    Dim NextNextItem As Long, NextMenuItemId As Long, FixedString As String
    MenuHandle& = GetMenu(ParentWindow&)
    MenuItemCount& = GetMenuItemCount(MenuHandle&)
    For NextItem& = 0& To MenuItemCount& - 1
        SubMenu& = GetSubMenu(MenuHandle&, NextItem&)
        NextMenuItemCount& = GetMenuItemCount(SubMenu&)
        For NextNextItem& = 0& To NextMenuItemCount& - 1
             NextMenuItemId& = GetMenuItemID(SubMenu&, NextNextItem&)
             FixedString$ = String(100, " ")
             Call GetMenuString(SubMenu&, NextMenuItemId&, FixedString$, 100, 1)
             If InStr(LCase(FixedString$), LCase(StringToGet$)) Then
                  Call SendMessageLong(ParentWindow&, WM_COMMAND, NextMenuItemId&, 0&)
                  Exit Sub
             End If
        Next NextNextItem&
    Next NextItem&
End Sub

Public Sub WaitForWindowByTitle(ParentWindow As Long, WindowText As String)
    Dim FindThisWindow As Long
    Do: DoEvents
        FindThisWindow& = Findchildbytitle(ParentWindow&, WindowText$)
    Loop Until FindThisWindow& <> 0&
End Sub

Public Sub WaitForWindowByClass(ParentWindow As Long, ClassWindow As String)
    Dim FindThisWindow As Long
    Do: DoEvents
        FindThisWindow& = FindChildByClass(ParentWindow&, ClassWindow$)
    Loop Until FindThisWindow& <> 0&
End Sub

Public Sub WaitForListToLoad(ListBox As Long)
    Dim FirstCount As Long, SecondCount As Long, ThirdCount As Long
    Dim LastCount As Long
    Do: DoEvents
        FirstCount& = ListCount(ListBox&)
        Yield 0.5
        SecondCount& = ListCount(ListBox&)
        Yield 0.5
        ThirdCount& = ListCount(ListBox&)
    Loop Until FirstCount& <> SecondCount& And ThirdCount& <> FirstCount&
    Yield 0.5
    LastCount& = ListCount(ListBox&)
End Sub

Public Function ListSearch(ListBox As ListBox, SearchString As String) As Boolean
    Dim Search As Long
    On Error Resume Next
    For Search& = 0 To ListCount(ListBox.hwnd) - 1
        If ListBox.list(Search&) = SearchString$ Then
            ListSearch = True
            Form1.List1.RemoveItem (Search)
            Exit Function
        End If
    Next Search&
End Function

Public Function ListSearchTwo(ListBox As ListBox, SearchString As String) As Boolean
    Dim Search As Long
    On Error Resume Next
    For Search& = 0 To ListCount(ListBox.hwnd) - 1
        If InStr(ListBox.list(Search&), SearchString$) <> 0& Then
            ListSearchTwo = True
            Exit Function
        End If
    Next Search&
End Function

Public Function ComboSearch(ComboBox As ComboBox, SearchString As String) As Boolean
    Dim Search As Long
    On Error Resume Next
    For Search& = 0 To ComboCount(ComboBox.hwnd) - 1
        If ComboBox.list(Search&) = SearchString$ Then
            ComboSearch = True
            Exit Function
        End If
    Next Search&
End Function

Public Function StringCount(InThisString As String, FindString As String) As Long
    Dim LenString As Long, Count As Long
    For LenString& = 1 To Len(InThisString$)
        If InStr(LenString&, InThisString$, FindString$) = LenString& Then
            Count& = Count& + 1
        End If
    Next LenString&
    StringCount& = Count&
End Function

Public Function StringSearch(InThisString As String, FindString As String) As Boolean
    Dim LenString As Long, Count As Long
    If InStr(InThisString$, FindString$) <> 0& Then
         StringSearch = True
         Exit Function
        ElseIf InStr(InThisString$, FindString$) = 0& Then
         StringSearch = False
         Exit Function
    End If
End Function

Public Sub ListRemoveNull(ListBox As ListBox)
    Dim Count As Long
    Do: DoEvents
       If TrimSpaces(ListBox.list(Count&)) = "" Then ListBox.RemoveItem (Count&)
       Count& = Count& + 1
    Loop Until Count& >= ListCount(ListBox.hwnd)
End Sub

Public Sub ComboRemoveNull(ComboBox As ComboBox)
    Dim Count As Long
    Do: DoEvents
       If TrimSpaces(ComboBox.list(Count&)) = "" Then ComboBox.RemoveItem (Count&)
       Count& = Count& + 1
    Loop Until Count& >= ComboCount(ComboBox.hwnd)
End Sub

Public Function TrimHtml(TrimThisString As String) As String
    Dim LenString As Long, Instance As Long, NewMain As String
    Dim Instance2 As Long
    NewMain$ = ReplaceCharacters(TrimThisString$, "<B>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</B>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "<S>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</S>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "<I>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</I>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "<U>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</U>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "<SUB>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</SUB>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</SUP>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "<HTML>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</HTML>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "<FONT>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "</FONT>", "")
    NewMain$ = ReplaceCharacters(NewMain$, "<BR>", "")
    If InStr(NewMain$, "<FONT COLOR=") <> 0& Then
        Do: DoEvents
            If Right("<FONT COLOR=", Len("<FONT COLOR=") + 10) = "" Then
                NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT COLOR=") - 1)
               ElseIf Right("<FONT COLOR=", Len("<FONT COLOR=") + 10) <> "" Then
                NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT COLOR=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<FONT COLOR=") - 21)
            End If
       Loop Until InStr(NewMain$, "<FONT COLOR=") = 0&
    End If
    If InStr(NewMain$, "<FONT SIZE=") <> 0& Then
        Do: DoEvents
            If Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) = ">" Then
                If Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) = "" Then
                    NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1)
                   ElseIf Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) <> "" Then
                    NewMain$ = ReplaceCharacters(Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<FONT SIZE=") - 13), ">", "")
                End If
               ElseIf Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) <> ">" Then
                If Right("<FONT SIZE=", Len("<FONT SIZE=") + 3) = "" Then
                    NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1)
                   ElseIf Right("<FONT SIZE=", Len("<FONT SIZE=") + 3) <> "" Then
                    NewMain$ = ReplaceCharacters(Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<FONT SIZE=") - 12), ">", "")
                End If
            End If
        Loop Until InStr(NewMain$, "<FONT SIZE=") = 0&
   End If
   If InStr(NewMain$, "<BODY BGCOLOR=") <> 0& Then
      Do: DoEvents
         If Right("<BODY BGCOLOR=", Len("<BODY BGCOLOR=") + 12) = "" Then
             NewMain$ = Left(NewMain$, InStr(NewMain$, "<BODY BGCOLOR=") - 1)
            ElseIf Right("<BODY BGCOLOR=", Len("<BODY BGCOLOR=") + 12) <> "" Then
             NewMain$ = Left(NewMain$, InStr(NewMain$, "<BODY BGCOLOR=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<BODY BGCOLOR=") - 23)
         End If
      Loop Until InStr(NewMain$, "<BODY BGCOLOR=") = 0&
   End If
   TrimHtml$ = NewMain$
End Function

Public Function ModuleLoadList(ListBox As Control, Optional EmphasisOnModuleName As Boolean = False) As String
    Dim VbWindow As Long, VbMdi As Long, VbaWindow As Long
    Dim VbCombo1 As Long, VbCombo2 As Long, VbComboCount As Long
    Dim Index As Long, VbComboItem As String * 256, VbComboL As Long
    Dim VbComboDrop As Long
    VbWindow& = FindWindow("wndclass_desked_gsk", vbNullString)
    VbMdi& = FindWindowEx(VbWindow&, 0&, "MDIClient", vbNullString)
    VbaWindow& = FindWindowEx(VbMdi&, 0&, "VbaWindow", vbNullString)
    VbCombo1& = FindWindowEx(VbaWindow&, 0&, "ComboBox", vbNullString)
    VbCombo2& = FindWindowEx(VbaWindow&, VbCombo1&, "ComboBox", vbNullString)
    If InStr(GetCaption(VbWindow&), "Form") = 0& Then
        Call PostMessage(VbCombo2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(VbCombo2&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            VbComboDrop& = FindWindow("#32769", vbNullString)
            VbComboL& = FindWindowEx(VbComboDrop&, 0&, "ComboLBox", vbNullString)
        Loop Until IsWindowVisible(VbComboL&) <> 0&
        VbComboCount& = ComboCount(VbCombo2&)
        For Index& = 0 To VbComboCount& - 1
            DoEvents
            Call SendMessageByString(VbComboL&, LB_GETTEXT, Index&, VbComboItem$)
            ListBox.AddItem VbComboItem$
        Next Index&
        Call PostMessage(VbCombo2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(VbCombo2&, WM_LBUTTONUP, 0&, 0&)
        ListBox.RemoveItem 0
        If EmphasisOnModuleName = True Then ModuleLoadList$ = "</S></B></U></I>" & VbComboCount& - 1 & " subs/functions extracted from <B><U>" & ReplaceCharacters(GetInstance(GetCaption(VbWindow&), "[", 3), " (Code)]", "</B></U> module.")
        If EmphasisOnModuleName = False Then ModuleLoadList$ = VbComboCount& - 1 & " subs/functions extracted from " & ReplaceCharacters(GetInstance(GetCaption(VbWindow&), "[", 3), " (Code)]", " module.")
       ElseIf InStr(GetCaption(VbWindow&), "Form") <> 0& Then
        ModuleLoadList$ = "No module could be found."
        Exit Function
    End If
End Function

Public Function ModuleFindFromList(ListBox As Control, SearchThisString As String) As String
    Dim SubString As String, FunctionString As String, DeclareString As String
    Dim PublicSubString As String, PrivateSubString As String
    Dim PublicFunctionString As String, PrivateFunctionString As String
    Dim EndSubString As String, EndFunctionString As String
    SubString$ = "Sub "
     FunctionString$ = "Function "
      PublicSubString$ = "Public Sub "
       PrivateSubString$ = "Private Sub "
        PublicFunctionString$ = "Public Function "
       PrivateFunctionString$ = "Private Function "
      EndSubString$ = "End Sub"
     EndFunctionString$ = "End Function"
    DeclareString$ = "Declare "
   If InStr(SearchThisString$, PublicSubString$ & ListBox.Text) <> 0& Then
        ModuleFindFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicSubString$ & ListBox.Text))
        ModuleFindFromList$ = ModuleFindFromList$ & Mid(ModuleFindFromList$, Len(ModuleFindFromList$), InStr(SearchThisString$, EndSubString$))
        ModuleFindFromList$ = Left(ModuleFindFromList$, InStr(ModuleFindFromList$, vbCrLf & EndSubString$) + 8)
        Exit Function
       ElseIf InStr(SearchThisString$, DeclareString$ & SubString$ & ListBox.Text) <> 0& Or InStr(SearchThisString$, DeclareString$ & FunctionString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = "No matches were found."
            Exit Function
       ElseIf InStr(SearchThisString$, "Public " & DeclareString$ & SubString$ & ListBox.Text) <> 0& Or InStr(SearchThisString$, "Public " & DeclareString$ & FunctionString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = "No matches were found."
            Exit Function
       ElseIf InStr(SearchThisString$, "Private " & DeclareString$ & SubString$ & ListBox.Text) <> 0& Or InStr(SearchThisString$, "Private " & DeclareString$ & FunctionString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = "No matches were found."
            Exit Function
       ElseIf InStr(SearchThisString$, PublicFunctionString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicFunctionString$ & ListBox.Text))
            ModuleFindFromList$ = ModuleFindFromList$ & Mid(ModuleFindFromList$, Len(ModuleFindFromList$), InStr(SearchThisString$, EndFunctionString$))
            ModuleFindFromList$ = Left(ModuleFindFromList$, InStr(ModuleFindFromList$, vbCrLf & EndFunctionString$) + 13)
            Exit Function
       ElseIf InStr(SearchThisString$, PrivateSubString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateSubString$ & ListBox.Text))
            ModuleFindFromList$ = ModuleFindFromList$ & Mid(ModuleFindFromList$, Len(ModuleFindFromList$), InStr(SearchThisString$, EndSubString$))
            ModuleFindFromList$ = Left(ModuleFindFromList$, InStr(ModuleFindFromList$, vbCrLf & EndSubString$) + 8)
            Exit Function
       ElseIf InStr(SearchThisString$, SubString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, SubString$ & ListBox.Text))
            ModuleFindFromList$ = ModuleFindFromList$ & Mid(ModuleFindFromList$, Len(ModuleFindFromList$), InStr(SearchThisString$, EndSubString$))
            ModuleFindFromList$ = Left(ModuleFindFromList$, InStr(ModuleFindFromList$, vbCrLf & EndSubString$) + 8)
            Exit Function
       ElseIf InStr(SearchThisString$, PrivateFunctionString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateFunctionString$ & ListBox.Text))
            ModuleFindFromList$ = ModuleFindFromList$ & Mid(ModuleFindFromList$, Len(ModuleFindFromList$), InStr(SearchThisString$, EndFunctionString$))
            ModuleFindFromList$ = Left(ModuleFindFromList$, InStr(ModuleFindFromList$, vbCrLf & EndFunctionString$) + 13)
            Exit Function
       ElseIf InStr(SearchThisString$, FunctionString$ & ListBox.Text) <> 0& Then
            ModuleFindFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, FunctionString$ & ListBox.Text))
            ModuleFindFromList$ = ModuleFindFromList$ & Mid(ModuleFindFromList$, Len(ModuleFindFromList$), InStr(SearchThisString$, EndFunctionString$))
            ModuleFindFromList$ = Left(ModuleFindFromList$, InStr(ModuleFindFromList$, vbCrLf & EndFunctionString$) + 13)
            Exit Function
       Else
            ModuleFindFromList$ = "No matches were found."
            Exit Function
   End If
End Function

Public Function ModuleFindFromListIndex(ListBox As Control, ListIndex As Long, SearchThisString As String) As String
    Dim SubString As String, FunctionString As String, DeclareString As String
    Dim PublicSubString As String, PrivateSubString As String
    Dim PublicFunctionString As String, PrivateFunctionString As String
    Dim EndSubString As String, EndFunctionString As String
    SubString$ = "Sub "
     FunctionString$ = "Function "
      PublicSubString$ = "Public Sub "
       PrivateSubString$ = "Private Sub "
        PublicFunctionString$ = "Public Function "
       PrivateFunctionString$ = "Private Function "
      EndSubString$ = "End Sub"
     EndFunctionString$ = "End Function"
    DeclareString$ = "Declare "
   If InStr(SearchThisString$, PublicSubString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleFindFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicSubString$ & ListBox.list(ListIndex&)))
        ModuleFindFromListIndex$ = ModuleFindFromListIndex$ & Mid(ModuleFindFromListIndex$, Len(ModuleFindFromListIndex$), InStr(SearchThisString$, EndSubString$))
        ModuleFindFromListIndex$ = Left(ModuleFindFromListIndex$, InStr(ModuleFindFromListIndex$, vbCrLf & EndSubString$) + 8)
        Exit Function
       ElseIf InStr(SearchThisString$, DeclareString$ & SubString$ & ListBox.list(ListIndex&)) <> 0& Or InStr(SearchThisString$, DeclareString$ & FunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = "No matches were found."
            Exit Function
       ElseIf InStr(SearchThisString$, "Public " & DeclareString$ & SubString$ & ListBox.list(ListIndex&)) <> 0& Or InStr(SearchThisString$, "Public " & DeclareString$ & FunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = "No matches were found."
            Exit Function
       ElseIf InStr(SearchThisString$, "Private " & DeclareString$ & SubString$ & ListBox.list(ListIndex&)) <> 0& Or InStr(SearchThisString$, "Private " & DeclareString$ & FunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = "No matches were found."
            Exit Function
       ElseIf InStr(SearchThisString$, PublicFunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicFunctionString$ & ListBox.list(ListIndex&)))
            ModuleFindFromListIndex$ = ModuleFindFromListIndex$ & Mid(ModuleFindFromListIndex$, Len(ModuleFindFromListIndex$), InStr(SearchThisString$, EndFunctionString$))
            ModuleFindFromListIndex$ = Left(ModuleFindFromListIndex$, InStr(ModuleFindFromListIndex$, vbCrLf & EndFunctionString$) + 13)
            Exit Function
       ElseIf InStr(SearchThisString$, PrivateSubString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateSubString$ & ListBox.list(ListIndex&)))
            ModuleFindFromListIndex$ = ModuleFindFromListIndex$ & Mid(ModuleFindFromListIndex$, Len(ModuleFindFromListIndex$), InStr(SearchThisString$, EndSubString$))
            ModuleFindFromListIndex$ = Left(ModuleFindFromListIndex$, InStr(ModuleFindFromListIndex$, vbCrLf & EndSubString$) + 8)
            Exit Function
       ElseIf InStr(SearchThisString$, SubString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, SubString$ & ListBox.list(ListIndex&)))
            ModuleFindFromListIndex$ = ModuleFindFromListIndex$ & Mid(ModuleFindFromListIndex$, Len(ModuleFindFromListIndex$), InStr(SearchThisString$, EndSubString$))
            ModuleFindFromListIndex$ = Left(ModuleFindFromListIndex$, InStr(ModuleFindFromListIndex$, vbCrLf & EndSubString$) + 8)
            Exit Function
       ElseIf InStr(SearchThisString$, PrivateFunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateFunctionString$ & ListBox.list(ListIndex&)))
            ModuleFindFromListIndex$ = ModuleFindFromListIndex$ & Mid(ModuleFindFromListIndex$, Len(ModuleFindFromListIndex$), InStr(SearchThisString$, EndFunctionString$))
            ModuleFindFromListIndex$ = Left(ModuleFindFromListIndex$, InStr(ModuleFindFromListIndex$, vbCrLf & EndFunctionString$) + 13)
            Exit Function
       ElseIf InStr(SearchThisString$, FunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
            ModuleFindFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, FunctionString$ & ListBox.list(ListIndex&)))
            ModuleFindFromListIndex$ = ModuleFindFromListIndex$ & Mid(ModuleFindFromListIndex$, Len(ModuleFindFromListIndex$), InStr(SearchThisString$, EndFunctionString$))
            ModuleFindFromListIndex$ = Left(ModuleFindFromListIndex$, InStr(ModuleFindFromListIndex$, vbCrLf & EndFunctionString$) + 13)
            Exit Function
       Else
            ModuleFindFromListIndex$ = "No matches were found."
            Exit Function
   End If
End Function

Public Function ModuleSubOrFunctionTitle(FromThisString As String) As String
    Dim SubString As String, FunctionString As String, DeclareString As String
    Dim PublicSubString As String, PrivateSubString As String, prepstring As String
    Dim PublicFunctionString As String, PrivateFunctionString As String
    SubString$ = "Sub "
     FunctionString$ = "Function "
      PublicSubString$ = "Public Sub "
       PrivateSubString$ = "Private Sub "
      PublicFunctionString$ = "Public Function "
    PrivateFunctionString$ = "Private Function "
    DeclareString$ = "Declare "
   If InStr(FromThisString$, PublicSubString$) <> 0& Then
       prepstring$ = Mid(FromThisString$, Len(PublicSubString$) + 1)
       ModuleSubOrFunctionTitle$ = Left(prepstring$, InStr(prepstring$, "(") - 1)
       Exit Function
      ElseIf InStr(FromThisString$, PrivateSubString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PrivateSubString$) + 1)
           ModuleSubOrFunctionTitle$ = Left(prepstring$, InStr(prepstring$, "(") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, PublicFunctionString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PublicFunctionString$) + 1)
           ModuleSubOrFunctionTitle$ = Left(prepstring$, InStr(prepstring$, "(") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, PrivateFunctionString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PrivateFunctionString$) + 1)
           ModuleSubOrFunctionTitle$ = Left(prepstring$, InStr(prepstring$, "(") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, "Public " & DeclareString$ & FunctionString$) <> 0& Or InStr(FromThisString$, "Public " & DeclareString$ & SubString$) <> 0& Then
           ModuleSubOrFunctionTitle$ = "No sub or function was found."
           Exit Function
      ElseIf InStr(FromThisString$, "Private " & DeclareString$ & FunctionString$) <> 0& Or InStr(FromThisString$, "Private " & DeclareString$ & SubString$) <> 0& Then
           ModuleSubOrFunctionTitle$ = "No sub or function was found."
           Exit Function
      ElseIf InStr(FromThisString$, DeclareString$ & FunctionString$) <> 0& Or InStr(FromThisString$, DeclareString$ & SubString$) <> 0& Then
           ModuleSubOrFunctionTitle$ = "No sub or function was found."
           Exit Function
      ElseIf InStr(FromThisString$, SubString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(SubString$) + 1)
           ModuleSubOrFunctionTitle$ = Left(prepstring$, InStr(prepstring$, "(") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, FunctionString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(FunctionString$) + 1)
           ModuleSubOrFunctionTitle$ = Left(prepstring$, InStr(prepstring$, "(") - 1)
           Exit Function
      Else
           ModuleSubOrFunctionTitle$ = "No sub or function was found."
           Exit Function
   End If
End Function

Public Sub ModuleSubsAndFunctions(ListBox As Control, SearchThisString As String)
    Dim CountLines As Long, NextLine As Long, prepstring As String
    CountLines& = LineCount(SearchThisString$)
   On Error Resume Next
   If CountLines& <= 0& Then Exit Sub
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleSubOrFunctionTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No sub or function was found." And InStr(prepstring$, "=") = 0& Then
               ListBox.AddItem prepstring$
           End If
   Next NextLine&
End Sub

Public Function ModuleDecFromList(ListBox As Control, SearchThisString As String) As String
    Dim DecSubString As String, DecFunctionString As String, DeclareString As String
    Dim PublicDecSubString As String, PrivateDecSubString As String
    Dim PublicDecFunctionString As String, PrivateDecFunctionString As String
    DecSubString$ = "Declare Sub "
     DecFunctionString$ = "Declare Function "
      PublicDecSubString$ = "Public Declare Sub "
       PrivateDecSubString$ = "Private Declare Sub "
      PublicDecFunctionString$ = "Public Declare Function "
    PrivateDecFunctionString$ = "Private Declare Function "
   If InStr(SearchThisString$, PublicDecSubString$ & ListBox.Text) <> 0& Then
        ModuleDecFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicDecSubString$ & ListBox.Text))
        ModuleDecFromList$ = Left(ModuleDecFromList$, InStr(ModuleDecFromList$, vbNewLine))
        ModuleDecFromList$ = ReplaceCharacters(ModuleDecFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateDecSubString$ & ListBox.Text) <> 0& Then
        ModuleDecFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateDecSubString$ & ListBox.Text))
        ModuleDecFromList$ = Left(ModuleDecFromList$, InStr(ModuleDecFromList$, vbNewLine))
        ModuleDecFromList$ = ReplaceCharacters(ModuleDecFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PublicDecFunctionString$ & ListBox.Text) <> 0& Then
        ModuleDecFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicDecFunctionString$ & ListBox.Text))
        ModuleDecFromList$ = Left(ModuleDecFromList$, InStr(ModuleDecFromList$, vbNewLine))
        ModuleDecFromList$ = ReplaceCharacters(ModuleDecFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateDecFunctionString$ & ListBox.Text) <> 0& Then
        ModuleDecFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateDecFunctionString$ & ListBox.Text))
        ModuleDecFromList$ = Left(ModuleDecFromList$, InStr(ModuleDecFromList$, vbNewLine))
        ModuleDecFromList$ = ReplaceCharacters(ModuleDecFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, DecSubString$ & ListBox.Text) <> 0& Then
        ModuleDecFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, DecSubString$ & ListBox.Text))
        ModuleDecFromList$ = Left(ModuleDecFromList$, InStr(ModuleDecFromList$, vbNewLine))
        ModuleDecFromList$ = ReplaceCharacters(ModuleDecFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, DecFunctionString$ & ListBox.Text) <> 0& Then
        ModuleDecFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, DecFunctionString$ & ListBox.Text))
        ModuleDecFromList$ = Left(ModuleDecFromList$, InStr(ModuleDecFromList$, vbNewLine))
        ModuleDecFromList$ = ReplaceCharacters(ModuleDecFromList$, vbCr, "")
        Exit Function
       Else
        ModuleDecFromList$ = "No matches were found."
        Exit Function
   End If
End Function

Public Function ModuleDecFromListIndex(ListBox As Control, ListIndex As Long, SearchThisString As String) As String
    Dim DecSubString As String, DecFunctionString As String, DeclareString As String
    Dim PublicDecSubString As String, PrivateDecSubString As String
    Dim PublicDecFunctionString As String, PrivateDecFunctionString As String
    DecSubString$ = "Declare Sub "
     DecFunctionString$ = "Declare Function "
      PublicDecSubString$ = "Public Declare Sub "
       PrivateDecSubString$ = "Private Declare Sub "
      PublicDecFunctionString$ = "Public Declare Function "
    PrivateDecFunctionString$ = "Private Declare Function "
   If InStr(SearchThisString$, PublicDecSubString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleDecFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicDecSubString$ & ListBox.list(ListIndex&)))
        ModuleDecFromListIndex$ = Left(ModuleDecFromListIndex$, InStr(ModuleDecFromListIndex$, vbNewLine))
        ModuleDecFromListIndex$ = ReplaceCharacters(ModuleDecFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateDecSubString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleDecFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateDecSubString$ & ListBox.list(ListIndex&)))
        ModuleDecFromListIndex$ = Left(ModuleDecFromListIndex$, InStr(ModuleDecFromListIndex$, vbNewLine))
        ModuleDecFromListIndex$ = ReplaceCharacters(ModuleDecFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PublicDecFunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleDecFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicDecFunctionString$ & ListBox.list(ListIndex&)))
        ModuleDecFromListIndex$ = Left(ModuleDecFromListIndex$, InStr(ModuleDecFromListIndex$, vbNewLine))
        ModuleDecFromListIndex$ = ReplaceCharacters(ModuleDecFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateDecFunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleDecFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateDecFunctionString$ & ListBox.list(ListIndex&)))
        ModuleDecFromListIndex$ = Left(ModuleDecFromListIndex$, InStr(ModuleDecFromListIndex$, vbNewLine))
        ModuleDecFromListIndex$ = ReplaceCharacters(ModuleDecFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, DecSubString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleDecFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, DecSubString$ & ListBox.list(ListIndex&)))
        ModuleDecFromListIndex$ = Left(ModuleDecFromListIndex$, InStr(ModuleDecFromListIndex$, vbNewLine))
        ModuleDecFromListIndex$ = ReplaceCharacters(ModuleDecFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, DecFunctionString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleDecFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, DecFunctionString$ & ListBox.list(ListIndex&)))
        ModuleDecFromListIndex$ = Left(ModuleDecFromListIndex$, InStr(ModuleDecFromListIndex$, vbNewLine))
        ModuleDecFromListIndex$ = ReplaceCharacters(ModuleDecFromListIndex$, vbCr, "")
        Exit Function
       Else
        ModuleDecFromListIndex$ = "No matches were found."
        Exit Function
   End If
End Function

Public Function ModuleDecTitle(FromThisString As String) As String
    Dim DecSubString As String, DecFunctionString As String, DeclareString As String
    Dim PublicDecSubString As String, PrivateDecSubString As String, prepstring As String
    Dim PublicDecFunctionString As String, PrivateDecFunctionString As String
    DecSubString$ = "Declare Sub "
     DecFunctionString$ = "Declare Function "
      PublicDecSubString$ = "Public Declare Sub "
       PrivateDecSubString$ = "Private Declare Sub "
      PublicDecFunctionString$ = "Public Declare Function "
    PrivateDecFunctionString$ = "Private Declare Function "
   If InStr(FromThisString$, PublicDecFunctionString$) <> 0& Then
       prepstring$ = Mid(FromThisString$, Len(PublicDecFunctionString$) + 1)
       ModuleDecTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
       Exit Function
      ElseIf InStr(FromThisString$, PrivateDecFunctionString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PrivateDecFunctionString$) + 1)
           ModuleDecTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, PublicDecSubString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PublicDecSubString$) + 1)
           ModuleDecTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, PrivateDecSubString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PrivateDecSubString$) + 1)
           ModuleDecTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, DecSubString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(DecSubString$) + 1)
           ModuleDecTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, DecFunctionString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(DecFunctionString$) + 1)
           ModuleDecTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      Else
           ModuleDecTitle$ = "No declaration was found."
           Exit Function
   End If
End Function

Public Sub ModuleDeclarations(ListBox As Control, SearchThisString As String)
    Dim CountLines As Long, NextLine As Long, prepstring As String
    CountLines& = LineCount(SearchThisString$)
   On Error Resume Next
   If CountLines& <= 0& Then Exit Sub
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleDecTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No declaration was found." And InStr(prepstring$, "=") = 0& Then
               ListBox.AddItem prepstring$
           End If
   Next NextLine&
End Sub

Public Function ModuleConstFromList(ListBox As Control, SearchThisString As String) As String
    Dim GlobalConString As String, PublicConString As String, PrivateConString As String
    Dim ConString As String, prepstring As String
    ConString$ = "Const "
     GlobalConString$ = "Global Const "
     PublicConString$ = "Public Const "
    PrivateConString$ = "Private Const "
   If InStr(SearchThisString$, GlobalConString$ & ListBox.Text) <> 0& Then
        ModuleConstFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, GlobalConString$ & ListBox.Text))
        ModuleConstFromList$ = Left(ModuleConstFromList$, InStr(ModuleConstFromList$, vbNewLine))
        ModuleConstFromList$ = ReplaceCharacters(ModuleConstFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PublicConString$ & ListBox.Text) <> 0& Then
        ModuleConstFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicConString$ & ListBox.Text))
        ModuleConstFromList$ = Left(ModuleConstFromList$, InStr(ModuleConstFromList$, vbNewLine))
        ModuleConstFromList$ = ReplaceCharacters(ModuleConstFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateConString$ & ListBox.Text) <> 0& Then
        ModuleConstFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateConString$ & ListBox.Text))
        ModuleConstFromList$ = Left(ModuleConstFromList$, InStr(ModuleConstFromList$, vbNewLine))
        ModuleConstFromList$ = ReplaceCharacters(ModuleConstFromList$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, ConString$ & ListBox.Text) <> 0& Then
        ModuleConstFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, ConString$ & ListBox.Text))
        ModuleConstFromList$ = Left(ModuleConstFromList$, InStr(ModuleConstFromList$, vbNewLine))
        ModuleConstFromList$ = ReplaceCharacters(ModuleConstFromList$, vbCr, "")
        Exit Function
       Else
        ModuleConstFromList$ = "No matches were found."
   End If
End Function

Public Function ModuleConstFromListIndex(ListBox As Control, ListIndex As Long, SearchThisString As String) As String
    Dim GlobalConString As String, PublicConString As String, PrivateConString As String
    Dim ConString As String, prepstring As String
    ConString$ = "Const "
     GlobalConString$ = "Global Const "
     PublicConString$ = "Public Const "
    PrivateConString$ = "Private Const "
   If InStr(SearchThisString$, GlobalConString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleConstFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, GlobalConString$ & ListBox.list(ListIndex&)))
        ModuleConstFromListIndex$ = Left(ModuleConstFromListIndex$, InStr(ModuleConstFromListIndex$, vbNewLine))
        ModuleConstFromListIndex$ = ReplaceCharacters(ModuleConstFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PublicConString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleConstFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicConString$ & ListBox.list(ListIndex&)))
        ModuleConstFromListIndex$ = Left(ModuleConstFromListIndex$, InStr(ModuleConstFromListIndex$, vbNewLine))
        ModuleConstFromListIndex$ = ReplaceCharacters(ModuleConstFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateConString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleConstFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateConString$ & ListBox.list(ListIndex&)))
        ModuleConstFromListIndex$ = Left(ModuleConstFromListIndex$, InStr(ModuleConstFromListIndex$, vbNewLine))
        ModuleConstFromListIndex$ = ReplaceCharacters(ModuleConstFromListIndex$, vbCr, "")
        Exit Function
       ElseIf InStr(SearchThisString$, ConString$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleConstFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, ConString$ & ListBox.list(ListIndex&)))
        ModuleConstFromListIndex$ = Left(ModuleConstFromListIndex$, InStr(ModuleConstFromListIndex$, vbNewLine))
        ModuleConstFromListIndex$ = ReplaceCharacters(ModuleConstFromListIndex$, vbCr, "")
        Exit Function
       Else
        ModuleConstFromListIndex$ = "No matches were found."
   End If
End Function

Public Function ModuleConstTitle(FromThisString As String) As String
    Dim GlobalConString As String, PublicConString As String, PrivateConString As String
    Dim ConString As String, prepstring As String
    ConString$ = "Const "
     GlobalConString$ = "Global Const "
     PublicConString$ = "Public Const "
    PrivateConString$ = "Private Const "
   If InStr(FromThisString$, GlobalConString$) <> 0& Then
       prepstring$ = Mid(FromThisString$, Len(GlobalConString$) + 1)
       ModuleConstTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
       Exit Function
      ElseIf InStr(FromThisString$, PublicConString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PublicConString$) + 1)
           ModuleConstTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, PrivateConString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PrivateConString$) + 1)
           ModuleConstTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      ElseIf InStr(FromThisString$, ConString$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(ConString$) + 1)
           ModuleConstTitle$ = Left(prepstring$, InStr(prepstring$, " ") - 1)
           Exit Function
      Else
           ModuleConstTitle$ = "No constant was found."
           Exit Function
   End If
End Function

Public Sub ModuleConstants(ListBox As Control, SearchThisString As String)
    Dim CountLines As Long, NextLine As Long, prepstring As String
    CountLines& = LineCount(SearchThisString$)
   On Error Resume Next
   If CountLines& <= 0& Then Exit Sub
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleConstTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No constant was found." And InStr(prepstring$, "=") = 0& Then
               ListBox.AddItem prepstring$
           End If
   Next NextLine&
End Sub

Public Function ModuleTypeFromList(ListBox As Control, SearchThisString As String) As String
    Dim GlobalType As String, PublicType As String, PrivateType As String
    Dim EndType As String
    GlobalType$ = "Global Type "
     PublicType$ = "Public Type "
     PrivateType$ = "Private Type "
    EndType$ = "End Type"
   If InStr(SearchThisString$, GlobalType$ & ListBox.Text) <> 0& Then
        ModuleTypeFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, GlobalType$ & ListBox.Text))
        ModuleTypeFromList$ = ModuleTypeFromList$ & Mid(ModuleTypeFromList$, Len(ModuleTypeFromList$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromList$ = Left(ModuleTypeFromList$, InStr(ModuleTypeFromList$, vbCrLf & EndType$) + 9)
        Exit Function
       ElseIf InStr(SearchThisString$, PublicType$ & ListBox.Text) <> 0& Then
        ModuleTypeFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicType$ & ListBox.Text))
        ModuleTypeFromList$ = ModuleTypeFromList$ & Mid(ModuleTypeFromList$, Len(ModuleTypeFromList$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromList$ = Left(ModuleTypeFromList$, InStr(ModuleTypeFromList$, vbCrLf & EndType$) + 9)
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateType$ & ListBox.Text) <> 0& Then
        ModuleTypeFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateType$ & ListBox.Text))
        ModuleTypeFromList$ = ModuleTypeFromList$ & Mid(ModuleTypeFromList$, Len(ModuleTypeFromList$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromList$ = Left(ModuleTypeFromList$, InStr(ModuleTypeFromList$, vbCrLf & EndType$) + 9)
        Exit Function
       ElseIf InStr(SearchThisString$, "Type " & ListBox.Text) <> 0& Then
        ModuleTypeFromList$ = Mid(SearchThisString$, InStr(SearchThisString$, "Type " & ListBox.Text))
        ModuleTypeFromList$ = ModuleTypeFromList$ & Mid(ModuleTypeFromList$, Len(ModuleTypeFromList$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromList$ = Left(ModuleTypeFromList$, InStr(ModuleTypeFromList$, vbCrLf & EndType$) + 9)
        Exit Function
       Else
        ModuleTypeFromList$ = "No matches were found."
        Exit Function
   End If
End Function

Public Function ModuleTypeFromListIndex(ListBox As Control, ListIndex As Long, SearchThisString As String) As String
    Dim GlobalType As String, PublicType As String, PrivateType As String
    Dim EndType As String
    GlobalType$ = "Global Type "
     PublicType$ = "Public Type "
     PrivateType$ = "Private Type "
    EndType$ = "End Type"
   If InStr(SearchThisString$, GlobalType$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleTypeFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, GlobalType$ & ListBox.list(ListIndex&)))
        ModuleTypeFromListIndex$ = ModuleTypeFromListIndex$ & Mid(ModuleTypeFromListIndex$, Len(ModuleTypeFromListIndex$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromListIndex$ = Left(ModuleTypeFromListIndex$, InStr(ModuleTypeFromListIndex$, vbCrLf & EndType$) + 9)
        Exit Function
       ElseIf InStr(SearchThisString$, PublicType$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleTypeFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PublicType$ & ListBox.list(ListIndex&)))
        ModuleTypeFromListIndex$ = ModuleTypeFromListIndex$ & Mid(ModuleTypeFromListIndex$, Len(ModuleTypeFromListIndex$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromListIndex$ = Left(ModuleTypeFromListIndex$, InStr(ModuleTypeFromListIndex$, vbCrLf & EndType$) + 9)
        Exit Function
       ElseIf InStr(SearchThisString$, PrivateType$ & ListBox.list(ListIndex&)) <> 0& Then
        ModuleTypeFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, PrivateType$ & ListBox.list(ListIndex&)))
        ModuleTypeFromListIndex$ = ModuleTypeFromListIndex$ & Mid(ModuleTypeFromListIndex$, Len(ModuleTypeFromListIndex$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromListIndex$ = Left(ModuleTypeFromListIndex$, InStr(ModuleTypeFromListIndex$, vbCrLf & EndType$) + 9)
        Exit Function
       ElseIf InStr(SearchThisString$, "Type " & ListBox.list(ListIndex&)) <> 0& Then
        ModuleTypeFromListIndex$ = Mid(SearchThisString$, InStr(SearchThisString$, "Type " & ListBox.list(ListIndex&)))
        ModuleTypeFromListIndex$ = ModuleTypeFromListIndex$ & Mid(ModuleTypeFromListIndex$, Len(ModuleTypeFromListIndex$), InStr(SearchThisString$, EndType$))
        ModuleTypeFromListIndex$ = Left(ModuleTypeFromListIndex$, InStr(ModuleTypeFromListIndex$, vbCrLf & EndType$) + 9)
        Exit Function
       Else
        ModuleTypeFromListIndex$ = "No matches were found."
        Exit Function
   End If
End Function

Public Function ModuleTypeTitle(FromThisString As String) As String
    Dim GlobalType As String, PublicType As String, PrivateType As String
    Dim EndType As String, prepstring As String
    GlobalType$ = "Global Type "
     PublicType$ = "Public Type "
     PrivateType$ = "Private Type "
   If InStr(FromThisString$, GlobalType$) <> 0& Then
       prepstring$ = Mid(FromThisString$, Len(GlobalType$) + 1)
       ModuleTypeTitle$ = Left(prepstring$, Len(Right(GlobalType$, Len(GlobalType$))))
       Exit Function
      ElseIf InStr(FromThisString$, PublicType$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PublicType$) + 1)
           ModuleTypeTitle$ = Left(prepstring$, Len(Right(PublicType$, Len(PublicType$))))
           Exit Function
      ElseIf InStr(FromThisString$, PrivateType$) <> 0& Then
           prepstring$ = Mid(FromThisString$, Len(PrivateType$) + 1)
           ModuleTypeTitle$ = Left(prepstring$, Len(Right(PrivateType$, Len(PrivateType$))))
           Exit Function
      ElseIf InStr(FromThisString$, "Type ") <> 0& Then
           prepstring$ = Mid(FromThisString$, Len("Type ") + 1)
           ModuleTypeTitle$ = Left(prepstring$, Len(Right("Type ", Len("Type "))))
           Exit Function
      Else
           ModuleTypeTitle$ = "No type was found."
           Exit Function
   End If
End Function

Public Sub ModuleTypes(ListBox As Control, SearchThisString As String)
    Dim CountLines As Long, NextLine As Long, prepstring As String
    CountLines& = LineCount(SearchThisString$)
   On Error Resume Next
   If CountLines& <= 0& Then Exit Sub
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleTypeTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No type was found." Then
               ListBox.AddItem prepstring$
           End If
   Next NextLine&
End Sub

Public Function ModuleSubAndFunctionCount(SearchThisString As String) As Long
    Dim CountLines As Long, NextLine As Long, prepstring As String
    Dim PrepCount As Long
    CountLines& = LineCount(SearchThisString$)
    PrepCount& = 0&
   On Error Resume Next
   If CountLines& <= 0& Then Exit Function
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleSubOrFunctionTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No sub or function was found." And InStr(prepstring$, "=") = 0& Then
               PrepCount& = PrepCount& + 1
           End If
   Next NextLine&
    ModuleSubAndFunctionCount& = PrepCount&
End Function

Public Function ModuleDeclarationCount(SearchThisString As String) As Long
    Dim CountLines As Long, NextLine As Long, prepstring As String
    Dim PrepCount As Long
    CountLines& = LineCount(SearchThisString$)
    PrepCount& = 0&
   On Error Resume Next
   If CountLines& <= 0& Then Exit Function
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleDecTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No declaration was found." And InStr(prepstring$, "=") = 0& Then
               PrepCount& = PrepCount& + 1
           End If
   Next NextLine&
    ModuleDeclarationCount& = PrepCount&
End Function

Public Function ModuleConstantCount(SearchThisString As String) As Long
    Dim CountLines As Long, NextLine As Long, prepstring As String
    Dim PrepCount As Long
    CountLines& = LineCount(SearchThisString$)
    PrepCount& = 0&
   On Error Resume Next
   If CountLines& <= 0& Then Exit Function
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleConstTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No constant was found." And InStr(prepstring$, "=") = 0& Then
               PrepCount& = PrepCount& + 1
           End If
   Next NextLine&
    ModuleConstantCount& = PrepCount&
End Function

Public Function ModuleTypeCount(SearchThisString As String) As Long
    Dim CountLines As Long, NextLine As Long, prepstring As String
    Dim PrepCount As Long
    CountLines& = LineCount(SearchThisString$)
    PrepCount& = 0&
   On Error Resume Next
   If CountLines& <= 0& Then Exit Function
   For NextLine& = 1 To CountLines&
       prepstring$ = ModuleTypeTitle(LineFromString(SearchThisString$, NextLine&))
           If prepstring$ <> "No type was found." And InStr(prepstring$, "=") = 0& Then
               PrepCount& = PrepCount& + 1
           End If
   Next NextLine&
    ModuleTypeCount& = PrepCount&
End Function

Public Function LineFromString(MyString As String, line As Long) As String
    'Thanks to DoS for this one
    Dim theline As String, Count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    Count& = LineCount(MyString$)
    If line& > Count& Then Exit Function
    If line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceCharacters(theline$, Chr(13), "")
        theline$ = ReplaceCharacters(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then FSpot = Len(MyString$)
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceCharacters(theline$, Chr(13), "")
        theline$ = ReplaceCharacters(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function

Public Function ListGetText(ListBox As Long, Index As Long) As String
    Dim ListText As String * 256
    Call SendMessageByString(ListBox&, LB_GETTEXT, Index&, ListText$)
    ListGetText$ = ListText$
End Function

Public Function ComboGetText(ComboBox As Long, Index As Long) As String
    Dim ComboText As String * 256
    Call SendMessageByString(ComboBox&, Cb_GetLbText, Index&, ComboText$)
    ComboGetText$ = ComboText$
End Function

Public Function RandomNumber(MaxNumber As Long) As Long
    Call Randomize
    RandomNumber& = Int((Val(MaxNumber&) * Rnd) + 1)
End Function

Public Function RandomLetter() As String
    Dim Random As Long
    Randomize
    Random& = Int(Rnd * 26) + 1
    If Random& = 1 Then RandomLetter$ = "a"
    If Random& = 2 Then RandomLetter$ = "b"
    If Random& = 3 Then RandomLetter$ = "c"
    If Random& = 4 Then RandomLetter$ = "d"
    If Random& = 5 Then RandomLetter$ = "e"
    If Random& = 6 Then RandomLetter$ = "f"
    If Random& = 7 Then RandomLetter$ = "g"
    If Random& = 8 Then RandomLetter$ = "h"
    If Random& = 9 Then RandomLetter$ = "i"
    If Random& = 10 Then RandomLetter$ = "j"
    If Random& = 11 Then RandomLetter$ = "k"
    If Random& = 12 Then RandomLetter$ = "l"
    If Random& = 13 Then RandomLetter$ = "m"
    If Random& = 14 Then RandomLetter$ = "n"
    If Random& = 15 Then RandomLetter$ = "o"
    If Random& = 16 Then RandomLetter$ = "p"
    If Random& = 17 Then RandomLetter$ = "q"
    If Random& = 18 Then RandomLetter$ = "r"
    If Random& = 19 Then RandomLetter$ = "s"
    If Random& = 20 Then RandomLetter$ = "t"
    If Random& = 21 Then RandomLetter$ = "u"
    If Random& = 22 Then RandomLetter$ = "v"
    If Random& = 23 Then RandomLetter$ = "w"
    If Random& = 24 Then RandomLetter$ = "x"
    If Random& = 25 Then RandomLetter$ = "y"
    If Random& = 26 Then RandomLetter$ = "z"
End Function

Public Sub TextLoad(LoadInThis As String, FilePath As String)
    Dim InputString As String
    On Error Resume Next
    Open FilePath$ For Input As #1
    InputString$ = Input(LOF(1), #1)
    Close #1
    LoadInThis$ = InputString$
End Sub

Public Sub TextSave(SaveThis As String, FilePath As String)
    Dim OutputString As String
    On Error Resume Next
    OutputString$ = SaveThis$
    Open FilePath$ For Output As #1
    Print #1, OutputString$
    Close #1
End Sub

Public Sub WavPlay(FilePath As String)
    Call sndPlaySound(FilePath$, Snd_Flags)
End Sub

Public Sub WavLoop(FilePath As String)
    Call sndPlaySound(FilePath$, Snd_Flags2)
End Sub

Public Sub WavStop()
    Call WavPlay(" ")
End Sub

Public Function GetFromINI(Section As String, KeyString As String, FilePath As String) As String
   Dim FixedString As String
   FixedString$ = String(750, Chr(0))
   KeyString$ = LCase$(KeyString$)
   GetFromINI$ = Left(FixedString$, GetPrivateProfileString(Section$, ByVal KeyString$, "", FixedString$, Len(FixedString$), FilePath$))
End Function

Public Sub WriteToINI(Section As String, KeyString As String, KeyValue As String, FilePath As String)
    Call WritePrivateProfileString(Section$, KeyString$, KeyValue$, FilePath$)
End Sub

Public Sub DisableX(WinHandle As Long)
    Dim SystemMenu As Long
    SystemMenu& = GetSystemMenu(WinHandle&, 0)
    Call RemoveMenu(SystemMenu&, 6, MF_BYPOSITION)
End Sub

Public Sub EnableX(WinHandle As Long)
    Dim SystemMenu As Long
    SystemMenu& = GetSystemMenu(WinHandle&, 1)
    Call RemoveMenu(SystemMenu&, 6, MF_BYPOSITION)
End Sub

Public Sub ListLoad(Directory As String, TheList As ListBox)
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

Public Sub ListSave(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.list(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub SystemTrayAddIcon(Form As Form)
    With SysTray
        .cbSize = Len(SysTray)
        .hwnd = Form.hwnd
        .uID = vbNull
        .uFlags = Sys_Icon Or Sys_Tip Or Sys_Message
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Form.Icon
        .szTip = Form.Caption & vbNullChar
    End With
    Call Shell_NotifyIcon(Sys_Add, SysTray)
End Sub

Public Sub SystemTrayAction()
    With SysTray
        .uFlags = Sys_Icon
        .uCallbackMessage = vbRightButton
        Call MsgBox(UserSN)
    End With
End Sub

Public Sub SystemTrayRemoveIcon(Form As Form)
    With SysTray
        .hwnd = Form.hwnd
    End With
    Call Shell_NotifyIcon(Sys_Delete, SysTray)
End Sub

Public Function IsSequential(FirstNumber As Long, SecondNumber As Long) As Boolean
    If SecondNumber& = (FirstNumber& - 1) Or SecondNumber& = (FirstNumber& + 1) Then
        IsSequential = True
        Exit Function
    End If
End Function

Public Sub Border3d(Control As Control, Style As Long, ThickNess As Long)
    Dim ObjectHeight As Long, ObjectWidth As Long, ObjectLeft As Long
    Dim ObjectTop As Long, OldScaleMode As Long, OldDrawWidth As Long
    Dim TopLeftShade As Long, BottomRightShade As Long, Thick As Long
    On Error Resume Next
    Control.Parent.AutoRedraw = True
    If ThickNess& <= 0 Then ThickNess& = 1
    If ThickNess& > 8 Then ThickNess& = 8
    OldScaleMode& = Control.Parent.ScaleMode
    OldDrawWidth& = Control.Parent.DrawWidth
    Control.Parent.ScaleMode = 3
    Control.Parent.DrawWidth = 1
    ObjectHeight& = Control.Height
    ObjectWidth& = Control.Width
    ObjectLeft& = Control.Left
    ObjectTop& = Control.Top
    Select Case Style&
        Case 1:
            TopLeftShade& = QBColor(8)
            BottomRightShade& = QBColor(15)
        Case 2:
            TopLeftShade& = QBColor(15)
            BottomRightShade& = QBColor(8)
        Case 3:
            TopLeftShade& = RGB(0, 0, 255)
            BottomRightShade& = QBColor(1)
    End Select
    For Thick& = 1 To ThickNess&
        Control.Parent.Line ((ObjectLeft& - Thick&), (ObjectTop& - Thick&))-Step((ObjectWidth& + (Thick& * 2) - 1), 0), TopLeftShade&
        Control.Parent.Line -Step(0, (ObjectHeight& + (Thick& * 2) - 1)), BottomRightShade&
        Control.Parent.Line -Step(-(ObjectWidth& + (Thick& * 2) - 1), 0), BottomRightShade&
        Control.Parent.Line -Step(0, -(ObjectHeight& + (Thick& * 2) - 1)), TopLeftShade&
    Next Thick&
    If ThickNess& > 2 Then
        Control.Parent.Line ((ObjectLeft& - ThickNess& - 1), (ObjectTop& - ThickNess& - 1))-Step((ObjectWidth& + ((ThickNess& + 1) * 2) - 1), 0), QBColor(0)
        Control.Parent.Line -Step(0, (ObjectHeight& + ((ThickNess& + 1) * 2) - 1)), QBColor(0)
        Control.Parent.Line -Step(-(ObjectWidth& + ((ThickNess& + 1) * 2) - 1), 0), QBColor(0)
        Control.Parent.Line -Step(0, -(ObjectHeight& + ((ThickNess& + 1) * 2) - 1)), QBColor(0)
    End If
    Control.Parent.ScaleMode = OldScaleMode&
    Control.Parent.DrawWidth = OldDrawWidth&
End Sub

Public Function FirstCharacter(ThisString As String, HtmlString As String) As String
    Dim prepstring As String, MidString As String, Space As Long
    Dim SpaceString As String, MidSpaceString As String
    On Error Resume Next
    If InStr(ThisString$, " ") = 0& Then
        MidString$ = Mid(ThisString$, 1, 1)
        MidString$ = "<" & HtmlString$ & ">" & MidString$ & "</" & HtmlString$ & ">"
        prepstring$ = MidString$ & Mid(ThisString$, 2)
        FirstCharacter$ = prepstring$
        Exit Function
       ElseIf InStr(ThisString$, " ") <> 0& Then
        For Space& = 1 To StringCount(ThisString$, " ") + 1
            SpaceString$ = GetInstance(ThisString$, " ", Space&)
            If TrimSpaces(SpaceString$) <> "" Then
                MidSpaceString$ = Mid(SpaceString$, 1, 1)
                MidSpaceString$ = "<" & HtmlString$ & ">" & MidSpaceString$ & "</" & HtmlString$ & ">"
                prepstring$ = prepstring$ & MidSpaceString$ & Mid(SpaceString$, 2) & " "
            End If
        Next Space&
        FirstCharacter$ = prepstring$
        Exit Function
    End If
End Function

Public Function FirstCharacterSecond(ThisString As String, HtmlString As String, HtmlString2 As String) As String
    Dim prepstring As String, MidString As String, Space As Long
    Dim SpaceString As String, MidSpaceString As String
    On Error Resume Next
    If InStr(ThisString$, " ") = 0& Then
        MidString$ = Mid(ThisString$, 1, 1)
        MidString$ = "<" & HtmlString$ & ">" & MidString$ & "</" & HtmlString$ & ">"
        prepstring$ = MidString$ & "<" & HtmlString2$ & ">" & Mid(ThisString$, 2) & "</" & HtmlString2$ & ">"
        FirstCharacterSecond$ = prepstring$
        Exit Function
       ElseIf InStr(ThisString$, " ") <> 0& Then
        For Space& = 1 To StringCount(ThisString$, " ") + 1
            SpaceString$ = GetInstance(ThisString$, " ", Space&)
            If TrimSpaces(SpaceString$) <> "" Then
                MidSpaceString$ = Mid(SpaceString$, 1, 1)
                MidSpaceString$ = "<" & HtmlString$ & ">" & MidSpaceString$ & "</" & HtmlString$ & ">"
                prepstring$ = prepstring$ & MidSpaceString$ & "<" & HtmlString2$ & ">" & Mid(SpaceString$, 2) & "</" & HtmlString2$ & ">" & " "
            End If
        Next Space&
        FirstCharacterSecond$ = prepstring$
        Exit Function
    End If
End Function

Public Function FindAChatWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Find a Chat")
   If AolChild& <> 0& Then
     FindAChatWindow& = AolChild&
     Exit Function
    Else
      AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", "Find a Chat")
      FindAChatWindow& = AolChild&
      Exit Function
   End If
End Function

Public Function FindErrorWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Error")
   If AolChild& <> 0& Then
     FindErrorWindow& = AolChild&
     Exit Function
    Else
      AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", "Error")
      FindErrorWindow& = AolChild&
      Exit Function
   End If
End Function

Public Function FindLocatedWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If InStr(GetCaption(AolChild&), "Locate ") <> 0& And GetCaption(AolChild&) <> "Locate Member Online" Then
        FindLocatedWindow& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         If InStr(GetCaption(AolChild&), "Locate ") <> 0& And GetCaption(AolChild&) <> "Locate Member Online" Then
             FindLocatedWindow& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindLocatedWindow& = AolChild&
End Function

Public Function FindRoom() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim RichText As Long, AolStatic As Long, AolGlyph As Long
    Dim ListBox As Long, AolIcon As Long, ComboBox As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    RichText& = FindWindowEx(AolChild&, 0&, "RICHCNTL", vbNullString)
    ListBox& = FindWindowEx(AolChild&, 0&, "_AOL_Listbox", vbNullString)
    AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    AolIcon& = FindWindowEx(AolChild&, 0&, "_AOL_Icon", vbNullString)
    AolGlyph& = FindWindowEx(AolChild&, 0&, "_AOL_Glyph", vbNullString)
    ComboBox& = FindWindowEx(AolChild&, 0&, "_AOL_Combobox", vbNullString)
    If RichText& <> 0& And ListBox& <> 0& And AolStatic& <> 0& And AolIcon& <> 0& And AolGlyph& <> 0& And ComboBox& <> 0& Then
        FindRoom& = AolChild&
        Exit Function
       Else
        Do
            AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            RichText& = FindWindowEx(AolChild&, 0&, "RICHCNTL", vbNullString)
            ListBox& = FindWindowEx(AolChild&, 0&, "_AOL_Listbox", vbNullString)
            AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
            AolIcon& = FindWindowEx(AolChild&, 0&, "_AOL_Icon", vbNullString)
            AolGlyph& = FindWindowEx(AolChild&, 0&, "_AOL_Glyph", vbNullString)
            ComboBox& = FindWindowEx(AolChild&, 0&, "_AOL_Combobox", vbNullString)
            If RichText& <> 0& And ListBox& <> 0& And AolStatic& <> 0& And AolIcon& <> 0& And AolGlyph& <> 0& And ComboBox& <> 0& Then
                FindRoom& = AolChild&
                Exit Function
            End If
        Loop Until AolChild& = 0&
    End If
    FindRoom& = AolChild&
End Function

Public Sub RoomSend(SendString As String, Optional ClearBefore As Boolean = False)
    If FindRoom& = 0& Then Exit Sub
    Dim RichText1 As Long, RichText2 As Long, TextOfRich As String
    RichText1& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
    RichText2& = FindWindowEx(FindRoom&, RichText1&, "RICHCNTL", vbNullString)
    If ClearBefore = True Then Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, SendString$)
    Call SendMessageLong(RichText2&, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Sub RoomSendSafe(SendString As String, Optional SendAfterPlaceBack As Boolean = False)
    If FindRoom& = 0& Then Exit Sub
    Dim RichText1 As Long, RichText2 As Long, TextOfRich As String
    RichText1& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
    RichText2& = FindWindowEx(FindRoom&, RichText1&, "RICHCNTL", vbNullString)
    TextOfRich$ = GetText(RichText2&)
    Call SendMessageByString(RichText2&, WM_CLEAR, 0&, 0&)
    Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, SendString$)
    Call SendMessageLong(RichText2&, WM_CHAR, ENTER_KEY, 0&)
    Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, TextOfRich$)
    If SendAfterPlaceBack = True Then Call SendMessageLong(RichText2&, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Sub ScrollString(ScrollThis As String, Optional Delay As Single = ".6")
    Dim PreString  As String, handler
    On Error GoTo handler
    If Mid(ScrollThis$, Len(ScrollThis$), 1) <> vbLf Then ScrollThis$ = ScrollThis$ & vbCrLf
    Do While InStr(ScrollThis$, vbCr) <> 0&
        If TrimSpaces(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1)) <> "" Then
            If Len(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1)) > 92 Then
                Call ScrollSplitString(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1))
               ElseIf Len(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1)) <= 92 Then
                Call RoomSend(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1))
                Yield Val(Delay)
            End If
        End If
        ScrollThis$ = Mid(ScrollThis$, InStr(ScrollThis$, vbCrLf) + 2)
    Loop
handler:
End Sub

Public Sub ScrollSplitString(SendString As String, Optional Delay As Single = "0.6")
    Dim LenString As Long
    If Len(SendString$) <= 92 Then
        RoomSend SendString$
        Exit Sub
       ElseIf Len(SendString$) > 92 Then
        RoomSend Mid(SendString$, 1, 92)
        Call Yield(Val(Delay))
        For LenString& = 1 To Len(SendString$) / 92
            RoomSend Mid(SendString$, (LenString& * 92) + 1, (LenString& * 92))
            Call Yield(Val(Delay))
        Next LenString&
    End If
End Sub

Public Sub ScrollProfile(ScreenName As String, Optional Delay As Single = "0.6")
    Call ScrollString(ProfileGet(ScreenName$), Delay)
End Sub

Public Sub ListScroll(ListBox As ListBox, Optional Delay As Single = "0.6")
    Dim ListIndex As Long
    For ListIndex& = 0 To ListBox.ListCount - 1
        Call RoomSend(ListBox.list(ListIndex&))
        Yield Val(Delay)
    Next ListIndex&
End Sub

Public Sub ComboScroll(ComboBox As ComboBox, Optional Delay As Single = "0.6")
    Dim ComboIndex As Long
    For ComboIndex& = 0 To ComboBox.ListCount - 1
        Call RoomSend(ComboBox.list(ComboIndex&))
        Yield Val(Delay)
    Next ComboIndex&
End Sub

Public Function RoomName() As String
    If FindRoom& = 0& Then Exit Function
    RoomName$ = GetCaption(FindRoom&)
End Function

Public Sub RoomClear()
    If FindRoom& = 0& Then Exit Sub
    Dim RichText As Long
    RichText& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, "")
End Sub

Public Function RoomCount() As Long
   If FindRoom& = 0& Then Exit Function
    Dim AolList As Long
    AolList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    RoomCount& = ListCount(AolList&)
End Function

Public Function RoomSendRich() As Long
    Dim RichText As Long
    RichText& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
    RoomSendRich& = FindWindowEx(FindRoom&, RichText&, "RICHCNTL", vbNullString)
End Function

Public Function RoomGetText() As String
    Dim RichText As Long
    RichText& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
    RoomGetText$ = GetText(RichText&)
End Function

Public Function RoomLocator(ScreenName As String, RoomList As Control, Optional BustIfFull As Boolean = True, Optional LimitTriesOnBust As Long = "20") As String
   '*TESTED*, if you know the rooms the person goes in this subs kicks ass
    Dim ListIndex As Long
    If RoomSearch(ScreenName$) = True Then
        RoomLocator$ = ScreenName & " has been found"
        Exit Function
       Else
        RoomLocator$ = ScreenName & " was not found"
    End If
    For ListIndex& = 0 To RoomList.ListCount - 1
       If BustIfFull = False Then
           Call KeyWord("aol://2719:2-2-" & RoomList.list(ListIndex&))
           WaitForOkorRoom RoomList.list(ListIndex&)
          ElseIf BustIfFull = True Then
           Call RoomForceEnter("aol://2719:2-2-", RoomList.list(ListIndex&), False, 0.2, LimitTriesOnBust&)
       End If
       Yield 0.6
       If RoomSearch(ScreenName$) = True Then
           Yield 0.6
           RoomLocator$ = ScreenName & " has been found"
           Exit Function
          Else
           RoomLocator$ = ScreenName & " was not found"
       End If
       Yield 2
    Next ListIndex&
End Function

Public Function RoomLocatorBuddy(ScreenName As String, Optional BustIfFull As Boolean = True, Optional LimitTriesOnBust As Long = "20") As String
    Dim ListIndex As Long, BuddyChat As String
    If RoomSearch(ScreenName$) = True Then
        RoomLocatorBuddy$ = ScreenName & " has been found"
        Exit Function
       Else
        RoomLocatorBuddy$ = ScreenName & " was not found"
    End If
    For ListIndex& = 1 To 99
       If ListIndex& = 1 Or ListIndex& = 2 Or ListIndex& = 3 Or ListIndex& = 4 Or ListIndex& = 5 Or ListIndex& = 6 Or ListIndex& = 7 Or ListIndex& = 8 Or ListIndex& = 9 Then
           BuddyChat$ = ScreenName & "0" & ListIndex&
          Else
           BuddyChat$ = ScreenName & ListIndex&
       End If
       If BustIfFull = False Then
           Call KeyWord("aol://2719:2-2-" & BuddyChat$)
           WaitForOkorRoom BuddyChat$
          ElseIf BustIfFull = True Then
           Call RoomForceEnter("aol://2719:2-2-", BuddyChat$, False, 2, LimitTriesOnBust&)
       End If
       Yield 0.6
       If RoomSearch(ScreenName$) = True Then
           Yield 0.6
           RoomLocatorBuddy$ = ScreenName & " has been found"
           Exit Function
          Else
           RoomLocatorBuddy$ = ScreenName & " was not found"
       End If
       Yield 2
    Next ListIndex&
End Function

Public Function FindMailBox() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolStatic As Long, AolImage As Long, AolGlyph As Long
    Dim AolIcon As Long, TabControl As Long, TabPage As Long
    Dim AolTree As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    AolImage& = FindWindowEx(AolChild&, 0&, "_AOL_Image", vbNullString)
    AolGlyph& = FindWindowEx(AolChild&, 0&, "_AOL_Glyph", vbNullString)
    AolIcon& = FindWindowEx(AolChild&, 0&, "_AOL_Icon", vbNullString)
    TabControl& = FindWindowEx(AolChild&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    If AolStatic& <> 0& And AolImage& <> 0& And AolGlyph& <> 0& And AolIcon& <> 0& And TabControl& <> 0& And TabPage& <> 0 And AolTree& <> 0& Then
        FindMailBox& = AolChild&
        Exit Function
       Else
        Do
            AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
            AolImage& = FindWindowEx(AolChild&, 0&, "_AOL_Image", vbNullString)
            AolGlyph& = FindWindowEx(AolChild&, 0&, "_AOL_Glyph", vbNullString)
            AolIcon& = FindWindowEx(AolChild&, 0&, "_AOL_Icon", vbNullString)
            TabControl& = FindWindowEx(AolChild&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
            If AolStatic& <> 0& And AolImage& <> 0& And AolGlyph& <> 0& And AolIcon& <> 0& And TabControl& <> 0& And TabPage& <> 0 And AolTree& <> 0& Then
                FindMailBox& = AolChild&
                Exit Function
            End If
        Loop Until AolChild& = 0&
    End If
    FindMailBox& = AolChild&
End Function

Public Function FindSendWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolStatic1 As Long, AolStatic2 As Long, AolStatic3 As Long
    Dim AolStatic4 As Long, AolStatic5 As Long, AolStatic6 As Long
    Dim AolStatic7 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
    AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
    AolStatic4& = FindWindowEx(AolChild&, AolStatic3&, "_AOL_Static", vbNullString)
    AolStatic5& = FindWindowEx(AolChild&, AolStatic4&, "_AOL_Static", vbNullString)
    AolStatic6& = FindWindowEx(AolChild&, AolStatic5&, "_AOL_Static", vbNullString)
    AolStatic7& = FindWindowEx(AolChild&, AolStatic6&, "_AOL_Static", vbNullString)
    If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& And AolStatic4& <> 0& And AolStatic5& <> 0& And AolStatic6& <> 0& And AolStatic7& <> 0& Then
        If GetText(AolStatic7&) = "Send Now" Then
            FindSendWindow& = AolChild&
            Exit Function
        End If
       Else
        Do
            AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
            AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
            AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
            AolStatic4& = FindWindowEx(AolChild&, AolStatic3&, "_AOL_Static", vbNullString)
            AolStatic5& = FindWindowEx(AolChild&, AolStatic4&, "_AOL_Static", vbNullString)
            AolStatic6& = FindWindowEx(AolChild&, AolStatic5&, "_AOL_Static", vbNullString)
            AolStatic7& = FindWindowEx(AolChild&, AolStatic6&, "_AOL_Static", vbNullString)
            If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& And AolStatic4& <> 0& And AolStatic5& <> 0& And AolStatic6& <> 0& And AolStatic7& <> 0& Then
                If GetText(AolStatic7&) = "Send Now" Then
                    FindSendWindow& = AolChild&
                    Exit Function
                End If
            End If
        Loop Until AolChild& = 0&
    End If
    FindSendWindow& = AolChild&
End Function

Public Function FindFwdWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    If InStr(GetCaption(AolChild&), "Fwd:") <> 0& Then
        FindFwdWindow& = AolChild&
        Exit Function
       Else
        Do
            AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            If InStr(GetCaption(AolChild&), "Fwd:") <> 0& Then
                FindFwdWindow& = AolChild&
                Exit Function
            End If
        Loop Until AolChild& = 0&
    End If
    FindFwdWindow& = AolChild&
End Function

Public Function FindForwardWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolStatic1 As Long, AolStatic2 As Long, AolStatic3 As Long
    Dim AolStatic4 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
    AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
    AolStatic4& = FindWindowEx(AolChild&, AolStatic3&, "_AOL_Static", vbNullString)
    If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& And AolStatic4& <> 0& Then
        If GetText(AolStatic4&) = "Forward" Then
            FindForwardWindow& = AolChild&
            Exit Function
        End If
       Else
        Do
            AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
            AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
            AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
            AolStatic4& = FindWindowEx(AolChild&, AolStatic3&, "_AOL_Static", vbNullString)
            If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& And AolStatic4& <> 0& Then
                If GetText(AolStatic4&) = "Forward" Then
                    FindForwardWindow& = AolChild&
                    Exit Function
                End If
            End If
        Loop Until AolChild& = 0&
    End If
    FindForwardWindow& = AolChild&
End Function

Public Function FindReWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If InStr(GetCaption(AolChild&), "Re:") <> 0& Then
        FindReWindow& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         If InStr(GetCaption(AolChild&), "Re:") <> 0& Then
             FindReWindow& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindReWindow& = AolChild&
End Function

Public Function FindReplyWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolStatic1 As Long, AolStatic2 As Long, AolStatic3 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
    AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
   If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& Then
        If GetText(AolStatic3&) = "Reply" Then
            FindReplyWindow& = AolChild&
            Exit Function
        End If
    Else
       Do
         AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
         AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
         AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
          If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& Then
           If GetText(AolStatic3&) = "Reply" Then
               FindReplyWindow& = AolChild&
               Exit Function
           End If
          End If
       Loop Until AolChild& = 0&
   End If
    FindReplyWindow& = AolChild&
End Function

Public Function FindFlashMailBox() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If InStr(GetCaption(AolChild&), "/Saved Mail") <> 0& Then
        FindFlashMailBox& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         If InStr(GetCaption(AolChild&), "/Saved Mail") <> 0& Then
             FindFlashMailBox& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindFlashMailBox& = AolChild&
End Function

Public Function FindIm() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If InStr(GetCaption(AolChild&), "Instant Message") <> 0& Then
        FindIm& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         If InStr(GetCaption(AolChild&), "Instant Message") <> 0& Then
             FindIm& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindIm& = AolChild&
End Function

Public Function FindImFromAim() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If InStr(GetCaption(AolChild&), "Instant Message") = 1& Then
        FindImFromAim& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         If InStr(GetCaption(AolChild&), "Instant Message") = 1& Then
             FindImFromAim& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindImFromAim& = AolChild&
End Function

Public Function FindBuddyList() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If GetCaption(AolChild&) = UserSN & "'s Buddy List" Or GetCaption(AolChild&) = UserSN & "'s Buddy Lists" Then
        FindBuddyList& = AolChild&
        Exit Function
     Else
         Do
           AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            If GetCaption(AolChild&) = UserSN & "'s Buddy List" Or GetCaption(AolChild&) = UserSN & "'s Buddy Lists" Then
                FindBuddyList& = AolChild&
                Exit Function
            End If
         Loop Until AolChild& = 0&
   End If
    FindBuddyList& = AolChild&
End Function

Public Function FindBuddyView() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If GetCaption(AolChild&) = "Buddy List Window" Then
        FindBuddyView& = AolChild&
        Exit Function
     Else
         Do
           AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            If GetCaption(AolChild&) = "Buddy List Window" Then
                FindBuddyView& = AolChild&
                Exit Function
            End If
         Loop Until AolChild& = 0&
   End If
    FindBuddyView& = AolChild&
End Function

Public Function FindWelcome() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim WelcomeCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    WelcomeCaption$ = GetCaption(AolChild&)
   If InStr(WelcomeCaption$, "Welcome, ") <> 0& Then
        FindWelcome& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
        WelcomeCaption$ = GetCaption(AolChild&)
         If InStr(WelcomeCaption$, "Welcome, ") <> 0& Then
             FindWelcome& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindWelcome& = AolChild&
End Function

Public Function UserSN() As String
   If FindWelcome& = 0& Then Exit Function
    UserSN$ = Mid$(GetCaption(FindWelcome&), 10, (InStr(GetCaption(FindWelcome&), "!") - 10))
End Function

Public Function Online() As Boolean
    If FindWelcome& <> 0& Then
        Online = True
       ElseIf FindWelcome& = 0& Then
        Online = False
    End If
End Function

Public Sub MailRemoveFwd()
    Dim AolEdit1 As Long, AolEdit2 As Long, AolEdit3 As Long
   If InStr(GetCaption(FindFwdWindow&), "Fwd:") = 0& Then Exit Sub
    AolEdit1& = FindWindowEx(FindFwdWindow&, 0&, "_AOL_Edit", vbNullString)
    AolEdit2& = FindWindowEx(FindFwdWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
    AolEdit3& = FindWindowEx(FindFwdWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(AolEdit3&, WM_SETTEXT, 0&, Mid(GetCaption(FindSendWindow&), 6))
End Sub

Public Sub MailRemoveRe()
    Dim AolEdit1 As Long, AolEdit2 As Long, AolEdit3 As Long
   If InStr(GetCaption(FindReWindow&), "Re:") = 0& Then Exit Sub
    AolEdit1& = FindWindowEx(FindReWindow&, 0&, "_AOL_Edit", vbNullString)
    AolEdit2& = FindWindowEx(FindReWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
    AolEdit3& = FindWindowEx(FindReWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(AolEdit3&, WM_SETTEXT, 0&, Mid(GetCaption(FindSendWindow&), 5))
End Sub

Public Sub MailRemoveErrorNames(ListBox As Control)
    Dim AolView As Long, ErrorText As String, ListIndex As Long, ListError
     AolView& = FindWindowEx(FindErrorWindow, 0, "_AOL_View", vbNullString)
    ErrorText$ = GetText(AolView&)
    
    If FindErrorWindow = 0 Then Exit Sub
   On Error GoTo ListError
    Do: DoEvents
        If InStr(ErrorText$, ListBox.list(ListIndex)) <> 1 Then
  Form1.Label15.Caption = (ListIndex)
            ListBox.RemoveItem (ListIndex)
            ListIndex = ListIndex + 1
        End If
    Loop Until ListCount(ListBox.hwnd)
ListError:
End Sub

Public Function MailErrorNameCount() As Long
    Dim AolView As Long, ErrorText As String
    If FindErrorWindow& = 0& Then Exit Function
    AolView& = FindWindowEx(FindErrorWindow&, 0&, "_AOL_View", vbNullString)
    ErrorText$ = GetText(AolView&)
    MailErrorNameCount& = StringCount(LCase(TrimSpaces(ErrorText$)), LCase(TrimSpaces(" - This is not a known member.")))
End Function

Public Function MailStatusNew(MailIndex As Long) As String
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim StatusWindow As Long, AolIcon1 As Long, AolIcon2 As Long
    Dim AolView As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabPage& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    AolIcon1& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
    AolIcon2& = FindWindowEx(FindMailBox&, AolIcon1&, "_AOL_Icon", vbNullString)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        StatusWindow& = FindMailStatusWindow&
        AolView& = FindWindowEx(StatusWindow&, 0&, "_AOL_View", vbNullString)
    Loop Until StatusWindow& <> 0& And AolView& <> 0& And GetText(AolView&) <> ""
    MailStatusNew$ = GetText(AolView&)
    Call WinClose(FindMailStatusWindow&)
End Function

Public Function MailStatusOld(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim TabPage2 As Long, StatusWindow As Long, AolIcon1 As Long
    Dim AolIcon2 As Long, AolView As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    If TabPage2& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    AolIcon1& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
    AolIcon2& = FindWindowEx(FindMailBox&, AolIcon1&, "_AOL_Icon", vbNullString)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        StatusWindow& = FindMailStatusWindow&
        AolView& = FindWindowEx(StatusWindow&, 0&, "_AOL_View", vbNullString)
    Loop Until StatusWindow& <> 0& And AolView& <> 0& And GetText(AolView&) <> ""
    MailStatusOld$ = GetText(AolView&)
    Call WinClose(FindMailStatusWindow&)
End Function

Public Function MailStatusSent(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim TabPage2 As Long, TabPage3 As Long, StatusWindow As Long
    Dim AolIcon1 As Long, AolIcon2 As Long, AolView As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
    If TabPage3& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    AolIcon1& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
    AolIcon2& = FindWindowEx(FindMailBox&, AolIcon1&, "_AOL_Icon", vbNullString)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        StatusWindow& = FindMailStatusWindow&
        AolView& = FindWindowEx(StatusWindow&, 0&, "_AOL_View", vbNullString)
    Loop Until StatusWindow& <> 0& And AolView& <> 0& And GetText(AolView&) <> ""
    MailStatusSent$ = GetText(AolView&)
    Call WinClose(FindMailStatusWindow&)
End Function

Public Sub KeyWord(KwString As String, Optional ClearBefore As Boolean = True)
    Dim AolFrame As Long, AolToolbar As Long, Toolbar As Long
    Dim ComboBox As Long, Edit As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    ComboBox& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    Edit& = FindWindowEx(ComboBox&, 0&, "Edit", vbNullString)
    If ClearBefore = True Then Call SendMessageByString(Edit&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, KwString$)
    Call SendMessageLong(Edit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(Edit&, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Sub ClearHistory()
    Dim AolFrame As Long, AolToolbar As Long, Toolbar As Long
    Dim ComboBox As Long, Edit As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    ComboBox& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    Edit& = FindWindowEx(ComboBox&, 0&, "Edit", vbNullString)
    Call SendMessageLong(ComboBox&, Cb_ResetContent, 0&, 0&)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, "Type Keyword or Web Address here and click Go")
End Sub

Public Sub PopUpIcon(IconNumber As Long, Character As String)
    Dim Message1 As Long, Message2 As Long, AolFrame As Long
    Dim AolToolbar As Long, Toolbar As Long, AolIcon As Long
    Dim NextOfClass As Long, AscCharacter As Long
    Message1& = FindWindow("#32768", vbNullString)
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    For NextOfClass& = 1 To IconNumber&
        AolIcon& = GetWindow(AolIcon&, 2)
    Next NextOfClass&
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Message2& = FindWindow("#32768", vbNullString)
    Loop Until Message2& <> Message1&
    AscCharacter& = Asc(Character$)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
End Sub

Public Sub PopUpIconDbl(IconNumber As Long, Character As String, Character2 As String)
    Dim Message1 As Long, Message2 As Long, AolFrame As Long
    Dim AolToolbar As Long, Toolbar As Long, AolIcon As Long
    Dim NextOfClass As Long, AscCharacter As Long, AscCharacter2 As Long
    Message1& = FindWindow("#32768", vbNullString)
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    For NextOfClass& = 1 To IconNumber&
        AolIcon& = GetWindow(AolIcon&, 2)
    Next NextOfClass&
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Message2& = FindWindow("#32768", vbNullString)
    Loop Until Message2& <> Message1&
    AscCharacter& = Asc(Character$)
    AscCharacter2& = Asc(Character2$)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter2&, 0&)
End Sub

Public Function RoomIsPrivate() As Boolean
    Dim AolImage As Long
    AolImage& = FindWindowEx(FindRoom&, 0&, "_AOL_Image", vbNullString)
    If IsWindowVisible(AolImage&) = False Then
        RoomIsPrivate = True
       ElseIf IsWindowVisible(AolImage&) = True Then
        RoomIsPrivate = False
    End If
End Function

Public Sub AntiIdle_45()
    Dim AolModal As Long, ModalIcon As Long, AolPalette As Long
    Dim PaletteIcon As Long
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    ModalIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    AolPalette& = FindWindow("_AOL_Palette", vbNullString)
    PaletteIcon& = FindWindowEx(AolPalette&, 0&, "_AOL_Icon", vbNullString)
    If AolModal& <> 0& And ModalIcon& <> 0& Then
        Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
       ElseIf AolPalette& <> 0& And PaletteIcon& <> 0& Then
        Call PostMessage(PaletteIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(PaletteIcon&, WM_LBUTTONUP, 0&, 0&)
    End If
End Sub

Public Sub KillModal()
    Dim AolModal As Long, AolIcon As Long
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    AolIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub KillTimer()
    Dim AolPalette As Long, AolIcon As Long
    AolPalette& = FindWindow("_AOL_Palette", vbNullString)
    AolIcon& = FindWindowEx(AolPalette&, 0&, "_AOL_Icon", vbNullString)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub DisableTimer()
    Dim AolTimeKeeper As Long
    AolTimeKeeper& = FindWindow("_AOL_TimeKeeper", vbNullString)
    Call WinDisable(AolTimeKeeper&)
End Sub

Public Sub EnableTimer()
    Dim AolTimeKeeper As Long
    AolTimeKeeper& = FindWindow("_AOL_TimeKeeper", vbNullString)
    Call WinEnable(AolTimeKeeper&)
End Sub

Public Sub InstantMessage(Person As String, message As String)
    Dim AolFrame As Long, AolMdi As Long, ImSendWindow As Long
    Dim RichText As Long, AolEdit As Long, AolIcon As Long
    Dim MessageOk As Long, OKButton As Long
    Call PopUpIcon(9, "I")
    Do: DoEvents
        AolFrame& = FindWindow("AOL Frame25", vbNullString)
        AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
        ImSendWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Send Instant Message")
        AolEdit& = FindWindowEx(ImSendWindow&, 0&, "_AOL_Edit", vbNullString)
        RichText& = FindWindowEx(ImSendWindow&, 0&, "RICHCNTL", vbNullString)
        AolIcon& = NextOfClassByCount(ImSendWindow&, "_AOL_Icon", 9)
    Loop Until ImSendWindow& <> 0& And AolEdit& <> 0& And RichText& <> 0& And AolIcon& <> 0&
    Call SendMessageByString(AolEdit&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWindowEx(AolMdi&, 0&, "AOL Child", "Send Instant Message") = 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(ImSendWindow&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
End Sub

Public Sub InstantMessageFast(Person As String, message As String)
    Dim AolFrame As Long, AolMdi As Long, ImSendWindow As Long
    Dim RichText As Long, AolIcon As Long
    Dim MessageOk As Long, OKButton As Long
    Call KeyWord("aol://9293:" & Person$)
    Do: DoEvents
        AolFrame& = FindWindow("AOL Frame25", vbNullString)
        AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
        ImSendWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Send Instant Message")
        RichText& = FindWindowEx(ImSendWindow&, 0&, "RICHCNTL", vbNullString)
        AolIcon& = NextOfClassByCount(ImSendWindow&, "_AOL_Icon", 9)
    Loop Until ImSendWindow& <> 0& And RichText& <> 0& And AolIcon& <> 0&
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWindowEx(AolMdi&, 0&, "AOL Child", "Send Instant Message") = 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(ImSendWindow&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
End Sub

Public Sub MassInstantMessage(ScreenNameList As Control, message As String, Optional Delay As Single = "0.6")
    Dim ListIndex As Long
    On Error Resume Next
    For ListIndex& = 0 To ScreenNameList.ListCount - 1
        Call InstantMessage(ScreenNameList.list(ListIndex&), message$)
        Call Yield(Val(Delay))
    Next ListIndex&
End Sub

Public Sub ImsOff()
    Call InstantMessageFast("$IM_OFF", "Http://chronx.cjb.net")
End Sub

Public Sub ImsOn()
    Call InstantMessageFast("$IM_ON", "Http://chronx.cjb.net")
End Sub

Public Sub ImsOffSpecific(Person As String)
    Call InstantMessageFast("$IM_OFF " & Person$, "Now Ignoring " & Person$)
End Sub

Public Sub ImsOnSpecific(Person As String)
    Call InstantMessageFast("$IM_ON " & Person$, "No Longer Ignoring " & Person$)
End Sub

Public Sub MailSend(Person As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
    Dim AolFrame As Long, AolToolbar As Long, Toolbar As Long
    Dim AolIcon1 As Long, AolEdit1 As Long
    Dim AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim AolIcon2 As Long, AolModal As Long, AolIcon3 As Long
    Dim CheckBox As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon1& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
    Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
        RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
        AolIcon2& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
    Loop Until FindSendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0& And AolIcon2& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(AolEdit3&, WM_SETTEXT, 0&, Subject$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    If CheckReturnReceipts = True Then
        CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
    End If
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolIcon3& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolIcon3& <> 0&
    If AolModal& <> 0& And FindWindowEx(AolMdi&, 0&, "AOL Child", "Write Mail") = 0& Then
        Call PostMessage(AolIcon3&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon3&, WM_LBUTTONUP, 0&, 0&)
        Exit Sub
       ElseIf FindWindowEx(AolMdi&, 0&, "AOL Child", "Write Mail") = 0& And AolModal& = 0& Then
        Exit Sub
    End If
End Sub

Public Sub MailSendNoKill(Person As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
    Dim AolFrame As Long, AolToolbar As Long, Toolbar As Long
    Dim AolIcon1 As Long, AolEdit1 As Long
    Dim AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim AolIcon2 As Long, CheckBox As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon1& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
    Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
        RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
        AolIcon2& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
    Loop Until FindSendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0& And AolIcon2& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(AolEdit3&, WM_SETTEXT, 0&, Subject$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    If CheckReturnReceipts = True Then
        CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
    End If
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub GhostOn()
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolIcon1 As Long, AolIcon2 As Long
    Dim PrivacyWindow As Long, CheckBox1 As Long, CheckBox2 As Long
    Dim CheckBox3 As Long, CheckBox4 As Long, CheckBox5 As Long
    Dim CheckBox6 As Long, CheckBox7 As Long, AolIcon3 As Long
    Dim AolIcon4 As Long, AolIcon5 As Long, MessageOk As Long
    Dim OKButton As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    If FindBuddyList& = 0& Then
        Call PopUpIcon(5, "B")
       Else
    End If
    Do: DoEvents
        AolIcon1& = NextOfClassByCount(FindBuddyList&, "_AOL_Icon", 5)
    Loop Until AolIcon1& <> 0&
    Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        PrivacyWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Privacy Preferences")
        CheckBox1& = FindWindowEx(PrivacyWindow&, 0&, "_AOL_Checkbox", vbNullString)
        CheckBox2& = FindWindowEx(PrivacyWindow&, CheckBox1&, "_AOL_Checkbox", vbNullString)
        CheckBox3& = FindWindowEx(PrivacyWindow&, CheckBox2&, "_AOL_Checkbox", vbNullString)
        CheckBox4& = FindWindowEx(PrivacyWindow&, CheckBox3&, "_AOL_Checkbox", vbNullString)
        CheckBox5& = FindWindowEx(PrivacyWindow&, CheckBox4&, "_AOL_Checkbox", vbNullString)
        CheckBox6& = FindWindowEx(PrivacyWindow&, CheckBox5&, "_AOL_Checkbox", vbNullString)
        CheckBox7& = FindWindowEx(PrivacyWindow&, CheckBox6&, "_AOL_Checkbox", vbNullString)
        AolIcon2& = FindWindowEx(PrivacyWindow&, 0&, "_AOL_Icon", vbNullString)
        AolIcon3& = FindWindowEx(PrivacyWindow&, AolIcon2&, "_AOL_Icon", vbNullString)
        AolIcon4& = FindWindowEx(PrivacyWindow&, AolIcon3&, "_AOL_Icon", vbNullString)
        AolIcon5& = FindWindowEx(PrivacyWindow&, AolIcon4&, "_AOL_Icon", vbNullString)
    Loop Until PrivacyWindow& <> 0& And CheckBox1& <> 0& And CheckBox2& <> 0& And CheckBox3& <> 0& And CheckBox4& <> 0& And CheckBox5& <> 0& And AolIcon2& <> 0& And AolIcon3& <> 0& And AolIcon4& <> 0& And AolIcon5& <> 0&
    Call PostMessage(CheckBox5&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(CheckBox5&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(CheckBox7&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(CheckBox7&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(AolIcon5&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon5&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(FindBuddyList&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Sub GhostOff()
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolIcon1 As Long, AolIcon2 As Long
    Dim PrivacyWindow As Long, CheckBox1 As Long, CheckBox2 As Long
    Dim CheckBox3 As Long, CheckBox4 As Long, CheckBox5 As Long
    Dim CheckBox6 As Long, CheckBox7 As Long, AolIcon3 As Long
    Dim AolIcon4 As Long, AolIcon5 As Long, MessageOk As Long
    Dim OKButton As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    If FindBuddyList& = 0& Then
        Call PopUpIcon(5, "B")
       Else
    End If
    Do: DoEvents
        AolIcon1& = NextOfClassByCount(FindBuddyList&, "_AOL_Icon", 5)
    Loop Until AolIcon1& <> 0&
    Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        PrivacyWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Privacy Preferences")
        CheckBox1& = FindWindowEx(PrivacyWindow&, 0&, "_AOL_Checkbox", vbNullString)
        CheckBox2& = FindWindowEx(PrivacyWindow&, CheckBox1&, "_AOL_Checkbox", vbNullString)
        CheckBox3& = FindWindowEx(PrivacyWindow&, CheckBox2&, "_AOL_Checkbox", vbNullString)
        CheckBox4& = FindWindowEx(PrivacyWindow&, CheckBox3&, "_AOL_Checkbox", vbNullString)
        CheckBox5& = FindWindowEx(PrivacyWindow&, CheckBox4&, "_AOL_Checkbox", vbNullString)
        CheckBox6& = FindWindowEx(PrivacyWindow&, CheckBox5&, "_AOL_Checkbox", vbNullString)
        CheckBox7& = FindWindowEx(PrivacyWindow&, CheckBox6&, "_AOL_Checkbox", vbNullString)
        AolIcon2& = FindWindowEx(PrivacyWindow&, 0&, "_AOL_Icon", vbNullString)
        AolIcon3& = FindWindowEx(PrivacyWindow&, AolIcon2&, "_AOL_Icon", vbNullString)
        AolIcon4& = FindWindowEx(PrivacyWindow&, AolIcon3&, "_AOL_Icon", vbNullString)
        AolIcon5& = FindWindowEx(PrivacyWindow&, AolIcon4&, "_AOL_Icon", vbNullString)
    Loop Until PrivacyWindow& <> 0& And CheckBox1& <> 0& And CheckBox2& <> 0& And CheckBox3& <> 0& And CheckBox4& <> 0& And CheckBox5& <> 0& And AolIcon2& <> 0& And AolIcon3& <> 0& And AolIcon4& <> 0& And AolIcon5& <> 0&
    Call PostMessage(CheckBox1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(CheckBox1&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(CheckBox7&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(CheckBox7&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(AolIcon5&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon5&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(FindBuddyList&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Function Guest() As Boolean
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AddressBook As Long, MessageOk As Long, OKButton As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    Call PopUpIcon(2, "A")
    Do: DoEvents
        AddressBook& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Address Book")
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until AddressBook& <> 0& Or MessageOk& <> 0&
    If AddressBook& <> 0& Then
        Call PostMessage(AddressBook&, WM_CLOSE, 0&, 0&)
        Guest = False
        Exit Function
       ElseIf MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Guest = True
        Exit Function
    End If
End Function

Public Function Guest2() As Boolean
    Dim MailPreferencesWindow As Long, FifthCheckBox As Long, AolSpin As Long
    Call PopUpIcon(2, "P")
    Do: DoEvents
        MailPreferencesWindow& = FindWindow("_AOL_Modal", "Mail Preferences")
        AolSpin& = FindWindowEx(MailPreferencesWindow&, 0&, "_AOL_Spin", vbNullString)
    Loop Until MailPreferencesWindow& <> 0& And AolSpin& <> 0&
    DoEvents
    FifthCheckBox& = NextOfClassByCount(MailPreferencesWindow&, "_AOL_Checkbox", 5)
    If IsWindowVisible(FifthCheckBox&) = 0& Then
        Guest2 = True
        Call WinClose(MailPreferencesWindow&)
        Exit Function
       ElseIf IsWindowVisible(FifthCheckBox&) <> 0& Then
        Guest2 = False
        Call WinClose(MailPreferencesWindow&)
        Exit Function
    End If
End Function

Public Function LocateMember(Person As String) As String
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim LocateMemberWindow As Long, AolEdit As Long, Static1 As Long
    Dim Static2 As Long, LocatedWindow As Long, message As Long
    Dim Button As Long, AolStatic As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    Call PopUpIcon(9, "L")
    Do: DoEvents
        LocateMemberWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Locate Member Online")
        AolEdit& = FindWindowEx(LocateMemberWindow&, 0&, "_AOL_Edit", vbNullString)
    Loop Until LocateMemberWindow& <> 7 And AolEdit& <> 0&
    Call SendMessageByString(AolEdit&, WM_SETTEXT, 0&, Person$)
    Call SendMessageLong(AolEdit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(AolEdit&, WM_CHAR, VK_RETURN, 0&)
    Do: DoEvents
        LocatedWindow& = FindLocatedWindow&
        message& = FindWindow("#32770", "America Online")
    Loop Until LocatedWindow& <> 0& Or message& <> 0&
    If LocatedWindow& <> 0& Then
        AolStatic& = FindWindowEx(LocatedWindow&, 0&, "_AOL_Static", vbNullString)
        LocateMember$ = GetText(AolStatic&)
        Call PostMessage(LocatedWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(LocateMemberWindow&, WM_CLOSE, 0&, 0&)
        Exit Function
       Else
        Static1& = FindWindowEx(message&, 0&, "Static", vbNullString)
        Static2& = FindWindowEx(message&, Static1&, "Static", vbNullString)
        Button& = FindWindowEx(message&, 0&, "Button", vbNullString)
        LocateMember$ = ReplaceCharacters(GetText(Static2&), "Member", Person$)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(LocateMemberWindow&, WM_CLOSE, 0&, 0&)
        Exit Function
    End If
End Function

Public Function LocateMemberFast(Person As String) As String
    Dim Static2 As Long, LocatedWindow As Long, message As Long
    Dim Button As Long, AolStatic As Long, Static1 As Long
    Call KeyWord("aol://3548:" & Person$)
    Do: DoEvents
        LocatedWindow& = FindLocatedWindow&
        message& = FindWindow("#32770", "America Online")
    Loop Until LocatedWindow& <> 0& Or message& <> 0&
    If LocatedWindow& <> 0& Then
        AolStatic& = FindWindowEx(LocatedWindow&, 0&, "_AOL_Static", vbNullString)
        LocateMemberFast$ = GetText(AolStatic&)
        Call PostMessage(LocatedWindow&, WM_CLOSE, 0&, 0&)
        Exit Function
       Else
        Static1& = FindWindowEx(message&, 0&, "Static", vbNullString)
        Static2& = FindWindowEx(message&, Static1&, "Static", vbNullString)
        Button& = FindWindowEx(message&, 0&, "Button", vbNullString)
        LocateMemberFast$ = ReplaceCharacters(GetText(Static2&), "Member", Person$)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Exit Function
    End If
End Function

Public Sub MassLocator(NamesList As Control)
    Dim ListIndex As Long
    On Error Resume Next
    For ListIndex& = 0& To NamesList.ListCount - 1
        NamesList.list(ListIndex&) = LocateMember(NamesList.list(ListIndex&))
        DoEvents
    Next ListIndex&
End Sub

Public Function ProfileGet(Person As String) As String
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim GetProfileWindow As Long, AolEdit As Long, GotProfileWindow As Long
    Dim AolView As Long, message As Long, Button As Long
    Dim Static1 As Long, Static2 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    Call PopUpIcon(9, "G")
    Do: DoEvents
        GetProfileWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Get a Member's Profile")
        AolEdit& = FindWindowEx(GetProfileWindow&, 0&, "_AOL_Edit", vbNullString)
    Loop Until GetProfileWindow& <> 7 And AolEdit& <> 0&
    Call SendMessageByString(AolEdit&, WM_SETTEXT, 0&, Person$)
    Call SendMessageLong(AolEdit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(AolEdit&, WM_CHAR, VK_RETURN, 0&)
    Do: DoEvents
        GotProfileWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Member Profile")
        message& = FindWindow("#32770", "America Online")
        AolView& = FindWindowEx(GotProfileWindow&, 0&, "_AOL_View", vbNullString)
    Loop Until GotProfileWindow& <> 0& And AolView& <> 0& Or message& <> 0&
    If GotProfileWindow& <> 0& Then
        Call Yield(3)
        If GetText(AolView&) = "" Then
            ProfileGet$ = Person$ & "' s profile is blank."
            Call PostMessage(GotProfileWindow&, WM_CLOSE, 0&, 0&)
            Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
            Exit Function
           ElseIf GetText(AolView&) <> "" Then
            ProfileGet$ = GetText(AolView&)
            Call PostMessage(GotProfileWindow&, WM_CLOSE, 0&, 0&)
            Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
            Exit Function
        End If
       Else
        Static1& = FindWindowEx(message&, 0&, "Static", vbNullString)
        Static2& = FindWindowEx(message&, Static1&, "Static", vbNullString)
        Button& = FindWindowEx(message&, 0&, "Button", vbNullString)
        ProfileGet$ = GetText(Static2&)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
        Exit Function
    End If
End Function

Public Sub ProfileAsPassWords(ScreenName As String, ListBox As Control)
    Dim InstanceChr As Long, ProfileString As String
    On Error Resume Next
    ProfileString$ = ReplaceCharacters(ReplaceCharacters(ReplaceCharacters(ReplaceCharacters(ProfileGet(ScreenName$), vbTab, " "), vbNullChar, " "), vbCr, " "), vbLf, "")
    If ProfileString$ = ScreenName$ & "' s profile is blank." Then Exit Sub
    For InstanceChr& = 1 To StringCount(ProfileString$, " ")
            ListBox.AddItem ProfileTrim(GetInstance(ProfileString$, " ", InstanceChr&))
    Next InstanceChr&
    Call ListRemoveNull(ListBox)
    Call ListKillDuplicates(ListBox)
End Sub

Public Function ProfileTrim(profile As String) As String
    Dim prepstring As String
    prepstring$ = ReplaceCharacters(profile$, "'", "")
    prepstring$ = ReplaceCharacters(prepstring$, "~", "")
    prepstring$ = ReplaceCharacters(prepstring$, "`", "")
    prepstring$ = ReplaceCharacters(prepstring$, "!", "")
    prepstring$ = ReplaceCharacters(prepstring$, "@", "")
    prepstring$ = ReplaceCharacters(prepstring$, "#", "")
    prepstring$ = ReplaceCharacters(prepstring$, "$", "")
    prepstring$ = ReplaceCharacters(prepstring$, "%", "")
    prepstring$ = ReplaceCharacters(prepstring$, "^", "")
    prepstring$ = ReplaceCharacters(prepstring$, "&", "")
    prepstring$ = ReplaceCharacters(prepstring$, "*", "")
    prepstring$ = ReplaceCharacters(prepstring$, "(", "")
    prepstring$ = ReplaceCharacters(prepstring$, ")", "")
    prepstring$ = ReplaceCharacters(prepstring$, "_", "")
    prepstring$ = ReplaceCharacters(prepstring$, "-", "")
    prepstring$ = ReplaceCharacters(prepstring$, "+", "")
    prepstring$ = ReplaceCharacters(prepstring$, "=", "")
    prepstring$ = ReplaceCharacters(prepstring$, "]", "")
    prepstring$ = ReplaceCharacters(prepstring$, "[", "")
    prepstring$ = ReplaceCharacters(prepstring$, "}", "")
    prepstring$ = ReplaceCharacters(prepstring$, "{", "")
    prepstring$ = ReplaceCharacters(prepstring$, Chr(34), "")
    prepstring$ = ReplaceCharacters(prepstring$, "|", "")
    prepstring$ = ReplaceCharacters(prepstring$, "\", "")
    prepstring$ = ReplaceCharacters(prepstring$, ":", "")
    prepstring$ = ReplaceCharacters(prepstring$, ";", "")
    prepstring$ = ReplaceCharacters(prepstring$, "?", "")
    prepstring$ = ReplaceCharacters(prepstring$, "/", "")
    prepstring$ = ReplaceCharacters(prepstring$, ">", "")
    prepstring$ = ReplaceCharacters(prepstring$, ".", "")
    prepstring$ = ReplaceCharacters(prepstring$, "<", "")
    prepstring$ = ReplaceCharacters(prepstring$, ",", "")
    ProfileTrim$ = prepstring$
End Function

Public Sub MailPrep(Person As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
    Dim AolFrame As Long, AolToolbar As Long, Toolbar As Long
    Dim AolIcon As Long, AolEdit1 As Long
    Dim AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim CheckBox As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolToolbar& = FindWindowEx(AolFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AolToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AolIcon& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
        RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
    Loop Until FindSendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(AolEdit3&, WM_SETTEXT, 0&, Subject$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    If CheckReturnReceipts = True Then
        CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
        Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
    End If
End Sub

Public Sub MailClickSend()
    Dim AolIcon As Long
    AolIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function ImCheck(Person As String) As String
    Dim AolFrame As Long, AolMdi As Long, ImSendWindow As Long
    Dim AolEdit As Long, AolIcon As Long
    Dim MessageOk As Long, OKButton As Long, Static1 As Long
    Dim Static2 As Long
    Call PopUpIcon(9, "I")
    Do: DoEvents
        AolFrame& = FindWindow("AOL Frame25", vbNullString)
        AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
        ImSendWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Send Instant Message")
        AolEdit& = FindWindowEx(ImSendWindow&, 0&, "_AOL_Edit", vbNullString)
        AolIcon& = NextOfClassByCount(ImSendWindow&, "_AOL_Icon", 10)
    Loop Until ImSendWindow& <> 0& And AolEdit& <> 0& And AolIcon& <> 0&
    Call SendMessageByString(AolEdit&, WM_SETTEXT, 0&, Person$)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Static1& = FindWindowEx(MessageOk&, 0&, "Static", vbNullString)
        Static2& = FindWindowEx(MessageOk&, Static1&, "Static", vbNullString)
        ImCheck$ = GetText(Static2&)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(ImSendWindow&, WM_CLOSE, 0&, 0&)
    End If
End Function

Public Sub MassImCheck(NamesList As Control)
    Dim ListIndex As Long
    On Error Resume Next
    For ListIndex& = 0& To NamesList.ListCount - 1
        NamesList.list(ListIndex&) = ImCheck(NamesList.list(ListIndex&))
        DoEvents
    Next ListIndex&
End Sub

Public Function RoomLastLineFull() As String
'My chat ocx will be out look for it
    If FindRoom& = 0& Then Exit Function
    Dim Count As Long, prepstring As String
    Count& = StringCount(RoomGetText$, vbCr)
    prepstring$ = GetInstance(RoomGetText$, vbCr, Count& + 1)
    RoomLastLineFull$ = ReplaceCharacters(prepstring$, vbTab, "    ")
End Function

Public Function RoomLastLineScreenName() As String
'My chat ocx will be out look for it
    If FindRoom& = 0& Then Exit Function
    RoomLastLineScreenName$ = Mid(RoomLastLineFull$, 1, InStr(RoomLastLineFull$, ":") - 1)
End Function

Public Function RoomLastLineMessage() As String
'My chat ocx will be out look for it
    If FindRoom& = 0& Then
        Exit Function
       ElseIf FindRoom& <> 0& Then
        RoomLastLineMessage$ = Mid(RoomLastLineFull$, InStr(RoomLastLineFull$, ":") + 6)
        Exit Function
    End If
End Function

Public Function RoomLastLineFullIndex(Optional line As Long = 0) As String
'My chat ocx will be out look for it
    If FindRoom& = 0& Then Exit Function
    Dim Count As Long, prepstring As String
    Count& = StringCount(RoomGetText$, vbCr)
    prepstring$ = GetInstance(RoomGetText$, vbCr, Count& - line&)
    RoomLastLineFullIndex$ = ReplaceCharacters(prepstring$, vbTab, "    ")
End Function

Public Function RoomLastLineScreenNameIndex(Optional line As Long = 0) As String
'My chat ocx will be out look for it
    If FindRoom& = 0& Then Exit Function
    RoomLastLineScreenNameIndex$ = Mid(RoomLastLineFullIndex$(line&), 1, InStr(RoomLastLineFullIndex$(line&), ":") - 1)
End Function

Public Function RoomLastLineMessageIndex(Optional line As Long = 0) As String
'My chat ocx will be out look for it
    If FindRoom& = 0& Then Exit Function
    RoomLastLineMessageIndex$ = Mid(RoomLastLineFullIndex$(line&), InStr(RoomLastLineFullIndex$(line&), ":") + 6)
End Function

Public Function ImScreenName() As String
    Dim AolStatic As Long
    AolStatic& = FindChildByClass(FindIm&, "_AOL_Static")
    If AolStatic& <> 0& Then
        ImScreenName$ = GetText(AolStatic&)
        Exit Function
       ElseIf AolStatic& = 0& Then
        ImScreenName$ = ""
        Exit Function
    End If
End Function

Public Function ImScreenNameFromAim() As String
    Dim ImCaption As String
    ImCaption$ = GetCaption(FindImFromAim&)
    If InStr(ImCaption$, "Instant Message From:") <> 0& Then
        ImScreenNameFromAim$ = Mid(ImCaption$, InStr(ImCaption$, ":") + 2)
    End If
End Function

Public Function ImSendRich() As Long
    Dim RichText As Long
    RichText& = FindWindowEx(FindIm&, 0&, "RICHCNTL", vbNullString)
    ImSendRich& = FindWindowEx(FindIm&, RichText&, "RICHCNTL", vbNullString)
End Function

Public Function ImGetText() As String
    Dim RichText As Long
    RichText& = FindWindowEx(FindIm&, 0&, "RICHCNTL", vbNullString)
    ImGetText$ = GetText(RichText&)
End Function

Public Function ImLastLineFull() As String
    Dim Count As Long, prepstring As String
    Count& = StringCount(ImGetText$, vbCr)
    prepstring$ = GetInstance(ImGetText$, vbCr, Count& + 1)
    ImLastLineFull$ = ReplaceCharacters(prepstring$, vbTab, "    ")
End Function

Public Function ImLastLineScreenName() As String
    ImLastLineScreenName$ = Mid(ImLastLineFull$, 1, InStr(ImLastLineFull$, ":") - 1)
    ImLastLineScreenName$ = ReplaceCharacters(ImLastLineScreenName$, Chr(10) & Chr(32), "")
End Function

Public Function ImLastLineMessage() As String
    ImLastLineMessage$ = Mid(ImLastLineFull$, InStr(ImLastLineFull$, ":") + 6)
End Function

Public Function ImLastLineFullIndex(Optional line As Long = 0) As String
    Dim Count As Long, prepstring As String
    Count& = StringCount(ImGetText$, vbCr)
    prepstring$ = GetInstance(ImGetText$, vbCr, Count& - line&)
    ImLastLineFullIndex$ = ReplaceCharacters(prepstring$, vbTab, "    ")
End Function

Public Function ImLastLineScreenNameIndex(Optional line As Long = 0) As String
    ImLastLineScreenNameIndex$ = Mid(ImLastLineFullIndex$(line&), 1, InStr(ImLastLineFullIndex$(line&), ":") - 1)
    ImLastLineScreenNameIndex$ = ReplaceCharacters(ImLastLineScreenNameIndex$(line&), Chr(10) & Chr(32), "")
End Function

Public Function ImLastLineMessageIndex(Optional line As Long = 0) As String
    ImLastLineMessageIndex$ = Mid(ImLastLineFullIndex$(line&), InStr(ImLastLineFullIndex$(line&), ":") + 6)
End Function

Public Sub ImCloseWindows()
    Do: DoEvents
        Call WinClose(FindIm&)
    Loop Until FindIm& = 0&
End Sub

Public Sub ImRespond(message As String)
    If FindIm& <> 0& And GetCaption(FindIm&) <> "Send Instant Message" Then
        Call InstantMessage(ImScreenName, message$)
        Call Yield(0.2)
        Call WinClose(FindIm&)
        Call Yield(1)
        Exit Sub
    End If
End Sub

Public Sub BuddyInvitation(Person As String, message As String, RoomOrUrl As String, Optional RoomOrWebUrl As String = "Room")
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolIcon1 As Long, AolIcon2 As Long, BuddyInviteWindow As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolEdit3 As Long
    Dim CheckBox1 As Long, CheckBox2 As Long, AolEdit4 As Long
    Dim MessageOk As Long, OKButton As Long
    If RoomOrWebUrl$ = "Room" Then
        If Len(RoomOrUrl$) > 20 Then Exit Sub
    End If
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    If FindBuddyView& = 0& Then
        Call PopUpIcon(9, "V")
    End If
    Do: DoEvents
        AolIcon1& = NextOfClassByCount(FindBuddyView&, "_AOL_Icon", 4)
    Loop Until FindBuddyView& <> 0& And AolIcon1& <> 0&
    Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        BuddyInviteWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Buddy Chat")
        AolIcon2& = FindWindowEx(BuddyInviteWindow&, 0&, "_AOL_Icon", vbNullString)
        AolEdit1& = FindWindowEx(BuddyInviteWindow&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(BuddyInviteWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolEdit3& = FindWindowEx(BuddyInviteWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
        CheckBox1& = FindWindowEx(BuddyInviteWindow&, 0&, "_AOL_Checkbox", vbNullString)
        CheckBox2& = FindWindowEx(BuddyInviteWindow&, CheckBox1&, "_AOL_Checkbox", vbNullString)
    Loop Until BuddyInviteWindow& <> 0& And AolIcon2& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And CheckBox1& <> 0& And CheckBox2& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, message$)
    If LCase(TrimSpaces(RoomOrWebUrl$)) Like LCase("Room") Then
        Call PostMessage(CheckBox1&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(CheckBox1&, WM_LBUTTONUP, 0&, 0&)
        Call SendMessageByString(AolEdit3&, WM_SETTEXT, 0&, RoomOrUrl$)
       ElseIf LCase(TrimSpaces(RoomOrWebUrl$)) Like LCase("WebUrl") Then
        Call PostMessage(CheckBox2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(CheckBox2&, WM_LBUTTONUP, 0&, 0&)
        AolEdit4& = FindWindowEx(BuddyInviteWindow&, AolEdit3&, "_AOL_Edit", vbNullString)
        Call SendMessageByString(AolEdit4&, WM_SETTEXT, 0&, RoomOrUrl$)
    End If
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
     MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWindowEx(AolMdi&, 0&, "AOL Child", "Buddy Chat") = 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(BuddyInviteWindow&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
End Sub

Public Sub BuddiesToList(ListBox As ListBox)
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
    Dim ProcessThread As Long, SearchIndex As Long
    ListHandle& = FindWindowEx(FindBuddyView&, 0&, "_AOL_Listbox", vbNullString)
    Call GetWindowThreadProcessId(ListHandle&, Process&)
    ProcessThread& = OpenProcess(Op_Flags, False, Process&)
    On Error Resume Next
    If ProcessThread& Then
        For SearchIndex& = 0 To ListCount(ListHandle&) - 1
            name$ = String(4, vbNullChar)
            ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
            ListHoldItem& = ListHoldItem& + 24
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
            Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
            ListHoldItem& = ListHoldItem& + 6
            name$ = String(16, vbNullChar)
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
            If InStr(name$, "(") = 0& Then
                ListBox.AddItem Mid(name$, 5)
            End If
        Next SearchIndex&
        Call CloseHandle(ProcessThread&)
    End If
End Sub

Public Function TimeOnline() As String
'Look for an example on how to use this sub at knk's site
    Dim AolModal As Long, AolIcon As Long, AolStatic As Long
    Call PopUpIcon(5, "O")
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", "America Online")
        AolIcon& = FindWindowEx(AolModal&, 0, "_AOL_Icon", vbNullString)
        AolStatic& = FindWindowEx(AolModal&, 0, "_AOL_Static", vbNullString)
    Loop Until AolModal& <> 0& And AolIcon& <> 0& And AolStatic& <> 0&
    TimeOnline$ = GetText(AolStatic&)
    Call PostMessage(AolIcon&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(AolIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Function MasterAccount() As Boolean
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim ParentalControlWindow As Long, AolIcon1 As Long, NextOfClass As Long
    Dim SecondParentalControlWindow As Long, AolModal As Long, AolIcon2 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    Call PopUpIcon(5, "C")
    Do: DoEvents
        ParentalControlWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", " Parental Controls")
        AolIcon1& = FindWindowEx(ParentalControlWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until ParentalControlWindow& <> 0& And AolIcon1& <> 0&
    Yield 3
    Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        SecondParentalControlWindow& = FindWindow("_AOL_Modal", "Parental Controls")
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
    Loop Until SecondParentalControlWindow& <> 0& Or AolModal& <> 0&
    If SecondParentalControlWindow& <> 0& Then
       Call PostMessage(SecondParentalControlWindow&, WM_CLOSE, 0&, 0&)
       Call PostMessage(ParentalControlWindow&, WM_CLOSE, 0&, 0&)
       MasterAccount = True
       Exit Function
      ElseIf AolModal& <> 0& Then
       AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
       Do: DoEvents
           Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
           Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
       Loop Until AolIcon2& <> 0&
       Call PostMessage(ParentalControlWindow&, WM_CLOSE, 0&, 0&)
       MasterAccount = False
       Exit Function
   End If
End Function

Public Function AolFrame() As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
End Function

Public Function AolMdi() As Long
    Dim AolFrame As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
End Function

Public Function AolChild() As Long
    Dim AolFrame As Long, AolMdi As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
End Function

Public Function AolModal() As Long
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
End Function

Public Sub MailCheckReturnReceipts(CheckReturnReceipts As Boolean)
    Dim CheckBox As Long
    If FindSendWindow& = 0& Then Exit Sub
    CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
    If CheckReturnReceipts = True Then
        Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
       ElseIf CheckReturnReceipts = False Then
        Call PostMessage(CheckBox&, BM_SETCHECK, False, 0&)
    End If
End Sub

Public Function FindSignOnScreen() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim SignOnCaption As String
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    SignOnCaption$ = GetCaption(AolChild&)
    If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
        FindSignOnScreen& = AolChild&
        Exit Function
       Else
        Do
            AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
                FindSignOnScreen& = AolChild&
                Exit Function
            End If
        Loop Until AolChild& = 0&
    End If
    FindSignOnScreen& = AolChild&
End Function

Public Sub GuestSetToGuest()
    Dim ComboBox As Long
    If FindSignOnScreen& = 0& Then Exit Sub
    ComboBox& = FindWindowEx(FindSignOnScreen&, 0&, "_AOL_Combobox", vbNullString)
    Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessageLong(ComboBox&, Cb_SetCursel, ComboCount(ComboBox&) - 1, 0&)
    Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub GuestSignOn(ScreenName As String, PassWord As String)
    If FindSignOnScreen& = 0& Then Exit Sub
    Dim MessageOk As Long, OKButton As Long, AolModal As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolIcon1 As Long
    Dim AolIcon2 As Long
    Call GuestSetToGuest
    Call GuestClickSignOn
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0 And AolIcon2& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, PassWord$)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
    If MessageOk& <> 0& Then
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
        Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, PassWord$)
        Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            MessageOk& = FindWindow("#32770", "America Online")
        Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
        If MessageOk& <> 0& And OKButton& <> 0& Then
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Exit Sub
           ElseIf FindWelcome& <> 0& Then
            Exit Sub
        End If
       ElseIf FindWelcome& <> 0& Then
        Exit Sub
    End If
End Sub

Public Sub GuestSetSnAndPw(ScreenName As String, PassWord As String)
    Dim AolModal As Long, AolEdit1 As Long, AolEdit2 As Long
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
    AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, PassWord$)
End Sub

Public Sub GuestClickSignOn()
    Dim AolIcon As Long
    If NextOfClassByCount(FindSignOnScreen&, "_AOL_Icon", 4) <> 0& Then
        AolIcon& = NextOfClassByCount(FindSignOnScreen&, "_AOL_Icon", 4)
       Else
        AolIcon& = NextOfClassByCount(FindSignOnScreen&, "_AOL_Icon", 3)
    End If
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function PhishCheck(ScreenName As String, PassWord As String, Optional SignOffIfTrue As Boolean = False) As Boolean
    If FindSignOnScreen& = 0& Then Exit Function
    Dim MessageOk As Long, OKButton As Long, AolModal As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolIcon1 As Long
    Dim AolIcon2 As Long
    Call GuestSetToGuest
    Call GuestClickSignOn
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolEdit1& = FindWindowEx(AolModal&, 0&, "_AOL_Edit", vbNullString)
        AolEdit2& = FindWindowEx(AolModal&, AolEdit1&, "_AOL_Edit", vbNullString)
        AolIcon2& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0 And AolIcon2& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(AolEdit2&, WM_SETTEXT, 0&, PassWord$)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
    If MessageOk& <> 0& Then
        If GetMessageText(MessageOk&) = "Incorrect name and/or password, please re-enter" Then
            PhishCheck = False
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Function
           ElseIf GetMessageText(MessageOk&) = "suspended" Then
            PhishCheck = False
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Function
           ElseIf InStr(GetMessageText(MessageOk&), "Your account ") = 1& Then
            PhishCheck = True
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            AolIcon2& = FindWindowEx(AolModal&, AolIcon2&, "_AOL_Icon", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
            Exit Function
        End If
       ElseIf FindWelcome& <> 0& Then
        PhishCheck = True
        If SignOffIfTrue = True Then Call RunMenuByString(AolFrame&, "&Sign Off")
        Exit Function
    End If
End Function

Public Sub PhishStatus(ListBox As ListBox, ListIndex As Long, Status As Long)
     If Status& = 1 Then
        If InStr(ListBox.list(ListIndex&), "[m]") <> 0& Then
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[s]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[s]", "[m]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[?]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[?]", "[m]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[h]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[h]", "[m]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[i]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[i]", "[m]")
            Exit Sub
        End If
       ElseIf Status& = 2 Then
        If InStr(ListBox.list(ListIndex&), "[s]") <> 0& Then
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[m]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[m]", "[s]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[?]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[?]", "[s]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[h]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[h]", "[s]")
            Exit Sub
          ElseIf InStr(ListBox.list(ListIndex&), "[i]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[i]", "[s]")
            Exit Sub
        End If
       ElseIf Status& = 3 Then
        If InStr(ListBox.list(ListIndex&), "[?]") <> 0& Then
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[m]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[m]", "[?]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[s]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[s]", "[?]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[h]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[h]", "[?]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[i]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[i]", "[?]")
            Exit Sub
        End If
       ElseIf Status& = 4 Then
        If InStr(ListBox.list(ListIndex&), "[h]") <> 0& Then
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[m]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[m]", "[h]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[s]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[s]", "[h]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[?]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[?]", "[h]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[i]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[i]", "[h]")
            Exit Sub
        End If
       ElseIf Status& = 5 Then
        If InStr(ListBox.list(ListIndex&), "[i]") <> 0& Then
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[m]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[m]", "[i]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[s]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[s]", "[i]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[?]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[?]", "[i]")
            Exit Sub
           ElseIf InStr(ListBox.list(ListIndex&), "[h]") <> 0& Then
            ListBox.list(ListIndex&) = ReplaceCharacters(ListBox.list(ListIndex&), "[h]", "[i]")
            Exit Sub
        End If
    End If
End Sub

Public Function PhishSubCount(ListBox As Control) As Long
    Dim SearchIndex As Long, CurrentCount As Long
    CurrentCount& = 0&
   For SearchIndex& = 0 To ListBox.ListCount - 1
      If InStr(ListBox.list(SearchIndex&), "[s]") <> 0& Then
          CurrentCount& = CurrentCount& + 1
      End If
   Next SearchIndex&
    PhishSubCount& = CurrentCount&
End Function

Public Function PhishMasterCount(ListBox As Control) As Long
    Dim SearchIndex As Long, CurrentCount As Long
    CurrentCount& = 0&
   For SearchIndex& = 0 To ListBox.ListCount - 1
      If InStr(ListBox.list(SearchIndex&), "[m]") <> 0& Then
          CurrentCount& = CurrentCount& + 1
      End If
   Next SearchIndex&
    PhishMasterCount& = CurrentCount&
End Function

Public Function PhishIcaseCount(ListBox As Control) As Long
    Dim SearchIndex As Long, CurrentCount As Long
    CurrentCount& = 0&
   For SearchIndex& = 0 To ListBox.ListCount - 1
      If InStr(ListBox.list(SearchIndex&), "[i]") <> 0& Then
          CurrentCount& = CurrentCount& + 1
      End If
   Next SearchIndex&
    PhishIcaseCount& = CurrentCount&
End Function

Public Function PhishHotMailCount(ListBox As Control) As Long
    Dim SearchIndex As Long, CurrentCount As Long
    CurrentCount& = 0&
   For SearchIndex& = 0 To ListBox.ListCount - 1
      If InStr(ListBox.list(SearchIndex&), "[h]") <> 0& Then
          CurrentCount& = CurrentCount& + 1
      End If
   Next SearchIndex&
    PhishHotMailCount& = CurrentCount&
End Function

Public Function PhishUnknownCount(ListBox As Control) As Long
    Dim SearchIndex As Long, CurrentCount As Long
    CurrentCount& = 0&
   For SearchIndex& = 0 To ListBox.ListCount - 1
      If InStr(ListBox.list(SearchIndex&), "[?]") <> 0& Then
          CurrentCount& = CurrentCount& + 1
      End If
   Next SearchIndex&
    PhishUnknownCount& = CurrentCount&
End Function

Public Sub KillWait()
    Dim AolModal As Long, AolIcon As Long
    Call RunMenuByString(AolFrame&, "&About America Online")
    Do: DoEvents
        AolModal& = FindWindow("_AOL_Modal", vbNullString)
        AolIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until AolModal& <> 0& And AolIcon& <> 0&
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function MailCountNew() As Long
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    MailCountNew& = ListCount(AolTree&)
End Function

Public Function MailCountOld() As Long
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
    MailCountOld& = ListCount(AolTree&)
End Function

Public Function MailCountSent() As Long
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim TabPage3 As Long, AolTree As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
    MailCountSent& = ListCount(AolTree&)
End Function

Public Function MailCountFlash() As Long
    Dim AolTree As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
    MailCountFlash& = ListCount(AolTree&)
End Function

Public Function MailSenderNew(MailIndex As Long) As String
    Dim LenSender As Long, FixedString As String, PrepSender As String
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSender& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSender& + 1, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance1& = InStr(FixedString$, vbTab)
    Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
    MailSenderNew$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function

Public Function MailSubjectNew(MailIndex As Long) As String
    Dim LenSubject As Long, FixedString As String, PrepSubject As String
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim Instance As Long, TreeCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance& = InStr(FixedString$, vbTab)
    Instance& = InStr(Instance& + 1, FixedString$, vbTab)
    FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
    MailSubjectNew$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function

Public Function MailSenderOld(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long, LenSender As Long, FixedString As String
    Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSender& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSender&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance1& = InStr(FixedString$, vbTab)
    Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
    MailSenderOld$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function

Public Function MailSubjectOld(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long, LenSubject As Long, FixedString As String
    Dim Instance As Long, TreeCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance& = InStr(FixedString$, vbTab)
    Instance& = InStr(Instance& + 1, FixedString$, vbTab)
    FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
    MailSubjectOld$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function

Public Function MailSenderSent(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long, LenSender As Long, FixedString As String
    Dim TabPage3 As Long
    Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSender& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSender&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance1& = InStr(FixedString$, vbTab)
    Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
    MailSenderSent$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function

Public Function MailSubjectSent(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long, LenSubject As Long, FixedString As String
    Dim Instance As Long, TreeCount As Long, TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance& = InStr(FixedString$, vbTab)
    Instance& = InStr(Instance& + 1, FixedString$, vbTab)
    FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
    MailSubjectSent$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function

Public Function MailSenderFlash(MailIndex As Long) As String
    Dim LenSender As Long, FixedString As String, AolTree As Long
    Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSender& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSender&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance1& = InStr(FixedString$, vbTab)
    Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
    MailSenderFlash$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function

Public Function MailSubjectFlash(MailIndex As Long) As String
    Dim LenSubject As Long, FixedString As String, PrepSubject As String
    Dim AolTree As Long, Instance As Long, TreeCount As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
    TreeCount& = ListCount(AolTree&)
   If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    Instance& = InStr(FixedString$, vbTab)
    Instance& = InStr(Instance& + 1, FixedString$, vbTab)
    FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
    MailSubjectFlash$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function

Public Function MailSenderSubjectNew(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long
    Dim AolTree As Long, LenSubject As Long, FixedString As String
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage1& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage1&, 0&, "_AOL_Tree", vbNullString)
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    MailSenderSubjectNew$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function

Public Function MailSenderSubjectOld(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long, LenSubject As Long, FixedString As String
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    MailSenderSubjectOld$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function

Public Function MailSenderSubjectSent(MailIndex As Long) As String
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long, LenSubject As Long, FixedString As String
    Dim TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Function
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    MailSenderSubjectSent$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function

Public Function MailSenderSubjectFlash(MailIndex As Long) As String
    Dim LenSubject As Long, FixedString As String, AolTree As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, MailIndex&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, MailIndex&, FixedString$)
    MailSenderSubjectFlash$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function

Public Sub MailSetPreferences(Optional ConfirmMail As Boolean = False, Optional CloseMail As Boolean = True, Optional ConfirmMarked As Boolean = False, Optional RetainSent As Boolean = False, Optional RetainRead As Boolean = False, Optional HyperLinks As Boolean = False)
    Dim MailPreferencesWindow As Long, AolIcon As Long
    Dim ConfirmMailCheckBox As Long, CloseMailCheckBox As Long, ConfirmMarkedCheckBox As Long
    Dim RetainSentCheckBox As Long, RetainReadCheckBox As Long, HyperLinksCheckBox As Long
    Call PopUpIcon(2, "P")
   Do: DoEvents
    MailPreferencesWindow& = FindWindow("_AOL_Modal", "Mail Preferences")
    ConfirmMailCheckBox& = FindWindowEx(MailPreferencesWindow&, 0&, "_AOL_Checkbox", vbNullString)
    CloseMailCheckBox& = FindWindowEx(MailPreferencesWindow&, ConfirmMailCheckBox&, "_AOL_Checkbox", vbNullString)
    ConfirmMarkedCheckBox& = FindWindowEx(MailPreferencesWindow&, CloseMailCheckBox&, "_AOL_Checkbox", vbNullString)
    RetainSentCheckBox& = FindWindowEx(MailPreferencesWindow&, ConfirmMarkedCheckBox&, "_AOL_Checkbox", vbNullString)
    RetainReadCheckBox& = FindWindowEx(MailPreferencesWindow&, RetainSentCheckBox&, "_AOL_Checkbox", vbNullString)
    HyperLinksCheckBox& = NextOfClassByCount(MailPreferencesWindow&, "_AOL_Checkbox", 8)
    AolIcon& = FindWindowEx(MailPreferencesWindow&, 0&, "_AOL_Icon", vbNullString)
   Loop Until MailPreferencesWindow& <> 0& And ConfirmMailCheckBox& <> 0& And CloseMailCheckBox& <> 0& And ConfirmMarkedCheckBox& <> 0& And RetainSentCheckBox& <> 0& And HyperLinksCheckBox& <> 0& And AolIcon& <> 0&
   If ConfirmMail = False Then
      Call CheckBoxSetValue(ConfirmMailCheckBox&, False)
     ElseIf ConfirmMail = True Then
      Call CheckBoxSetValue(ConfirmMailCheckBox&, True)
   End If
   If CloseMail = True Then
      Call CheckBoxSetValue(CloseMailCheckBox&, True)
     ElseIf ConfirmMail = False Then
      Call CheckBoxSetValue(CloseMailCheckBox&, False)
   End If
   If ConfirmMarked = False Then
      Call CheckBoxSetValue(ConfirmMarkedCheckBox&, False)
     ElseIf ConfirmMarked = True Then
      Call CheckBoxSetValue(ConfirmMarkedCheckBox&, True)
   End If
   If RetainSent = False Then
      Call CheckBoxSetValue(RetainSentCheckBox&, False)
     ElseIf RetainSent = True Then
      Call CheckBoxSetValue(RetainSentCheckBox&, True)
   End If
   If RetainRead = False Then
      Call CheckBoxSetValue(RetainReadCheckBox&, False)
     ElseIf RetainRead = True Then
      Call CheckBoxSetValue(RetainReadCheckBox&, True)
   End If
   If HyperLinks = False Then
      Call CheckBoxSetValue(HyperLinksCheckBox&, False)
     ElseIf HyperLinks = True Then
      Call CheckBoxSetValue(HyperLinksCheckBox&, True)
   End If
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailOpenNew()
    Call PopUpIcon(2, "R")
End Sub

Public Sub MailOpenOld()
    Call PopUpIcon(2, "O")
End Sub

Public Sub MailOpenSent()
    Call PopUpIcon(2, "S")
End Sub

Public Sub MailOpenFlash(Optional CheckIfGuest As Boolean = True)
   If CheckIfGuest = True Then
      If Guest = True Then
        Exit Sub
       Else
        Call PopUpIconDbl(2, "d", "I")
      End If
     ElseIf CheckIfGuest = False Then
      Call PopUpIconDbl(2, "d", "I")
   End If
End Sub

Public Sub MailToListNew(ListBox As ListBox, Optional NumberIndex As Boolean = True)
    Dim LenSubject As Long, FixedString As String, PrepSubject As String
    Dim TabControl As Long, TabPage As Long, AolTree As Long, AddThisMail As Long
    ListBox.Clear
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   For AddThisMail& = 0& To ListCount(AolTree&) - 1
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, AddThisMail&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, AddThisMail&, FixedString$)
    PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
      If NumberIndex = True Then
        ListBox.AddItem AddThisMail& + 1 & ". " & PrepSubject$
       Else
        ListBox.AddItem PrepSubject$
      End If
   Next AddThisMail&
End Sub

Public Sub MailToListOld(ListBox As ListBox, Optional NumberIndex As Boolean = True)
    Dim LenSubject As Long, FixedString As String, PrepSubject As String
    Dim TabControl As Long, TabPage As Long, AolTree As Long, AddThisMail As Long
    Dim TabPage2 As Long
    ListBox.Clear
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   For AddThisMail& = 0& To ListCount(AolTree&) - 1
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, AddThisMail&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, AddThisMail&, FixedString$)
    PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
      If NumberIndex = True Then
        ListBox.AddItem AddThisMail& + 1 & ". " & PrepSubject$
       Else
        ListBox.AddItem PrepSubject$
      End If
   Next AddThisMail&
End Sub

Public Sub MailToListSent(ListBox As ListBox, Optional NumberIndex As Boolean = True)
    Dim LenSubject As Long, FixedString As String, PrepSubject As String
    Dim TabControl As Long, TabPage As Long, AolTree As Long, AddThisMail As Long
    Dim TabPage2 As Long, TabPage3 As Long
    ListBox.Clear
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   For AddThisMail& = 0& To ListCount(AolTree&) - 1
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, AddThisMail&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, AddThisMail&, FixedString$)
    PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
      If NumberIndex = True Then
        ListBox.AddItem AddThisMail& + 1 & ". " & PrepSubject$
       Else
        ListBox.AddItem PrepSubject$
      End If
   Next AddThisMail&
End Sub

Public Sub MailToListFlash(ListBox As ListBox, Optional NumberIndex As Boolean = True)
    Dim LenSubject As Long, FixedString As String, PrepSubject As String
    Dim AolTree As Long, AddThisMail As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   For AddThisMail& = 0& To ListCount(AolTree&) - 1
    LenSubject& = SendMessageLong(AolTree&, LB_GETTEXTLEN, AddThisMail&, 0&)
    FixedString$ = String(LenSubject&, 0&)
    Call SendMessageByString(AolTree&, LB_GETTEXT, AddThisMail&, FixedString$)
    PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
      If NumberIndex = True Then
        ListBox.AddItem AddThisMail& + 1 & ". " & PrepSubject$
       Else
        ListBox.AddItem PrepSubject$
      End If
   Next AddThisMail&
End Sub

Public Sub MailOpenNewIndex(MailIndex As Long)
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    Call PostMessage(AolTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(AolTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenOldIndex(MailIndex As Long)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    Call PostMessage(AolTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(AolTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenSentIndex(MailIndex As Long)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim TabPage2 As Long, TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    Call PostMessage(AolTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(AolTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenFlashIndex(MailIndex As Long)
    Dim AolTree As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    Call PostMessage(AolTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(AolTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub

Public Sub MailOpenNewSender(MailSender As String)
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim Searchlist As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderNew(Searchlist&)) = LCase(MailSender$) Then
          Call MailOpenNewIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailOpenOldSender(MailSender As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderOld(Searchlist&)) = LCase(MailSender$) Then
          Call MailOpenOldIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailOpenSentSender(MailSender As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long, TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderSent(Searchlist&)) = LCase(MailSender$) Then
          Call MailOpenSentIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailOpenFlashSender(MailSender As String)
    Dim AolTree As Long, Searchlist As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderFlash(Searchlist&)) = LCase(MailSender$) Then
          Call MailOpenFlashIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailOpenNewSubject(MailSubject As String)
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim Searchlist As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectNew(Searchlist&)) = LCase(MailSubject$) Then
          Call MailOpenNewIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailOpenOldSubject(MailSubject As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectOld(Searchlist&)) = LCase(MailSubject$) Then
          Call MailOpenOldIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailOpenSentSubject(MailSubject As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long, TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectSent(Searchlist&)) = LCase(MailSubject$) Then
          Call MailOpenSentIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailOpenFlashSubject(MailSubject As String)
    Dim AolTree As Long, Searchlist As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectFlash(Searchlist&)) = LCase(MailSubject$) Then
          Call MailOpenFlashIndex(Searchlist&)
          Exit Sub
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteNewIndex(MailIndex As Long)
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim AolIcon As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   If MailIndex& > ListCount(AolTree&) - 1 Or MailIndex& < 0& Then Exit Sub
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    AolIcon& = NextOfClassByCount(FindMailBox&, "_AOL_Icon", 7)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailDeleteOldIndex(MailIndex As Long)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim AolIcon As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    AolIcon& = NextOfClassByCount(FindMailBox&, "_AOL_Icon", 7)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailDeleteSentIndex(MailIndex As Long)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim AolIcon As Long, TabPage2 As Long, TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    AolIcon& = NextOfClassByCount(FindMailBox&, "_AOL_Icon", 7)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub MailDeleteFlashIndex(MailIndex As Long)
    Dim AolTree As Long, AolIcon As Long, MessageOk As Long
    Dim OKButton As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessageLong(AolTree&, LB_SETCURSEL, MailIndex&, 0&)
    AolIcon& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Icon", vbNullString)
    AolIcon& = FindWindowEx(FindFlashMailBox&, AolIcon&, "_AOL_Icon", vbNullString)
    AolIcon& = FindWindowEx(FindFlashMailBox&, AolIcon&, "_AOL_Icon", vbNullString)
    AolIcon& = FindWindowEx(FindFlashMailBox&, AolIcon&, "_AOL_Icon", vbNullString)
    Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
   Do: DoEvents
    MessageOk& = FindWindow("#32770", "America Online")
    OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
   Loop Until MessageOk& <> 0& And OKButton& <> 0&
   If MessageOk& <> 0& Then
    Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
    Exit Sub
   End If
End Sub

Public Sub MailDeleteNewSender(MailSender As String)
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim Searchlist As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderNew(Searchlist&)) = LCase(MailSender$) Then
          Call MailDeleteNewIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteOldSender(MailSender As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderOld(Searchlist&)) = LCase(MailSender$) Then
          Call MailDeleteOldIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteSentSender(MailSender As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long, TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderSent(Searchlist&)) = LCase(MailSender$) Then
          Call MailDeleteSentIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteFlashSender(MailSender As String)
    Dim AolTree As Long, Searchlist As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSenderFlash(Searchlist&)) = LCase(MailSender$) Then
          Call MailDeleteFlashIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteNewSubject(MailSubject As String)
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim Searchlist As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectNew(Searchlist&)) = LCase(MailSubject$) Then
          Call MailDeleteNewIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteOldSubject(MailSubject As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectOld(Searchlist&)) = LCase(MailSubject$) Then
          Call MailDeleteOldIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteSentSubject(MailSubject As String)
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim Searchlist As Long, TabPage2 As Long, TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectSent(Searchlist&)) = LCase(MailSubject$) Then
          Call MailDeleteSentIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailDeleteFlashSubject(MailSubject As String)
    Dim AolTree As Long, Searchlist As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   For Searchlist& = 0& To ListCount(AolTree&) - 1
      If LCase(MailSubjectFlash(Searchlist&)) = LCase(MailSubject$) Then
          Call MailDeleteFlashIndex(Searchlist&)
          Searchlist& = Searchlist& - 1
      End If
   Next Searchlist&
End Sub

Public Sub MailKillDuplicatesNew()
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderNew(FirstCount&)) Like LCase(MailSenderNew(SecondCount&)) And LCase(MailSubjectNew(FirstCount&)) Like LCase(MailSubjectNew(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteNewIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesOld()
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderOld(FirstCount&)) Like LCase(MailSenderOld(SecondCount&)) And LCase(MailSubjectOld(FirstCount&)) Like LCase(MailSubjectOld(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteOldIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesSent()
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    Dim TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderSent(FirstCount&)) Like LCase(MailSenderSent(SecondCount&)) And LCase(MailSubjectSent(FirstCount&)) Like LCase(MailSubjectSent(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteSentIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesFlash()
    Dim AolTree As Long, FirstCount As Long, SecondCount As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderFlash(FirstCount&)) Like LCase(MailSenderFlash(SecondCount&)) And LCase(MailSubjectFlash(FirstCount&)) Like LCase(MailSubjectFlash(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteFlashIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesNewSender()
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderNew(FirstCount&)) Like LCase(MailSenderNew(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteNewIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesOldSender()
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderOld(FirstCount&)) Like LCase(MailSenderOld(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteOldIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesSentSender()
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    Dim TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderSent(FirstCount&)) Like LCase(MailSenderSent(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteSentIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesFlashSender()
    Dim AolTree As Long, FirstCount As Long, SecondCount As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSenderFlash(FirstCount&)) Like LCase(MailSenderFlash(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteFlashIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesNewSubject()
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSubjectNew(FirstCount&)) Like LCase(MailSubjectNew(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteNewIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesOldSubject()
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSubjectOld(FirstCount&)) Like LCase(MailSubjectOld(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteOldIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesSentSubject()
    Dim TabControl As Long, TabPage1 As Long, AolTree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    Dim TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSubjectSent(FirstCount&)) Like LCase(MailSubjectSent(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteSentIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesFlashSubject()
    Dim AolTree As Long, FirstCount As Long, SecondCount As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(AolTree&) - 1
      For SecondCount& = 0& To ListCount(AolTree&) - 1
         If LCase(MailSubjectFlash(FirstCount&)) Like LCase(MailSubjectFlash(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteFlashIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailCleanNew()
    Dim TabControl As Long, TabPage As Long, AolTree As Long
    Dim Count As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    AolTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   For Count& = 0 To ListCount(AolTree&) - 1
    Do: DoEvents
       Call MailDeleteNewIndex(Count&)
    Loop Until ListCount(AolTree&) = 0&
   Next Count&
End Sub

Public Sub MailCleanOld()
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim AolTree As Long, Count As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    AolTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   For Count& = 0 To ListCount(AolTree&) - 1
    Do: DoEvents
       Call MailDeleteOldIndex(Count&)
    Loop Until ListCount(AolTree&) = 0&
   Next Count&
End Sub

Public Sub MailCleanSent()
    Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
    Dim TabPage3 As Long, AolTree As Long, Count As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
    AolTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   For Count& = 0 To ListCount(AolTree&) - 1
    Do: DoEvents
       Call MailDeleteSentIndex(Count&)
    Loop Until ListCount(AolTree&) = 0&
   Next Count&
End Sub

Public Sub MailCleanFlash()
    Dim AolTree As Long, Count As Long
    AolTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
   For Count& = 0 To ListCount(AolTree&) - 1
    Do: DoEvents
       Call MailDeleteFlashIndex(Count&)
    Loop Until ListCount(AolTree&) = 0&
   Next Count&
End Sub

Public Function MailTosCheck(ScreenName As String) As String
'Aka Alive Check
    Dim ErrorWindow As Long, AolView As Long, ViewText As String
    Dim MessageOk As Long, OKButton As Long
    Call MailSendNoKill("*, " & ScreenName$, "Tos check.", "")
    Do: DoEvents
        ErrorWindow& = FindErrorWindow&
        AolView& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ViewText$ = GetText(AolView&)
    Loop Until ErrorWindow& <> 0& And AolView& <> 0 And ViewText$ <> ""
    If InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(ScreenName$ & " - This is not a known member."))) <> 0& Then
        MailTosCheck$ = "invalid"
       ElseIf InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(ScreenName$ & " - This member is currently not accepting e-mail from your account."))) <> 0& Then
        MailTosCheck$ = "valid, no mail"
       ElseIf InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(ScreenName$ & " - This member is currently not accepting e-mail attachments or embedded files."))) <> 0& Then
        MailTosCheck$ = "valid, no attached files"
       ElseIf InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(ScreenName$ & " - This member's mailbox is full."))) <> 0& Then
        MailTosCheck$ = "valid, full mailbox"
       ElseIf Len(ScreenName$) > 16 Then
        MailTosCheck$ = "invalid length"
       Else
        MailTosCheck$ = "valid"
    End If
    Call PostMessage(FindErrorWindow&, WM_CLOSE, 0&, 0&)
    Call PostMessage(FindSendWindow&, WM_CLOSE, 0&, 0&)
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", "&No")
    Loop Until MessageOk& <> 0& And OKButton& <> 0&
    Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Sub MailMassTosCheck(NamesList As Control)
'Check more then one sn at a time
    Dim ListIndex As Long
    On Error Resume Next
    For ListIndex& = 0& To NamesList.ListCount - 1
        NamesList.list(ListIndex&) = NamesList.list(ListIndex&) & ": " & MailTosCheck(NamesList.list(ListIndex&))
        DoEvents
    Next ListIndex&
End Sub

Public Sub MailForwardNew(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
    Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim SendIcon As Long, AolModal As Long, ModalIcon As Long
    Call MailOpenNewIndex(MailIndex&)
   Do: DoEvents
      ForwardWindow& = FindForwardWindow&
   Loop Until ForwardWindow& <> 0&
    ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
    Call ClickIcon(ForwardIcon&)
   Do: DoEvents
      SendWindow& = FindSendWindow&
      AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
      AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
      AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
      RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
   Loop Until SendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
    If TrimFwd = True Then Call MailRemoveFwd
    SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
    Call ClickIcon(SendIcon&)
   Do: DoEvents
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    ModalIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
   Loop Until (AolModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
      If AolModal& <> 0& Then
         Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
         Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
         Call WinClose(FindForwardWindow&)
         Exit Sub
        ElseIf AolModal& = 0& Then
         Call WinClose(FindForwardWindow&)
         Exit Sub
      End If
End Sub

Public Sub MailForwardOld(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
    Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim SendIcon As Long, AolModal As Long, ModalIcon As Long
    Call MailOpenOldIndex(MailIndex&)
   Do: DoEvents
      ForwardWindow& = FindForwardWindow&
   Loop Until ForwardWindow& <> 0&
    ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
    Call ClickIcon(ForwardIcon&)
   Do: DoEvents
      SendWindow& = FindSendWindow&
      AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
      AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
      AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
      RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
   Loop Until SendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
    If TrimFwd = True Then Call MailRemoveFwd
    SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
    Call ClickIcon(SendIcon&)
   Do: DoEvents
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    ModalIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
   Loop Until (AolModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
      If AolModal& <> 0& Then
         Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
         Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
         Call WinClose(FindForwardWindow&)
         Exit Sub
        ElseIf AolModal& = 0& Then
         Call WinClose(FindForwardWindow&)
         Exit Sub
      End If
End Sub

Public Sub MailForwardSent(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
    Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim SendIcon As Long, AolModal As Long, ModalIcon As Long
    Call MailOpenSentIndex(MailIndex&)
   Do: DoEvents
      ForwardWindow& = FindForwardWindow&
   Loop Until ForwardWindow& <> 0&
    ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
    Call ClickIcon(ForwardIcon&)
   Do: DoEvents
      SendWindow& = FindSendWindow&
      AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
      AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
      AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
      RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
   Loop Until SendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
    If TrimFwd = True Then Call MailRemoveFwd
    SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
    Call ClickIcon(SendIcon&)
   Do: DoEvents
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    ModalIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
   Loop Until (AolModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
      If AolModal& <> 0& Then
         Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
         Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
         Call WinClose(FindForwardWindow&)
         Exit Sub
        ElseIf AolModal& = 0& Then
         Call WinClose(FindForwardWindow&)
         Exit Sub
      End If
End Sub

Public Sub MailForwardFlash(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
    Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
    Dim AolEdit1 As Long, AolEdit2 As Long, AolEdit3 As Long, RichText As Long
    Dim SendIcon As Long, AolModal As Long, ModalIcon As Long
    Call MailOpenFlashIndex(MailIndex&)
   Do: DoEvents
      ForwardWindow& = FindForwardWindow&
   Loop Until ForwardWindow& <> 0&
    ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
    Call ClickIcon(ForwardIcon&)
   Do: DoEvents
      SendWindow& = FindSendWindow&
      AolEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
      AolEdit2& = FindWindowEx(FindSendWindow&, AolEdit1&, "_AOL_Edit", vbNullString)
      AolEdit3& = FindWindowEx(FindSendWindow&, AolEdit2&, "_AOL_Edit", vbNullString)
      RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
   Loop Until SendWindow& <> 0& And AolEdit1& <> 0& And AolEdit2& <> 0& And AolEdit3& <> 0& And RichText& <> 0&
    Call SendMessageByString(AolEdit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
    If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
    If TrimFwd = True Then Call MailRemoveFwd
    SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
    Call ClickIcon(SendIcon&)
   Do: DoEvents
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    ModalIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
   Loop Until (AolModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
      If AolModal& <> 0& Then
         Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
         Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
         Call WinClose(FindForwardWindow&)
         Exit Sub
        ElseIf AolModal& = 0& Then
         Call WinClose(FindForwardWindow&)
         Exit Sub
      End If
End Sub

Public Sub MailBomb(ScreenName As String, Subject As String, message As String, Optional MaxBombs As Long = "100", Optional Delay As Single = "2")
    Dim AolIcon As Long, MessageOk As Long, OkButton1 As Long
    Dim Bomb As Long, OkButton2 As Long
    Call MailSetPreferences(, False)
    Call MailPrep(ScreenName$, Subject$, message$)
    Do: DoEvents
        AolIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
    Loop Until FindSendWindow& <> 0& And AolIcon& <> 0&
    For Bomb& = 1 To MaxBombs&
        Call PostMessage(AolIcon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon&, WM_LBUTTONUP, 0&, 0&)
        If Bomb& >= MaxBombs& Then
            Call MailSetPreferences(, True)
            Call WinClose(FindWindowEx(AolMdi&, 0&, "AOL Child", "Write Mail"))
            Do: DoEvents
                MessageOk& = FindWindow("#32770", "America Online")
                OkButton1& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
                OkButton2& = FindWindowEx(MessageOk&, OkButton1&, "Button", vbNullString)
            Loop Until MessageOk& <> 0& And OkButton1& <> 0& And OkButton2& <> 0&
            If MessageOk& <> 0& Then
                Call PostMessage(OkButton2&, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(OkButton2&, WM_KEYUP, VK_SPACE, 0&)
                Exit Sub
               ElseIf MessageOk& = 0& And FindWindowEx(AolMdi&, 0&, "AOL Child", "Write Mail") = 0& Then
                Exit Sub
            End If
        End If
        Call Yield(Val(Delay))
    Next Bomb&
End Sub

Public Function MailListString(ListBox As ListBox, SearchString As String) As String
     Dim Search As Long, prepstring As String
     For Search& = 0 To ListBox.ListCount - 1
         If InStr(LCase(ListBox.list(Search&)), LCase(SearchString$)) <> 0& Then
             prepstring$ = prepstring$ & vbCrLf & ListBox.list(Search&)
         End If
     Next Search&
     MailListString$ = prepstring$
End Function

Public Function MailListStringTwo(ListBox As Control, Optional NumberIndex As Boolean = True) As String
    Dim CurrentCount As Long, prepstring As String
    For CurrentCount& = 0 To ListBox.ListCount - 1
        If NumberIndex = True Then
            prepstring$ = prepstring$ & CurrentCount& + 1 & "." & ListBox.list(CurrentCount&) & vbCrLf
        End If
    Next CurrentCount&
    MailListStringTwo$ = prepstring$
End Function

Public Function MailCountFromSnNew(ScreenName As String) As Long
    Dim MailIndex As Long, PrepCount As Long
    MailCountFromSnNew& = 0
    For MailIndex& = 0 To MailCountNew - 1
        If LCase(TrimSpaces(MailSenderNew(MailIndex&))) = LCase(TrimSpaces(ScreenName$)) Then
            PrepCount& = Val(PrepCount&) + 1
        End If
    Next MailIndex&
    MailCountFromSnNew& = PrepCount&
End Function

Public Function MailCountFromSnOld(ScreenName As String) As Long
    Dim MailIndex As Long, PrepCount As Long
    MailCountFromSnOld& = 0
    For MailIndex& = 0 To MailCountOld - 1
        If LCase(TrimSpaces(MailSenderOld(MailIndex&))) = LCase(TrimSpaces(ScreenName$)) Then
            PrepCount& = Val(PrepCount&) + 1
        End If
    Next MailIndex&
    MailCountFromSnOld& = PrepCount&
End Function

Public Function MailCountFromSnSent(ScreenName As String) As Long
    Dim MailIndex As Long, PrepCount As Long
    MailCountFromSnSent& = 0
    For MailIndex& = 0 To MailCountSent - 1
        If LCase(TrimSpaces(MailSenderSent(MailIndex&))) = LCase(TrimSpaces(ScreenName$)) Then
            PrepCount& = Val(PrepCount&) + 1
        End If
    Next MailIndex&
    MailCountFromSnSent& = PrepCount&
End Function

Public Function MailCountFromSnFlash(ScreenName As String) As Long
    Dim MailIndex As Long, PrepCount As Long
    MailCountFromSnFlash& = 0
    For MailIndex& = 0 To MailCountFlash - 1
        If LCase(TrimSpaces(MailSenderFlash(MailIndex&))) = LCase(TrimSpaces(ScreenName$)) Then
            PrepCount& = Val(PrepCount&) + 1
        End If
    Next MailIndex&
    MailCountFromSnFlash& = PrepCount&
End Function

Public Function RoomForceEnter(AolKeyWord As String, PrivateRoom As String, Optional CloseChatBeforeBust As Boolean = True, Optional Delay As Single = ".5", Optional StopAfterSoManyTries As Long = "1") As Long
   'Not what i made hookshot with but its still quite nice
    Dim MessageOk As Long, OKButton As Long, GetRoomName As String
   If CloseChatBeforeBust = True Then
       If FindRoom& <> 0& Then Call PostMessage(FindRoom&, WM_CLOSE, 0&, 0&)
        RoomForceEnter& = 0&
       Do: DoEvents
        Call KeyWord(AolKeyWord$ & PrivateRoom$)
         RoomForceEnter& = RoomForceEnter& + 1
           Do: DoEvents
            MessageOk& = FindWindow("#32770", "America Online")
           Loop Until MessageOk& <> 0& Or FindRoom& <> 0&
             If MessageOk& <> 0& Then
               OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
               Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
             End If
           Yield Val(Delay)
        If RoomForceEnter& >= StopAfterSoManyTries& Then Exit Do
       Loop Until FindRoom& <> 0&
       Label12.Caption = List2.ListCount
            Exit Function
      ElseIf CloseChatBeforeBust = False Then
        If FindRoom& <> 0& Then
         GetRoomName$ = LCase(TrimSpaces(RoomName$))
       Do: DoEvents
        Call KeyWord(AolKeyWord$ & PrivateRoom$)
         RoomForceEnter& = RoomForceEnter& + 1
           Do: DoEvents
            MessageOk& = FindWindow("#32770", "America Online")
           Loop Until MessageOk& <> 0& Or InStr(LCase(TrimSpaces(RoomName$)), LCase(TrimSpaces(PrivateRoom$))) <> 0&
             If MessageOk& <> 0& Then
               OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
               Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
             End If
           Yield Val(Delay)
        If RoomForceEnter& >= StopAfterSoManyTries& Then Exit Do
       Loop Until InStr(LCase(TrimSpaces(RoomName$)), LCase(TrimSpaces(PrivateRoom$))) <> 0&
        Exit Function
        ElseIf FindRoom& = 0& Then
       Do: DoEvents
        Call KeyWord(AolKeyWord$ & PrivateRoom$)
         RoomForceEnter& = RoomForceEnter& + 1
           Do: DoEvents
            MessageOk& = FindWindow("#32770", "America Online")
           Loop Until MessageOk& <> 0& Or FindRoom& <> 0&
             If MessageOk& <> 0& Then
               OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
               Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
               Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
             End If
           Yield Val(Delay)
        If RoomForceEnter& >= StopAfterSoManyTries& Then Exit Do
       Loop Until FindRoom& <> 0&
        Exit Function
        End If
      End If
End Function

Public Sub RoomSetPreferences(MembersArrive As Boolean, MembersLeave As Boolean, DoubleSpace As Boolean, Alphabatize As Boolean, Sounds As Boolean)
    Dim PreferencesWindow As Long, ChatPreferencesWindow As Long, AolIcon1 As Long
    Dim AolIcon2 As Long, MembersArriveCheckBox As Long, MembersLeaveCheckBox As Long
    Dim DoubleSpaceCheckBox As Long, AlphabatizeCheckBox As Long, SoundsCheckBox As Long
    Dim MessageOk As Long, OKButton As Long
    If FindRoom& = 0& Then
        Call PopUpIcon(5, "P")
        Do: DoEvents
            PreferencesWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Preferences")
            AolIcon1& = NextOfClassByCount(PreferencesWindow&, "_AOL_Icon", 5)
        Loop Until PreferencesWindow& <> 0& And AolIcon1& <> 0&
        Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
       ElseIf FindRoom& <> 0& Then
        AolIcon1& = NextOfClassByCount(FindRoom&, "_AOL_Icon", 10)
        Call PostMessage(AolIcon1&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(AolIcon1&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        ChatPreferencesWindow& = FindWindow("_AOL_Modal", "Chat Preferences")
        MembersArriveCheckBox& = FindWindowEx(ChatPreferencesWindow&, 0&, "_AOL_Checkbox", vbNullString)
        MembersLeaveCheckBox& = FindWindowEx(ChatPreferencesWindow&, MembersArriveCheckBox&, "_AOL_Checkbox", vbNullString)
        DoubleSpaceCheckBox& = FindWindowEx(ChatPreferencesWindow&, MembersLeaveCheckBox&, "_AOL_Checkbox", vbNullString)
        AlphabatizeCheckBox& = FindWindowEx(ChatPreferencesWindow&, DoubleSpaceCheckBox&, "_AOL_Checkbox", vbNullString)
        SoundsCheckBox& = FindWindowEx(ChatPreferencesWindow&, AlphabatizeCheckBox&, "_AOL_Checkbox", vbNullString)
        AolIcon2& = FindWindowEx(ChatPreferencesWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until ChatPreferencesWindow& <> 0& And MembersArriveCheckBox& <> 0& And MembersLeaveCheckBox& <> 0& And DoubleSpaceCheckBox& <> 0& And AlphabatizeCheckBox& <> 0& And SoundsCheckBox& <> 0& And AolIcon2& <> 0&
    Call CheckBoxSetValue(MembersArriveCheckBox&, MembersArrive)
    Call CheckBoxSetValue(MembersLeaveCheckBox&, MembersLeave)
    Call CheckBoxSetValue(DoubleSpaceCheckBox&, DoubleSpace)
    Call CheckBoxSetValue(AlphabatizeCheckBox&, Alphabatize)
    Call CheckBoxSetValue(SoundsCheckBox&, Sounds)
    Call PostMessage(AolIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AolIcon2&, WM_LBUTTONUP, 0&, 0&)
    If PreferencesWindow& <> 0& Then
        Call PostMessage(PreferencesWindow&, WM_CLOSE, 0&, 0&)
    End If
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
    Loop Until MessageOk& <> 0& And OKButton& <> 0& Or FindWindow("_AOL_Modal", "Chat Preferences") = 0&
    If MessageOk& <> 0& Then
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Exit Sub
    End If
End Sub

Public Sub WaitForOkorRoom(PrivateRoom As String)
    Dim MessageOk As Long, OKButton As Long, RoomCaption As String
    PrivateRoom$ = LCase(TrimSpaces(PrivateRoom$))
    Do: DoEvents
        RoomCaption$ = LCase(TrimSpaces(GetCaption(FindRoom&)))
        MessageOk& = FindWindow("#32770", "America Online")
    Loop Until MessageOk& <> 0& Or RoomCaption$ = PrivateRoom$
    If MessageOk& <> 0& Then
        Do: DoEvents
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Loop Until MessageOk& = 0& Or OKButton& = 0&
    End If
End Sub

Public Sub RoomGreeter(Optional MessageGreet As String = "Welcome, ")
    Dim ErrorHandle, SavedEntries As String
    On Error GoTo ErrorHandle
    If RoomLastLineScreenName$ = "OnlineHost" Then
        If InStr(RoomLastLineMessage$, " has entered the room") <> 0& Then
            If StringCount(SavedEntries$, Mid(RoomLastLineMessage$, 1, InStr(RoomLastLineMessage$, " has"))) >= 1 Then Exit Sub
            RoomSend MessageGreet$ & Mid(RoomLastLineMessage$, 1, InStr(RoomLastLineMessage$, " has"))
            SavedEntries$ = SavedEntries$ & RoomLastLineScreenName$
            Yield 1
            Exit Sub
        End If
    End If
ErrorHandle:
End Sub

Public Function FindOpenMail() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolStatic1 As Long, AolStatic2 As Long, AolStatic3 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
    AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
   If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& Then
        If GetText(AolStatic3&) = "Reply" Then
            FindOpenMail& = AolChild&
            Exit Function
        End If
    Else
       Do
         AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         AolStatic1& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
         AolStatic2& = FindWindowEx(AolChild&, AolStatic1&, "_AOL_Static", vbNullString)
         AolStatic3& = FindWindowEx(AolChild&, AolStatic2&, "_AOL_Static", vbNullString)
          If AolStatic1& <> 0& And AolStatic2& <> 0& And AolStatic3& <> 0& Then
           If GetText(AolStatic3&) = "Reply" Then
               FindOpenMail& = AolChild&
               Exit Function
           End If
          End If
       Loop Until AolChild& = 0&
   End If
    FindOpenMail& = AolChild&
End Function

Public Sub MailCloseWindows()
  If FindSendWindow& = 0& And FindFwdWindow& = 0& And FindReWindow& = 0& And FindOpenMail& = 0& And FindForwardWindow& = 0& Then Exit Sub
   Do: DoEvents
    If FindSendWindow& <> 0& Then
        Do: DoEvents
           Call PostMessage(FindSendWindow&, WM_CLOSE, 0&, 0&)
        Loop Until FindSendWindow& = 0&
      ElseIf FindFwdWindow& <> 0& Then
        Do: DoEvents
           Call PostMessage(FindFwdWindow&, WM_CLOSE, 0&, 0&)
        Loop Until FindFwdWindow& = 0&
      ElseIf FindForwardWindow& <> 0& Then
        Do: DoEvents
           Call PostMessage(FindForwardWindow&, WM_CLOSE, 0&, 0&)
        Loop Until FindForwardWindow& = 0&
      ElseIf FindReWindow& <> 0& Then
        Do: DoEvents
           Call PostMessage(FindReWindow&, WM_CLOSE, 0&, 0&)
        Loop Until FindReWindow& = 0&
      ElseIf FindOpenMail& <> 0& Then
        Do: DoEvents
           Call PostMessage(FindOpenMail&, WM_CLOSE, 0&, 0&)
        Loop Until FindOpenMail& = 0&
    End If
   Loop Until FindForwardWindow& = 0& And FindSendWindow& = 0& And FindFwdWindow& = 0& And FindReWindow& = 0& And FindOpenMail& = 0&
End Sub

Public Function FindSwitchWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If GetCaption(AolChild&) = "Switch Screen Names" Then
        FindSwitchWindow& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         If GetCaption(AolChild&) = "Switch Screen Names" Then
             FindSwitchWindow& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindSwitchWindow& = AolChild&
End Function

Public Function FindAboutWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolStatic As Long, CheckBox As Long, AolGlyph As Long
    Dim AolIcon1 As Long, AolIcon2 As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    AolIcon1& = FindWindowEx(AolChild&, 0&, "_AOL_Icon", vbNullString)
    AolIcon2& = FindWindowEx(AolChild&, AolIcon1&, "_AOL_Icon", vbNullString)
    AolGlyph& = FindWindowEx(AolChild&, 0&, "_AOL_Glyph", vbNullString)
    CheckBox& = FindWindowEx(AolChild&, 0&, "_AOL_Checkbox", vbNullString)
  If AolStatic& <> 0& And AolIcon1& <> 0& And AolIcon2& <> 0& And AolGlyph& <> 0& And CheckBox& <> 0& Then
        FindAboutWindow& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
        AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
        AolIcon1& = FindWindowEx(AolChild&, 0&, "_AOL_Icon", vbNullString)
        AolIcon2& = FindWindowEx(AolChild&, AolIcon1&, "_AOL_Icon", vbNullString)
        AolGlyph& = FindWindowEx(AolChild&, 0&, "_AOL_Glyph", vbNullString)
        CheckBox& = FindWindowEx(AolChild&, 0&, "_AOL_Checkbox", vbNullString)
         If AolStatic& <> 0& And AolIcon1& <> 0& And AolIcon2& <> 0& And AolGlyph& <> 0& And CheckBox& <> 0& Then
            FindAboutWindow& = AolChild&
            Exit Function
         End If
      Loop Until AolChild& = 0&
  End If
    FindAboutWindow& = AolChild&
End Function

Public Function FindMailStatusWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
   If InStr(GetCaption(AolChild&), "Status of ") <> 0& And GetCaption(AolChild&) <> "Locate Member Online" Then
        FindMailStatusWindow& = AolChild&
        Exit Function
    Else
      Do
        AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
         If InStr(GetCaption(AolChild&), "Status of ") <> 0& And GetCaption(AolChild&) <> "Locate Member Online" Then
             FindMailStatusWindow& = AolChild&
             Exit Function
         End If
      Loop Until AolChild& = 0&
   End If
    FindMailStatusWindow& = AolChild&
End Function

Public Sub RoomIgnoreByIndex(ListIndex As Long, Optional IgnoreOrUnignore As Boolean = True)
    Dim RoomList As Long, AboutWindow As Long, CheckBox As Long
    Dim CheckValue As Boolean
    RoomList& = FindWindowEx(FindRoom&, 0&, "_AOL_listbox", vbNullString)
    Call SendMessageLong(RoomList&, LB_SETCURSEL, ListIndex&, 0&)
    Call PostMessage(RoomList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do: DoEvents
        AboutWindow& = FindAboutWindow&
        CheckBox& = FindWindowEx(AboutWindow&, 0&, "_AOL_Checkbox", vbNullString)
    Loop Until AboutWindow& <> 0& And CheckBox& <> 0&
    If IgnoreOrUnignore = True Then
        Do: DoEvents
            CheckValue = CheckBoxGetValue(CheckBox&)
            DoEvents
            Call PostMessage(CheckBox&, WM_LBUTTONDOWN, 0&, 0&)
            DoEvents
            Call PostMessage(CheckBox&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        Loop Until CheckValue = True
       ElseIf IgnoreOrUnignore = False Then
        Do: DoEvents
            CheckValue = CheckBoxGetValue(CheckBox&)
            DoEvents
            Call PostMessage(CheckBox&, WM_LBUTTONDOWN, 0&, 0&)
            DoEvents
            Call PostMessage(CheckBox&, WM_LBUTTONUP, 0&, 0&)
            DoEvents
        Loop Until CheckValue = False
    End If
    DoEvents
    Call PostMessage(AboutWindow&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub RoomIgnoreByScreenName(ScreenName As String, Optional IgnoreOrUnignore As Boolean = True)
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
    Dim ProcessThread As Long, SearchIndex As Long
    ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    Call GetWindowThreadProcessId(ListHandle&, Process&)
    ProcessThread& = OpenProcess(Op_Flags, False, Process&)
    If ProcessThread& Then
        For SearchIndex& = 0 To ListCount(ListHandle&) - 1
            name$ = String(4, vbNullChar)
            ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
            ListHoldItem& = ListHoldItem& + 24
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
            Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
            ListHoldItem& = ListHoldItem& + 6
            name$ = String(16, vbNullChar)
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
            name$ = Left(name$, InStr(name$, vbNullChar) - 1)
                If LCase(TrimSpaces(name$)) <> LCase(TrimSpaces(UserSN$)) And LCase(TrimSpaces(name$)) = LCase(TrimSpaces(ScreenName$)) Then
                    SearchIndex& = SearchIndex&
                    Call RoomIgnoreByIndex(SearchIndex&, IgnoreOrUnignore)
                    Exit Sub
                End If
        Next SearchIndex&
        Call CloseHandle(ProcessThread&)
    End If
End Sub

Public Sub AddRoomToList(ListBox As ListBox, Optional AddUserSn As Boolean = False)
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
    Dim ProcessThread As Long, SearchIndex As Long
    ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    Call GetWindowThreadProcessId(ListHandle&, Process&)
    ProcessThread& = OpenProcess(Op_Flags, False, Process&)
    If ProcessThread& Then
        For SearchIndex& = 0 To ListCount(ListHandle&) - 1
            name$ = String(4, vbNullChar)
            ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
            ListHoldItem& = ListHoldItem& + 24
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
            Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
            ListHoldItem& = ListHoldItem& + 6
            name$ = String(16, vbNullChar)
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
            name$ = Left(name$, InStr(name$, vbNullChar) - 1)
                If AddUserSn = True Then
                    ListBox.AddItem name$
                   ElseIf AddUserSn = False Then
                    If name$ <> UserSN$ Then
                        ListBox.AddItem name$
                    End If
                End If
        Next SearchIndex&
        Call CloseHandle(ProcessThread&)
    End If
End Sub

Public Sub AddWhosChatting(ListBox As Control)
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, SearchIndex As Long
    Dim ProcessThread As Long, ListHandle3 As Long, WhosChatting As Long
    Do: DoEvents
        WhosChatting& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Who's Chatting")
        ListHandle3& = FindWindowEx(WhosChatting&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until WhosChatting& <> 0& And ListHandle3& <> 0&
    Call GetWindowThreadProcessId(ListHandle3&, Process&)
    ProcessThread& = OpenProcess(Op_Flags, False, Process&)
        If ProcessThread& Then
            For SearchIndex& = 0 To ListCount(ListHandle3&) - 1
                name$ = String(4, vbNullChar)
                ListHoldItem& = SendMessage(ListHandle3&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
                ListHoldItem& = ListHoldItem& + 24
                Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
                Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
                ListHoldItem& = ListHoldItem& + 6
                name$ = String(16, vbNullChar)
                Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
                name$ = Left(name$, InStr(name$, vbNullChar) - 1)
                    If name$ <> UserSN$ Then
                        ListBox.AddItem (name$)
                    End If
            Next SearchIndex&
    Call CloseHandle(ProcessThread&)
        End If
End Sub

Public Sub AutoListerFindAChat(ListBox As Control, Optional Limit As Long = "100")
'Ever used auto gather in inertia, well here is a sub much like the one monk had to make
'You can set the number of people to however many you want, my default is set to 100
'Great sub for spammers
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle1 As Long
    Dim ProcessThread As Long, SearchIndex As Long, ListHandle2 As Long
    Dim ListHandle3 As Long, SetIndex As Long, WhosChatting As Long
    Dim AolIcon As Long, Current As Long
    If FindAChatWindow& = 0& Then
        Call PopUpIcon(9, "F")
        Do: DoEvents
            ListHandle1& = FindWindowEx(FindAChatWindow&, 0&, "_AOL_Listbox", vbNullString)
            ListHandle2& = FindWindowEx(FindAChatWindow&, ListHandle1&, "_AOL_Listbox", vbNullString)
        Loop Until FindAChatWindow& <> 0& And ListHandle1& <> 0& And ListHandle2& <> 0&
        Call Yield(4)
    End If
    ListHandle1& = FindWindowEx(FindAChatWindow&, 0&, "_AOL_Listbox", vbNullString)
    ListHandle2& = FindWindowEx(FindAChatWindow&, ListHandle1&, "_AOL_Listbox", vbNullString)
    AolIcon& = NextOfClassByCount(FindAChatWindow&, "_AOL_Icon", 9)
    DoEvents
    For SetIndex& = 0 To ListCount(ListHandle2&) - 2
        Call ListSetFocus(ListHandle2&, SetIndex&)
        Call ClickIcon(AolIcon&)
        Do: DoEvents
            WhosChatting& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Who's Chatting")
            ListHandle3& = FindWindowEx(WhosChatting&, 0&, "_AOL_Listbox", vbNullString)
        Loop Until WhosChatting& <> 0& And ListHandle3& <> 0&
        Yield 1.5
        Call GetWindowThreadProcessId(ListHandle3&, Process&)
        ProcessThread& = OpenProcess(Op_Flags, False, Process&)
        If ProcessThread& Then
            For SearchIndex& = 0 To ListCount(ListHandle3&) - 1
                name$ = String(4, vbNullChar)
                ListHoldItem& = SendMessage(ListHandle3&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
                ListHoldItem& = ListHoldItem& + 24
                Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
                Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
                ListHoldItem& = ListHoldItem& + 6
                name$ = String(16, vbNullChar)
                Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
                name$ = Left(name$, InStr(name$, vbNullChar) - 1)
                If name$ <> UserSN$ Then
                    ListBox.AddItem TrimSpaces(name$)
                    Current& = Current& + 1
                    If Current& >= Limit& Then
                        Call WinClose(WhosChatting&)
                        Call WinClose(FindAChatWindow&)
                        Exit Sub
                    End If
                End If
                DoEvents
            Next SearchIndex&
            Call CloseHandle(ProcessThread&)
        End If
        Call Yield(0.5)
        Call WinClose(WhosChatting&)
        Yield 5
    Next SetIndex&
End Sub

Public Sub AutoListerLobbyChat(ListBox As Control, Optional Limit As Long = "100")
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
    Dim ProcessThread As Long, SearchIndex As Long, ChatRoom As Long
    Dim Current As Long
    If FindRoom& <> 0& Then Call WinClose(FindRoom&)
        Do: DoEvents
            Call PopUpIcon(9, "C")
            Do: DoEvents
                ChatRoom& = FindRoom&
            Loop Until ChatRoom& <> 0&
            DoEvents
            Yield 1.3
            ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
            Call GetWindowThreadProcessId(ListHandle&, Process&)
            ProcessThread& = OpenProcess(Op_Flags, False, Process&)
            If ProcessThread& Then
                For SearchIndex& = 0 To ListCount(ListHandle&) - 1
                    name$ = String(4, vbNullChar)
                    ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
                    ListHoldItem& = ListHoldItem& + 24
                    Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
                    Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
                    ListHoldItem& = ListHoldItem& + 6
                    name$ = String(16, vbNullChar)
                    Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
                    name$ = Left(name$, InStr(name$, vbNullChar) - 1)
                    If name$ <> UserSN$ Then
                    ListBox.AddItem name$
                    Current& = Current& + 1
                    If Current& >= Limit& Then Exit Sub
                    End If
                Next SearchIndex&
                Call CloseHandle(ProcessThread&)
            End If
            DoEvents
            Call WinClose(FindRoom&)
            Yield 5
        Loop
End Sub

Public Function RoomSearch(ScreenName As String) As Boolean
    Dim Process As Long, ListHoldItem As Long, name As String
    Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
    Dim ProcessThread As Long, SearchIndex As Long
    ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    Call GetWindowThreadProcessId(ListHandle&, Process&)
    ProcessThread& = OpenProcess(Op_Flags, False, Process&)
    If ProcessThread& Then
        For SearchIndex& = 0 To ListCount(ListHandle&) - 1
            name$ = String(4, vbNullChar)
            ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
            ListHoldItem& = ListHoldItem& + 24
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, 4, BytesRead&)
            Call RtlMoveMemory(ListHoldItem&, ByVal name$, 4)
            ListHoldItem& = ListHoldItem& + 6
            name$ = String(16, vbNullChar)
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, name$, Len(name$), BytesRead&)
            name$ = Left(name$, InStr(name$, vbNullChar) - 1)
                If LCase(TrimSpaces(ScreenName$)) Like LCase(TrimSpaces(name$)) Then
                    RoomSearch = True
                    Exit Function
                End If
        Next SearchIndex&
        Call CloseHandle(ProcessThread&)
    End If
End Function

Public Function CheckEmpowered() As Boolean
'Got Emp?
    Dim UserIm As Long, MessageOk As Long, OKButton As Long
    Call ImsOff
    Call InstantMessageFast(UserSN, "emp check")
    Do: DoEvents
        MessageOk& = FindWindow("#32770", "America Online")
        UserIm& = FindWindowEx(AolMdi&, 0&, "AOL Child", "  Instant Message To: " & UserSN)
    Loop Until MessageOk& <> 0& Or UserIm& <> 0& Or FindWindowEx(AolMdi&, 0&, "AOL Child", "Send Instant Message") = 0&
    If MessageOk& <> 0& Then
        CheckEmpowered = False
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(FindWindowEx(AolMdi&, 0&, "AOL Child", "Send Instant Message"), WM_CLOSE, 0&, 0&)
        Exit Function
       ElseIf UserIm& <> 0& Then
        CheckEmpowered = True
        Yield 0.6
        Call PostMessage(UserIm&, WM_CLOSE, 0&, 0&)
        Exit Function
    End If
End Function

Public Function FindDownLoadWindow() As Long
    Dim AolFrame As Long, AolMdi As Long, AolChild As Long
    Dim AolStatic As Long
    AolFrame& = FindWindow("AOL Frame25", vbNullString)
    AolMdi& = FindWindowEx(AolFrame&, 0&, "MDIClient", vbNullString)
    AolChild& = FindWindowEx(AolMdi&, 0&, "AOL Child", vbNullString)
    AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
    If InStr(GetCaption(AolChild&), "File Transfer") <> 0& And InStr(GetText(AolStatic&), "Downloading") <> 0& Then
        FindDownLoadWindow& = AolChild&
        Exit Function
       Else
        Do
            AolChild& = FindWindowEx(AolMdi&, AolChild&, "AOL Child", vbNullString)
            AolStatic& = FindWindowEx(AolChild&, 0&, "_AOL_Static", vbNullString)
            If InStr(GetCaption(AolChild&), "File Transfer") <> 0& And InStr(GetText(AolStatic&), "Downloading") <> 0& Then
                FindDownLoadWindow& = AolChild&
                Exit Function
            End If
        Loop Until AolChild& = 0&
    End If
    FindDownLoadWindow& = AolChild&
End Function

Public Function DownLoadStatus(Optional EmphasisOnStats As Boolean = True) As String
    If FindDownLoadWindow& = 0& Then
        DownLoadStatus$ = "Not currently downloading"
        Exit Function
    End If
    Dim AolStatic1 As Long, AolStatic2 As Long
    AolStatic1& = FindWindowEx(FindDownLoadWindow&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(FindDownLoadWindow&, AolStatic1&, "_AOL_Static", vbNullString)
    If EmphasisOnStats = True Then DownLoadStatus$ = "File transfer for: <b>" & GetInstance(GetText(AolStatic1&), " ", 3) & "</b>" & vbCrLf & "Percentage done: <b>" & ExtractNumeric(GetCaption(FindDownLoadWindow&)) & "%</b>" & vbCrLf & "Time remaining: <b>" & GetText(AolStatic2&)
    If EmphasisOnStats = False Then DownLoadStatus$ = "File transfer for: " & GetInstance(GetText(AolStatic1&), " ", 3) & vbCrLf & "Percentage done: " & ExtractNumeric(GetCaption(FindDownLoadWindow&)) & "%" & vbCrLf & "Time remaining: " & GetText(AolStatic2&)
End Function

Public Function FindUpLoadWindow() As Long
    Dim AolModal As Long, AolStatic As Long
    AolModal& = FindWindow("_AOL_Modal", vbNullString)
    AolStatic& = FindWindowEx(AolModal&, 0&, "_AOL_Static", vbNullString)
    If InStr(GetCaption(AolModal&), "File Transfer") <> 0& And InStr(GetText(AolStatic&), "Uploading") <> 0& Then
        FindUpLoadWindow& = AolModal&
        Exit Function
       Else
        Do
            AolModal& = FindWindow("_AOL_Modal", vbNullString)
            AolStatic& = FindWindowEx(AolModal&, 0&, "_AOL_Static", vbNullString)
            If InStr(GetCaption(AolModal&), "File Transfer") <> 0& And InStr(GetText(AolStatic&), "Uploading") <> 0& Then
                FindUpLoadWindow& = AolModal&
                Exit Function
            End If
        Loop Until AolModal& = 0&
    End If
    FindUpLoadWindow& = AolModal&
End Function

Public Function UpLoadStatus(Optional EmphasisOnStats As Boolean = True) As String
    If FindUpLoadWindow& = 0& Then
        UpLoadStatus$ = "Not currently uploading"
        Exit Function
    End If
    Dim AolStatic1 As Long, AolStatic2 As Long
    AolStatic1& = FindWindowEx(FindUpLoadWindow&, 0&, "_AOL_Static", vbNullString)
    AolStatic2& = FindWindowEx(FindUpLoadWindow&, AolStatic1&, "_AOL_Static", vbNullString)
    If EmphasisOnStats = True Then UpLoadStatus$ = "File transfer for: <b>" & GetInstance(GetText(AolStatic1&), " ", 3) & "</b>" & vbCrLf & "Percentage done: <b>" & ExtractNumeric(GetCaption(FindUpLoadWindow&)) & "%</b>" & vbCrLf & "Time remaining: <b>" & GetText(AolStatic2&)
    If EmphasisOnStats = False Then UpLoadStatus$ = "File transfer for: " & GetInstance(GetText(AolStatic1&), " ", 3) & vbCrLf & "Percentage done: " & ExtractNumeric(GetCaption(FindUpLoadWindow&)) & "%" & vbCrLf & "Time remaining: " & GetText(AolStatic2&)
End Function

Public Sub UpChat()
    Call WinEnable(AolFrame&)
    Call WinMinimize(FindUpLoadWindow&)
    Call WinDisable(FindUpLoadWindow&)
End Sub

Public Sub UnUpChat()
    Call WinDisable(AolFrame&)
    Call WinEnable(FindUpLoadWindow&)
    Call WinRestore(FindUpLoadWindow&)
End Sub

Public Function AimFindAimWindow() As Long
    AimFindAimWindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
End Function

Public Function AimUserSn() As String
    AimUserSn$ = Mid(GetCaption(AimFindAimWindow&), 1, InStr(GetCaption(AimFindAimWindow&), "'") - 1)
End Function

Public Function AimFindRoom() As Long
    AimFindRoom& = FindWindow("AIM_ChatWnd", vbNullString)
End Function

Public Sub AimRoomSend(SendString As String, Optional ClearBefore As Boolean = True)
    Dim WndAte1 As Long, WndAte2 As Long, OscarButton1 As Long
    Dim OscarButton2 As Long, OscarButton3 As Long, OscarButton4 As Long
    WndAte1& = FindWindowEx(AimFindRoom&, 0&, "WndAte32Class", vbNullString)
    WndAte2& = FindWindowEx(AimFindRoom&, WndAte1&, "WndAte32Class", vbNullString)
    OscarButton1& = FindWindowEx(AimFindRoom&, 0&, "_Oscar_IconBtn", vbNullString)
    OscarButton2& = FindWindowEx(AimFindRoom&, OscarButton1&, "_Oscar_IconBtn", vbNullString)
    OscarButton3& = FindWindowEx(AimFindRoom&, OscarButton2&, "_Oscar_IconBtn", vbNullString)
    OscarButton4& = FindWindowEx(AimFindRoom&, OscarButton3&, "_Oscar_IconBtn", vbNullString)
    If ClearBefore = True Then Call SendMessageByString(WndAte2&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(WndAte2&, WM_SETTEXT, 0&, SendString$)
    Call SendMessage(OscarButton4&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OscarButton4&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function AimRoomName() As String
    AimRoomName$ = ReplaceCharacters(GetCaption(AimFindRoom&), "Chat Room: ", "")
End Function

Public Function AimRoomCount() As Long
    Dim OscarTree As Long
    OscarTree& = FindWindowEx(AimFindRoom&, 0, "_Oscar_Tree", vbNullString)
    AimRoomCount& = ListCount(OscarTree&)
End Function

Public Function AimFindIm() As Long
    AimFindIm& = FindWindow("AIM_IMessage", vbNullString)
End Function

Public Function AimImScreenName() As String
    Dim ImCaption As String, PosOfDivider As Long, PrepScreenName As String
    ImCaption$ = GetCaption(AimFindOpenIm)
    If InStr(ImCaption$, " - Instant Message") <> 0& Then
       PosOfDivider& = InStr(ImCaption$, " -")
       PrepScreenName$ = Left(ImCaption$, PosOfDivider&)
       AimImScreenName$ = PrepScreenName$
      Else
       AimImScreenName$ = ""
    End If
End Function

Public Sub AimInstantMessage(Person As String, message As String)
    Dim OscarTab As Long, OscarIcon1 As Long, OscarPCombo As Long
    Dim Edit As Long, WndAte1 As Long, WndAte2 As Long, OscarIcon2 As Long
    Dim InfoMessage As Long, InfoButton As Long
    OscarTab& = FindWindowEx(AimFindAimWindow&, 0&, "_Oscar_TabGroup", vbNullString)
    OscarIcon1& = FindWindowEx(OscarTab&, 0&, "_Oscar_IconBtn", vbNullString)
    Call PostMessage(OscarIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OscarIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        OscarPCombo& = FindWindowEx(AimFindIm&, 0&, "_Oscar_PersistantCombo", vbNullString)
        Edit& = FindWindowEx(OscarPCombo&, 0&, "Edit", vbNullString)
        WndAte1& = FindWindowEx(AimFindIm&, 0&, "WndAte32Class", vbNullString)
        WndAte2& = FindWindowEx(AimFindIm&, WndAte1&, "WndAte32Class", vbNullString)
        OscarIcon2& = FindWindowEx(AimFindIm&, 0&, "_Oscar_IconBtn", vbNullString)
    Loop Until OscarPCombo& <> 0& And Edit& <> 0& And WndAte1& <> 0& And WndAte2& <> 0& And OscarIcon2& <> 0&
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(WndAte2&, WM_SETTEXT, 0&, message$)
    Call SendMessage(OscarIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OscarIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        InfoMessage& = FindWindow("#32770", "Information")
        InfoButton& = FindWindowEx(InfoMessage&, 0&, "Button", vbNullString)
    Loop Until InfoMessage& <> 0& And InfoButton& <> 0&
    If InfoMessage& <> 0& And InfoButton& <> 0& Then
        Call PostMessage(InfoButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(InfoButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(AimFindIm&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
End Sub

Public Function AimFindBuddyInvite() As Long
    AimFindBuddyInvite& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
End Function

Public Sub AimBuddyInvitation(Person As String, message As String, RoomName As String)
    Dim OscarTab As Long, OscarIcon1 As Long, OscarIcon2 As Long
    Dim OscarIcon3 As Long, OscarIcon4 As Long, OscarIcon5 As Long
    Dim Edit1 As Long, Edit2 As Long, Edit3 As Long, InfoMessage As Long
    Dim InfoButton As Long
    OscarTab& = FindWindowEx(AimFindAimWindow&, 0&, "_Oscar_TabGroup", vbNullString)
    OscarIcon1& = FindWindowEx(OscarTab&, 0&, "_Oscar_IconBtn", vbNullString)
    OscarIcon2& = FindWindowEx(OscarTab&, OscarIcon1&, "_Oscar_IconBtn", vbNullString)
    Call PostMessage(OscarIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OscarIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Edit1& = FindWindowEx(AimFindBuddyInvite&, 0&, "Edit", vbNullString)
        Edit2& = FindWindowEx(AimFindBuddyInvite&, Edit1&, "Edit", vbNullString)
        Edit3& = FindWindowEx(AimFindBuddyInvite&, Edit2&, "Edit", vbNullString)
        OscarIcon3& = FindWindowEx(AimFindBuddyInvite&, 0&, "_Oscar_IconBtn", vbNullString)
        OscarIcon4& = FindWindowEx(AimFindBuddyInvite&, OscarIcon3&, "_Oscar_IconBtn", vbNullString)
        OscarIcon5& = FindWindowEx(AimFindBuddyInvite&, OscarIcon4&, "_Oscar_IconBtn", vbNullString)
    Loop Until AimFindBuddyInvite& <> 0& And Edit1& <> 0& And Edit2& <> 0& And Edit3& <> 0& And OscarIcon3& <> 0& And OscarIcon4& <> 0& And OscarIcon5& <> 0&
    Call SendMessageByString(Edit1&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(Edit2&, WM_SETTEXT, 0&, message$)
    Call SendMessageByString(Edit3&, WM_SETTEXT, 0&, RoomName$)
    Call PostMessage(OscarIcon5&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OscarIcon5&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        InfoMessage& = FindWindow("#32770", "Information")
        InfoButton& = FindWindowEx(InfoMessage&, 0&, "Button", vbNullString)
    Loop Until InfoMessage& <> 0& And InfoButton& <> 0&
    If InfoMessage& <> 0& And InfoButton& <> 0& Then
        Call PostMessage(InfoButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(InfoButton&, WM_KEYUP, VK_SPACE, 0&)
    End If
End Sub

Public Sub AimRoomEnter(RoomName As String)
    Dim OscarTab As Long, OscarIcon1 As Long, OscarIcon2 As Long
    Dim OscarIcon3 As Long, OscarIcon4 As Long, OscarIcon5 As Long
    Dim Edit1 As Long, Edit2 As Long, Edit3 As Long
    OscarTab& = FindWindowEx(AimFindAimWindow&, 0&, "_Oscar_TabGroup", vbNullString)
    OscarIcon1& = FindWindowEx(OscarTab&, 0&, "_Oscar_IconBtn", vbNullString)
    OscarIcon2& = FindWindowEx(OscarTab&, OscarIcon1&, "_Oscar_IconBtn", vbNullString)
    Call PostMessage(OscarIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OscarIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Edit1& = FindWindowEx(AimFindBuddyInvite&, 0&, "Edit", vbNullString)
        Edit2& = FindWindowEx(AimFindBuddyInvite&, Edit1&, "Edit", vbNullString)
        Edit3& = FindWindowEx(AimFindBuddyInvite&, Edit2&, "Edit", vbNullString)
        OscarIcon3& = FindWindowEx(AimFindBuddyInvite&, 0&, "_Oscar_IconBtn", vbNullString)
        OscarIcon4& = FindWindowEx(AimFindBuddyInvite&, OscarIcon3&, "_Oscar_IconBtn", vbNullString)
        OscarIcon5& = FindWindowEx(AimFindBuddyInvite&, OscarIcon4&, "_Oscar_IconBtn", vbNullString)
    Loop Until AimFindBuddyInvite& <> 0& And Edit1& <> 0& And Edit2& <> 0& And Edit3& <> 0& And OscarIcon3& <> 0& And OscarIcon4& <> 0& And OscarIcon5& <> 0&
    Call SendMessageByString(Edit1&, WM_SETTEXT, 0&, AimUserSn$)
    Call SendMessageByString(Edit2&, WM_SETTEXT, 0&, "")
    Call SendMessageByString(Edit3&, WM_SETTEXT, 0&, RoomName$)
    Call PostMessage(OscarIcon5&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OscarIcon5&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Function AimFindOpenIm() As Long
   If InStr(GetCaption(AimFindIm&), " - Instant Message") <> 0& Then
        AimFindOpenIm = AimFindIm&
        Exit Function
       Else
        Do
            If InStr(GetCaption(AimFindIm&), " - Instant Message") <> 0& Then
                AimFindOpenIm = AimFindIm&
                Exit Function
            End If
       Loop Until AimFindIm& = 0&
   End If
   AimFindOpenIm = AimFindIm&
End Function

Public Function AimFindAcceptIm() As Long
    AimFindAcceptIm& = FindWindow("#32770", "Accept Message")
End Function

Public Function AimOnline() As Boolean
    If AimFindAimWindow& <> 0& Then
        AimOnline = True
       ElseIf AimFindAimWindow& = 0& Then
        AimOnline = False
    End If
End Function

Public Sub AimRoomClear()
    Dim WndAte As Long
    WndAte& = FindWindowEx(AimFindRoom&, 0&, "WndAte32Class", vbNullString)
    Call SendMessageByString(WndAte&, WM_SETTEXT, 0&, "")
End Sub

Public Function AimTrimTimeStamp(TrimThisString As String) As String
    Dim NewMain As String
    NewMain$ = TrimThisString$
    If InStr(NewMain$, "<!--(") <> 0& Then
        Do: DoEvents
            If Right("<!--(", Len("<!--(") + 18) = ">" Then
                If Right("<!--(", Len("<!--(") + 18) = "" Then
                    NewMain$ = Left(NewMain$, InStr(NewMain$, "<!--(") - 1)
                   ElseIf Right("<!--(", Len("<!--(") + 18) <> "" Then
                    NewMain$ = ReplaceCharacters(Left(NewMain$, InStr(NewMain$, "<!--(") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<!--(") - 19), ">", "")
                End If
               ElseIf Right("<!--(", Len("<!--(") + 18) <> ">" Then
                If Right("<!--(", Len("<!--(") + 19) = "" Then
                    NewMain$ = Left(NewMain$, InStr(NewMain$, "<!--(") - 1)
                   ElseIf Right("<!--(", Len("<!--(") + 19) <> "" Then
                    NewMain$ = ReplaceCharacters(Left(NewMain$, InStr(NewMain$, "<!--(") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<!--(") - 18), "->", "")
                End If
            End If
        Loop Until InStr(NewMain$, "<!--(") = 0&
    End If
    AimTrimTimeStamp$ = ReplaceCharacters(NewMain$, ">", "")
End Function

Public Function AimRoomGetText() As String
    Dim WndAte As Long, prepstring As String
    WndAte& = FindWindowEx(AimFindRoom&, 0&, "WndAte32Class", vbNullString)
    prepstring$ = TrimHtml(GetText(WndAte&))
    AimRoomGetText$ = prepstring$
End Function

Public Function AimFindSignOn() As Long
    AimFindSignOn& = FindWindow("#32770", "Sign On")
End Function

Public Function AimFindLocate() As Long
    AimFindLocate& = FindWindow("_Oscar_Locate", vbNullString)
End Function

Public Function AimProfileGet(Person As String) As String
    Dim OscarPCombo As Long, Edit As Long, Button As Long
    Dim WndAteClass As Long, AteClass As Long, prepstring As String
    Call RunMenuByString(AimFindAimWindow&, "Get Member Inf&o")
    OscarPCombo& = FindWindowEx(AimFindLocate&, 0&, "_Oscar_PersistantCombo", vbNullString)
    Edit& = FindWindowEx(OscarPCombo&, 0&, "Edit", vbNullString)
    Button& = FindWindowEx(AimFindLocate&, 0&, "Button", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, Person$)
    Call SendMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
    WndAteClass& = FindWindowEx(AimFindLocate&, 0&, "WndAte32Class", vbNullString)
    AteClass& = FindWindowEx(WndAteClass&, 0&, "Ate32Class", vbNullString)
    Yield 2
    prepstring$ = TrimHtml(GetText(AteClass&))
    AimProfileGet$ = prepstring$
    Call SendMessage(AimFindLocate&, WM_CLOSE, 0&, 0&)
End Function

Public Sub AimWebSearch(WebUrl As String)
    Dim Edit As Long, OscarIcon As Long
    Edit& = FindWindowEx(AimFindAimWindow&, 0&, "Edit", vbNullString)
    OscarIcon& = FindWindowEx(AimFindAimWindow&, 0&, "_Oscar_IconBtn", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, WebUrl$)
    Call SendMessage(OscarIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OscarIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub AimAddRoomToList(ListBox As ListBox)
    Dim OscarTree As Long, FixedString As String, Count As Long
    OscarTree& = FindWindowEx(AimFindRoom&, 0&, "_Oscar_Tree", vbNullString)
    For Count& = 0 To ListCount(OscarTree&) - 1
        FixedString$ = String(200, vbNullChar)
        Call SendMessageByString(OscarTree&, LB_GETTEXT, Count&, FixedString$)
        ListBox.AddItem FixedString$
    Next Count&
End Sub

Public Sub AimSignOn(ScreenName As String, PassWord As String)
    If AimFindSignOn& = 0& Then Exit Sub
    Dim Combo As Long, Edit1 As Long, Edit2 As Long
    Dim OscarIcon1 As Long, OscarIcon2 As Long, OscarIcon3 As Long
    Dim AimError As Long, Button1 As Long
    Combo& = FindWindowEx(AimFindSignOn&, 0&, "ComboBox", vbNullString)
    Edit1& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Edit2& = FindWindowEx(AimFindSignOn&, 0&, "Edit", vbNullString)
    Call SendMessageByString(Edit1&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessageByString(Edit2&, WM_SETTEXT, 0&, PassWord$)
    OscarIcon1& = FindWindowEx(AimFindSignOn&, 0&, "_Oscar_IconBtn", vbNullString)
    OscarIcon2& = FindWindowEx(AimFindSignOn&, OscarIcon1&, "_Oscar_IconBtn", vbNullString)
    OscarIcon3& = FindWindowEx(AimFindSignOn&, OscarIcon2&, "_Oscar_IconBtn", vbNullString)
    Call SendMessage(OscarIcon3&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OscarIcon3&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        AimError& = FindWindow("#32770", "AOL Instant Messenger (TM) Error")
        Button1& = FindWindowEx(AimError&, 0&, "Button", vbNullString)
    Loop Until AimError& <> 0& And Button1& <> 0& Or AimFindAimWindow& <> 0&
    If AimError& <> 0& And Button1& <> 0& Then
        Call SendMessage(Button1&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(Button1&, WM_KEYUP, VK_SPACE, 0&)
        Exit Sub
       ElseIf AimFindAimWindow& <> 0& Then
        Exit Sub
    End If
End Sub

Public Sub AimAddBuddiesToList(ListBox As ListBox)
    Dim OscarTree As Long, FixedString As String, Count As Long
    Dim OscarTab As Long
    OscarTab& = FindWindowEx(AimFindAimWindow&, 0, "_Oscar_TabGroup", vbNullString)
    OscarTree& = FindWindowEx(OscarTab&, 0&, "_Oscar_Tree", vbNullString)
    For Count& = 0 To ListCount(OscarTree&) - 1
        FixedString$ = String(200, vbNullChar)
        Call SendMessageByString(OscarTree&, LB_GETTEXT, Count&, FixedString$)
        If InStr(FixedString$, "(") = 0& Then ListBox.AddItem FixedString$
    Next Count&
End Sub

Public Sub AimMassInstantMessage(ScreenNameList As Control, message As String, Optional Delay As Single = "0.6")
    Dim ListIndex As Long
    On Error Resume Next
    For ListIndex& = 0 To ScreenNameList.ListCount - 1
        Call AimInstantMessage(ScreenNameList.list(ListIndex&), message$)
        Call Yield(Val(Delay))
    Next ListIndex&
End Sub


Public Sub MailUnreadToFlash()
    Dim SessionWindow As Long, AolIcon1 As Long, AolIcon2 As Long
    Dim Check1 As Long, Check2 As Long, Check3 As Long
    Dim Check4 As Long, Check5 As Long, Check6 As Long
    Dim AolModal As Long, ModalIcon As Long
    Call PopUpIcon(2, "t")
   Do: DoEvents
      SessionWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Automatic AOL")
      AolIcon1& = FindWindowEx(SessionWindow&, 0&, "_AOL_Icon", vbNullString)
      AolIcon2& = FindWindowEx(SessionWindow&, AolIcon1&, "_AOL_Icon", vbNullString)
      Check1& = FindWindowEx(SessionWindow&, 0&, "_AOL_Checkbox", vbNullString)
      Check2& = FindWindowEx(SessionWindow&, Check1&, "_AOL_Checkbox", vbNullString)
      Check3& = FindWindowEx(SessionWindow&, Check2&, "_AOL_Checkbox", vbNullString)
      Check4& = FindWindowEx(SessionWindow&, Check3&, "_AOL_Checkbox", vbNullString)
      Check5& = FindWindowEx(SessionWindow&, Check4&, "_AOL_Checkbox", vbNullString)
      Check6& = FindWindowEx(SessionWindow&, Check5&, "_AOL_Checkbox", vbNullString)
   Loop Until SessionWindow& <> 0& And AolIcon1& <> 0& And AolIcon2& <> 0& And Check1& <> 0& And Check2& <> 0& And Check4& <> 0& And Check5& <> 0& And Check6& <> 0&
    Call Yield(0.6)
    Call ClickIcon(AolIcon2&)
       Do: DoEvents
          AolModal& = FindWindow("_AOL_Modal", "Run Automatic AOL Now")
          ModalIcon& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
       Loop Until AolModal& <> 0& And ModalIcon& <> 0&
End Sub


Public Function SpyCodeGenerator(WinHandle As Long, Optional ParentName As String = "ParentWin", Optional ChildName As String = "ChildWin", Optional Emphasis As String = "OurWin") As String
    If IsWindow(WinHandle&) = 0& Then
        SpyCodeGenerator$ = "Handle does not exist."
        Exit Function
    End If
End Function


Public Sub ProfileEdit(Optional MemberName As String, Optional Location As String, Optional Birthday As String, Optional Gender As Integer, Optional MaritalStatus As String, Optional Hobbies As String, Optional ComputersUsed As String, Optional Occupation As String, Optional PersonalQuote As String)
    Dim MNEdit As Long, LEdit As Long, BEdit As Long
    Dim GEdit As Long, MSEdit As Long, HEdit As Long
    Dim CUEdit As Long, OEdit As Long, PQEdit As Long
    Dim NWatchModal As Long, ModalIcon As Long, ModalCheck As Long
    Dim ProfileWindow As Long, UpdateIcon As Long
    Call PopUpIcon(5, "y")
        ProfileWindow& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Edit Your Online Profile")
        NWatchModal& = FindWindow("_AOL_Modal", vbNullString)
        ModalIcon& = FindWindowEx(NWatchModal&, 0&, "_AOL_Icon", vbNullString)
        ModalCheck& = FindWindowEx(NWatchModal&, 0&, "_AOL_Checkbox", vbNullString)
End Sub

Public Sub ScanChat(ScanFor As String, ReplyWith As String, Optional GetFromUser As Boolean = False)
    If InStr(RoomLastLineMessage$, ScanFor$) <> 0& Then
        RoomSend ReplyWith$
        Yield 2
        Exit Sub
    End If
End Sub
Public Sub SpiralScroll(TextToReverse As String, Optional Delay As Single = "0.6")
    Dim Step As Long, NewString As String
    For Step& = 1 To Len(TextToReverse$)
        NewString$ = Mid(TextToReverse$, 1, Step&)
        IpartySend NewString$
        Call Yield(Val(Delay))
    Next Step&
    For Step& = 1 To Len(TextToReverse$)
        NewString$ = Mid(TextToReverse$, Step& + 1, Len(TextToReverse$) + 1)
        IpartySend NewString$
        Call Yield(Val(Delay))
    Next Step&
End Sub







Public Function AbleSnMaker() As Boolean
    Dim CreateOrDelete As Long, Create As Long, NoMasterModal As Long
    Dim NamesMsg5 As Long, OkButton5 As Long, CreateModal As Long
    Dim CreateList As Long, CreateIcon As Long, ModalCIcon As Long
    Dim OtherModal As Long, OtherIcon As Long
    CreateOrDelete& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Create or Delete Screen Names")
    If CreateOrDelete& = 0& Then Call PopUpIcon(5, "n")
    Do: DoEvents
        CreateOrDelete& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Create or Delete Screen Names")
        CreateList& = FindChildByClass(CreateOrDelete&, "_AOL_Listbox")
    Loop Until CreateOrDelete& <> 0& And CreateList& <> 0&
    DoEvents
    Call SendMessageLong(CreateList&, LB_SETCURSEL, 3, 0&)
    Call PostMessage(CreateList&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do: DoEvents
        Create& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Create a Screen Name")
        CreateIcon& = FindChildByClass(Create&, "_AOL_Icon")
    Loop Until Create& <> 0& And CreateIcon& <> 0&
    DoEvents
    Call ClickIcon(CreateIcon&)
    Do: DoEvents
        NamesMsg5& = FindWindow("#32770", "America Online")
        OkButton5& = FindWindowEx(NamesMsg5&, 0&, "Button", vbNullString)
        CreateModal& = FindWindow("_AOL_Modal", "Create a Screen Name")
        ModalCIcon& = NextOfClassByCount(CreateModal&, "_AOL_Icon", 2)
        OtherModal& = FindWindow("_AOL_Modal", vbNullString)
        OtherIcon& = FindChildByClass(OtherModal&, "_AOL_Icon")
    Loop Until (NamesMsg5& <> 0& And OkButton5& <> 0&) Or (CreateModal& <> 0& And ModalCIcon& <> 0&) Or (OtherModal& <> 0& And OtherIcon& <> 0&)
    If NamesMsg5& <> 0& Then
        Call PostMessage(OkButton5&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OkButton5&, WM_KEYUP, VK_SPACE, 0&)
        AbleSnMaker = False
        Exit Function
       ElseIf CreateModal& <> 0& Then
        Call ClickIcon(ModalCIcon&)
        Yield 0.8
        AbleSnMaker = True
        Exit Function
       ElseIf OtherModal& <> 0& Then
        Call ClickIcon(OtherIcon&)
        AbleSnMaker = False
        Exit Function
    End If
End Function

Public Sub CheckScreenName(ScreenNames As Control, Available As Control)
    Dim ListIndex As Long, CreateEdit As Long, CreateOrDelete As Long
    Dim CreateList As Long, Create As Long, CreateIcon As Long
    Dim CreateModal As Long, ModalCIcon As Long, SetPw As Long
    Dim SetPwIcon As Long, NamesMsg As Long, OKButton As Long
    On Error Resume Next
    If AbleSnMaker = False Then Exit Sub
    For ListIndex& = 0& To ScreenNames.ListCount - 1
        CreateOrDelete& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Create or Delete Screen Names")
        If CreateOrDelete& = 0& Then Call PopUpIcon(5, "n")
        Do: DoEvents
            CreateOrDelete& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Create or Delete Screen Names")
            CreateList& = FindChildByClass(CreateOrDelete&, "_AOL_Listbox")
        Loop Until CreateOrDelete& <> 0& And CreateList& <> 0&
        DoEvents
        Call SendMessageLong(CreateList&, LB_SETCURSEL, 3, 0&)
        Call PostMessage(CreateList&, WM_LBUTTONDBLCLK, 0&, 0&)
        Do: DoEvents
            Create& = FindWindowEx(AolMdi&, 0&, "AOL Child", "Create a Screen Name")
            CreateIcon& = FindChildByClass(Create&, "_AOL_Icon")
        Loop Until Create& <> 0& And CreateIcon& <> 0&
        DoEvents
        Call ClickIcon(CreateIcon&)
        Do: DoEvents
            CreateModal& = FindWindow("_AOL_Modal", "Create a Screen Name")
            ModalCIcon& = FindChildByClass(CreateModal&, "_AOL_Icon")
            CreateEdit& = FindChildByClass(CreateModal&, "_AOL_Edit")
        Loop Until (CreateModal& <> 0& And ModalCIcon& <> 0& And CreateEdit& <> 0&)
        Call SetText(CreateEdit&, ScreenNames.list(ListIndex&))
        Call ClickIcon(ModalCIcon&)
        Do: DoEvents
            SetPw& = FindWindow("_AOL_Modal", "Set Password")
            SetPwIcon& = NextOfClassByCount(SetPw&, "_AOL_Icon", 2)
            NamesMsg& = FindWindow("#32770", "America Online")
            OKButton& = FindWindowEx(NamesMsg&, 0&, "Button", vbNullString)
        Loop Until (SetPw& <> 0& And SetPwIcon& <> 0&) Or (NamesMsg& <> 0& And OKButton& <> 0&)
        If (SetPw& <> 0& And SetPwIcon& <> 0&) Then
            Call ClickIcon(SetPwIcon&)
            Yield 0.7
            Available.AddItem ScreenNames.list(ListIndex&)
            Resume Next
           ElseIf (NamesMsg& <> 0& And OKButton& <> 0&) Then
            Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
            Call ClickIcon(NextOfClassByCount(CreateModal&, "_AOL_Icon", 2))
            Yield 0.7
            Resume Next
        End If
    Next ListIndex&
End Sub
Public Sub StaffAimInstantMessage(Person As String, message As String)
    Dim OscarTab As Long, OscarIcon1 As Long, OscarPCombo As Long
    Dim Edit As Long, WndAte1 As Long, WndAte2 As Long, OscarIcon2 As Long
    Dim InfoMessage As Long, InfoButton As Long
    OscarTab& = FindWindowEx(AimFindAimWindow&, 0&, "_Oscar_TabGroup", vbNullString)
    OscarIcon1& = FindWindowEx(OscarTab&, 0&, "_Oscar_IconBtn2", vbNullString)
    Call PostMessage(OscarIcon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(OscarIcon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        OscarPCombo& = FindWindowEx(AimFindIm&, 0&, "_Oscar_PersistantCombo", vbNullString)
        Edit& = FindWindowEx(OscarPCombo&, 0&, "Edit", vbNullString)
        WndAte1& = FindWindowEx(AimFindIm&, 0&, "WndAte32Class", vbNullString)
        WndAte2& = FindWindowEx(AimFindIm&, WndAte1&, "WndAte32Class", vbNullString)
        OscarIcon2& = FindWindowEx(AimFindIm&, 0&, "_Oscar_IconBtn", vbNullString)
    Loop Until OscarPCombo& <> 0& And Edit& <> 0& And WndAte1& <> 0& And WndAte2& <> 0& And OscarIcon2& <> 0&
    Call SendMessageByString(Edit&, WM_SETTEXT, 0&, Person$)
    Call SendMessageByString(WndAte2&, WM_SETTEXT, 0&, message$)
    Call SendMessage(OscarIcon2&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(OscarIcon2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        InfoMessage& = FindWindow("#32770", "Information")
        InfoButton& = FindWindowEx(InfoMessage&, 0&, "Button", vbNullString)
    Loop Until InfoMessage& <> 0& And InfoButton& <> 0&
    If InfoMessage& <> 0& And InfoButton& <> 0& Then
        Call PostMessage(InfoButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(InfoButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(AimFindIm&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
End Sub


Public Sub Closebox()
ParHand1& = FindWindow("AOL Frame25", "America  Online")
OurParent& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "AOL Child", UserSN + "'s Online Mailbox")
WinClose (ourhandle)
End Sub


Sub timeout(Duration)
'Strictly old school code, but the older style is better for certain subs
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop

End Sub
