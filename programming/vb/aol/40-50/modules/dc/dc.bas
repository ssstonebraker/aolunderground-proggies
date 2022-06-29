Attribute VB_Name = "dc"
' DC Bas 2001 by: DC (fissuredc@aol.com)
' Made for AOL version 4 & 5
' Group : hTech
' E-mail me to join

Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLICK = &H203
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205

Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateEllipticRgn& Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePolyPolygonRgn& Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long)
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Global Const WINDING = 2
Global Const ALTERNATE = 1
Global Const RGN_OR = 2
Global Const RGN_DIFF = 4
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private vbTray As NOTIFYICONDATA
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long

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

Public Const SW_Hide = 0
Public Const SW_SHOW = 5
Public Const SW_RESTORE& = 9

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


Public Function FindForwardWindow() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim Rich1 As Long, Rich2 As Long, Combo As Long
    Dim FontCombo As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
    FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
    If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
        FindForwardWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Rich1& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            Rich2& = FindWindowEx(Child&, Rich1&, "RICHCNTL", vbNullString)
            Combo& = FindWindowEx(Child&, 0&, "_AOL_Combobox", vbNullString)
            FontCombo& = FindWindowEx(Child&, 0&, "_AOL_FontCombo", vbNullString)
            If Rich1& <> 0& And Rich2& = 0& And Combo& = 0& And FontCombo& = 0& Then
                FindForwardWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindForwardWindow& = 0&
End Function

Public Function FindSendWindow() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim SendStatic As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
    If SendStatic& <> 0& Then
        FindSendWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            SendStatic& = FindWindowEx(Child&, 0&, "_AOL_Static", "Send Now")
            If SendStatic& <> 0& Then
                FindSendWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindSendWindow& = 0&
End Function

Public Sub button(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub Sendtext(Chat As String)
    Dim room As Long, AORich As Long, AORich2 As Long
    room& = FindRoom&
    AORich& = FindWindowEx(room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Function FindIM() As Long
    Dim aol As Long, MDI As Long, Child As Long, caption As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    caption$ = GetCaption(Child&)
    If InStr(caption$, "Instant Message") = 1 Or InStr(caption$, "Instant Message") = 2 Or InStr(caption$, "Instant Message") = 3 Then
        FindIM& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            caption$ = GetCaption(Child&)
            If InStr(caption$, "Instant Message") = 1 Or InStr(caption$, "Instant Message") = 2 Or InStr(caption$, "Instant Message") = 3 Then
                FindIM& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindIM& = Child&
End Function
Public Function FindInfoWindow() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim AOLCheck As Long, AOLIcon As Long, aolstatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
    aolstatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And aolstatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(Child&, 0&, "_AOL_Checkbox", vbNullString)
            aolstatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(Child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(Child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And aolstatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindInfoWindow& = Child&
End Function
Public Function FindMailBox() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim TabControl As Long, TabPage As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
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

Public Function FindRoom() As Long
    Dim aol As Long, MDI As Long, Child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, aolstatic As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
    aolstatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And aolstatic& <> 0& Then
        FindRoom& = Child&
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(Child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(Child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(Child&, 0&, "_AOL_Icon", vbNullString)
            aolstatic& = FindWindowEx(Child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And aolstatic& <> 0& Then
                FindRoom& = Child&
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    FindRoom& = Child&
End Function

Public Sub Moveform(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub OnTop(FormName As Form)
    Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Sub SendMail(Person As String, Subject As String, message As String)
    Dim aol As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
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
    Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    DoEvents
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub SendMail5(Person As String, Subject As String, message As String)
    Dim aol As Long, MDI As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, OpenSend As Long, DoIt As Long
    Dim Rich As Long, EditTo As Long, EditCC As Long
    Dim EditSubject As Long, SendButton As Long
    Dim Combo As Long, fCombo As Long, ErrorWindow As Long
    Dim Button1 As Long, Button2 As Long
    Dim mine As Long, newbut As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
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
               For DoIt& = 1 To 15
            SendButton& = FindWindowEx(OpenSend&, SendButton&, "_AOL_Icon", vbNullString)
        Next DoIt&
    Loop Until OpenSend& <> 0& And EditTo& <> 0& And EditCC& <> 0& And EditSubject& <> 0& And Rich& <> 0& And SendButton& <> 0& And Combo& <> 0& And fCombo& <> 0& & SendButton& <> Button1& And SendButton& <> Button2&
        Call SendMessageByString(EditTo&, WM_SETTEXT, 0, Person$)
    DoEvents
        Call SendMessageByString(EditSubject&, WM_SETTEXT, 0, Subject$)
    DoEvents
        Call SendMessageByString(Rich&, WM_SETTEXT, 0, message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    End Sub
    Public Sub Pause(duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= duration
        DoEvents
    Loop
End Sub
Public Sub PrivateRoom5(room As String)
    Call Keyword("aol://2719:2-2-" & room$)
End Sub
Public Sub Keyword(KW As String)
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub InstantMessage(Person As String, message As String)
    Dim aol As Long, MDI As Long, IM As Long, Rich As Long
    Dim SendButton As Long, OK As Long, button As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
        SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
        SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
    Loop Until IM& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or IM& = 0&
    If OK& <> 0& Then
        button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Sub LoadText(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub
Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
Public Sub WindowHide(hWnd As Long)
    Call ShowWindow(hWnd&, SW_Hide)
End Sub
Public Sub WindowShow(hWnd As Long)
    Call ShowWindow(hWnd&, SW_SHOW)
End Sub
Public Sub RunMenu(TopMenu As Long, SubMenu As Long)
    Dim aol As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(aol&, WM_COMMAND, mnID&, 0&)
End Sub
Public Sub RunMenuByString(SearchString As String)
    Dim aol As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
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
Public Function IMSender() As String
    Dim IM As Long, caption As String
    caption$ = GetCaption(FindIM&)
    If InStr(caption$, ":") = 0& Then
        IMSender$ = ""
        Exit Function
    Else
        IMSender$ = Right(caption$, Len(caption$) - InStr(caption$, ":") - 1)
    End If
End Function
Public Sub imsoff()
    Call InstantMessage("$IM_OFF", "DC Bas")
End Sub
Public Sub imson()
    Call InstantMessage("$IM_ON", "DC Bas")
End Sub

Public Function GetListText(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, Buffer$)
    GetListText$ = Buffer$
End Function
Public Function UserSN() As String
    Dim aol As Long, MDI As Long, welcome As Long
    Dim Child As Long, UserString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(Child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        UserSN$ = UserString$
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(Child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                UserSN$ = UserString$
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    UserSN$ = ""
End Function
Public Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function




Public Function aol()
aol = FindWindow("AOL Frame25", vbNullString)
End Function


Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Public Sub CloseWindow(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub
