Attribute VB_Name = "I_Own_AIM"
' I Own AIM v2 *UPDATE* by seb
' Last updated 1/31/99 8:43:23

Option Explicit

' WinMM
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' Kernel
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' User
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

' Global & Public Const
Const EM_UNDO = &HC7
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
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

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
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE

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

Public Const ENTA = 13
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const EM_LINESCROLL = &HB6

Type RECT
   Left As Long
   Top As Long
   right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type
Sub Chat_AddRoom_Combo(cmb As ComboBox)
    Dim ChatRoom As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, Buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer, BuddyTree As Long

    ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatRoom& <> 0 Then
        Do
            BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            Buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, Buffer$)
            TabPos = InStr(Buffer$, Chr$(9))
            NameText$ = right$(Buffer$, (Len(Buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            For mooz = 0 To cmb.ListCount - 1
                If name$ = cmb.List(mooz) Then
                    Well% = 123
                    GoTo endz
                End If
            Next mooz
            If Well% <> 123 Then
                cmb.AddItem name$
            Else
            End If
endz:
        Next MooLoo
    End If
End Sub

Sub AIM_SignOn2(ScrnName As String, PASS As String)
    Dim SignOnS2 As Long, Comb As Long
    Dim CombEdit As Long, SetScreenN As Long
    Dim password As Long, SetPassW As Long, Help2 As Long
    Dim SetUp2 As Long, SignOn2 As Long
    Dim BuddyList As Long, Klick As Long
    
    SignOnS2& = FindWindow("#32770", "Sign On")
    Comb& = FindWindowEx(SignOnS2&, 0, "ComboBox", vbNullString)
    CombEdit& = FindWindowEx(Comb&, 0, "Edit", vbNullString)
    SetScreenN& = SendMessageByString(CombEdit&, WM_SETTEXT, 0, ScrnName$)
    password& = FindWindowEx(SignOnS2&, 0, "Edit", vbNullString)
    Pause 0.3
    SetPassW& = SendMessageByString(password&, WM_SETTEXT, 0, PASS$)
    Call SendMessageLong(SetPassW&, WM_CHAR, VK_RETURN, 0&)
    Call SendMessageLong(SetPassW&, WM_CHAR, VK_RETURN, 0&)
    Help2& = FindWindowEx(SignOnS2&, 0, "_Oscar_IconBtn", vbNullString)
    SetUp2& = FindWindowEx(SignOnS2&, Help2&, "_Oscar_IconBtn", vbNullString)
    SignOn2& = FindWindowEx(SignOnS2&, SetUp2&, "_Oscar_IconBtn", vbNullString)
    Klick& = SendMessage(SignOn2&, WM_LBUTTONDOWN, 0, 0&)
    Klick& = SendMessage(SignOn2&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub AIM_AddBuddyTo_Combo(cmb As ComboBox)
    Dim BuddyList As Long, TabGroup As Long
    Dim BuddyTree As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, Buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0 Then
        Do
            TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
            BuddyTree& = FindWindowEx(TabGroup&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            Buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, Buffer$)
            TabPos = InStr(Buffer$, Chr$(9))
            NameText$ = right$(Buffer$, (Len(Buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            If InStr(name$, "(") <> 0 And InStr(name$, ")") <> 0 Then
                GoTo HellNo
            End If
            For mooz = 0 To cmb.ListCount - 1
                If name$ = cmb.List(mooz) Then
                    Well% = 123
                    GoTo HellNo
                End If
            Next mooz
            If Well% <> 123 Then
                cmb.AddItem name$
            Else
            End If
HellNo:
        Next MooLoo
    End If
End Sub
Sub AIM_AddBuddyTo_List(lis As ListBox)
    Dim BuddyList As Long, TabGroup As Long
    Dim BuddyTree As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, Buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0 Then
        Do
            TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
            BuddyTree& = FindWindowEx(TabGroup&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            Buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, Buffer$)
            TabPos = InStr(Buffer$, Chr$(9))
            NameText$ = right$(Buffer$, (Len(Buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            If InStr(name$, "(") <> 0 And InStr(name$, ")") <> 0 Then
                GoTo HellNo
            End If
            For mooz = 0 To lis.ListCount - 1
                If name$ = lis.List(mooz) Then
                    Well% = 123
                    GoTo HellNo
                End If
            Next mooz
            If Well% <> 123 Then
                lis.AddItem name$
            Else
            End If
HellNo:
        Next MooLoo
    End If
End Sub
Sub Change_SignOnCaption(NewCap As String)
    Dim SignOnS As Long, SetCap As String

    SignOnS& = FindWindow("#32770", "Sign On")
    SetCap$ = SendMessageByString(SignOnS&, WM_SETTEXT, 0, NewCap$)
End Sub

Sub Chat_AddRoom_List(lis As ListBox)
    Dim ChatRoom As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, Buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer, BuddyTree As Long

    ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatRoom& <> 0 Then
        Do
            BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            Buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, Buffer$)
            TabPos = InStr(Buffer$, Chr$(9))
            NameText$ = right$(Buffer$, (Len(Buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            For mooz = 0 To lis.ListCount - 1
                If name$ = lis.List(mooz) Then
                    Well% = 123
                    GoTo endz
                End If
            Next mooz
            If Well% <> 123 Then
                lis.AddItem name$
            Else
            End If
endz:
        Next MooLoo
    End If
End Sub

Sub Chat_Flash_ON_OFF()
    Dim PrefWin As Long, ZeeWin As Long, Flash As Long
    Dim OKbuttin As Long, ChatWindow As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Flash& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Call SendMessage(Flash&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Flash&, WM_KEYUP, VK_SPACE, 0&)
    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub


Sub Chat_IgnoreInvites()
    Dim ChatWindow As Long, PrefWin As Long, ZeeWin As Long
    Dim Buttin1 As Long, Buttin2 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, Buttin6 As Long
    Dim Buttin7 As Long, Buttin8 As Long, IIButtin As Long
    Dim OKbuttin As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    Buttin6& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Buttin7& = FindWindowEx(ZeeWin&, Buttin6&, "Button", vbNullString)
    Buttin8& = FindWindowEx(ZeeWin&, Buttin7&, "Button", vbNullString)
    IIButtin& = FindWindowEx(ZeeWin&, Buttin8&, "Button", vbNullString)
    Call SendMessage(IIButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(IIButtin&, WM_KEYUP, VK_SPACE, 0&)
    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub


Sub Chat_SoundsON_OFF()
    Dim ChatWindow As Long, ZeeWin As Long, PrefWin As Long
    Dim Buttin2 As Long, Buttin As Long, PlayMess As Long
    Dim Buttin1 As Long, Buttin22 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, PlaySend As Long
    Dim OKbuttin As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin&, "Button", vbNullString)
    PlayMess& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Call SendMessage(PlayMess&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlayMess&, WM_KEYUP, VK_SPACE, 0&)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin22& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin22&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    PlaySend& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Call SendMessage(PlaySend&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlaySend&, WM_KEYUP, VK_SPACE, 0&)

    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub


Sub Click_Go()
    Dim BuddyList As Long, GoButtin As Long, Click As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(GoButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(GoButtin&, WM_LBUTTONUP, 0, 0&)
End Sub


Sub Click_SavePass()
    Dim SignOnS As Long, PASS As Long, SavePass As Long

    SignOnS& = FindWindow("#32770", "Sign On")
    PASS& = FindWindowEx(SignOnS&, 0, "Button", vbNullString)
    SavePass& = FindWindowEx(SignOnS&, PASS&, "Button", vbNullString)
    Call SendMessage(SavePass&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(SavePass&, WM_KEYUP, VK_SPACE, 0&)
End Sub


Sub AIM_Hide_AoCom()
    Dim BuddyList As Long, TabGroup As Long, Rounder As Long
    Dim AoCom As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    Rounder& = FindWindowEx(TabGroup&, 0, "WndAte32Class", vbNullString)
    AoCom& = FindWindowEx(Rounder&, 0, "Ate32Class", vbNullString)
    Call ShowWindow(AoCom&, SW_HIDE)
End Sub


Sub AIM_Set_PW(PASS As String)
    Dim SignOnS2 As Long, password As Long, SetPassW As Long

    SignOnS2& = FindWindow("#32770", "Sign On")
    password& = FindWindowEx(SignOnS2&, 0, "Edit", vbNullString)
    SetPassW& = SendMessageByString(password&, WM_SETTEXT, 0, PASS$)
End Sub

Sub AIM_Set_SN(ScrnName As String)
    Dim SignOnS2 As Long, Comb As Long, CombEdit As Long
    Dim SetScreenN As Long

    SignOnS2& = FindWindow("#32770", "Sign On")
    Comb& = FindWindowEx(SignOnS2&, 0, "ComboBox", vbNullString)
    CombEdit& = FindWindowEx(Comb&, 0, "Edit", vbNullString)
    SetScreenN& = SendMessageByString(CombEdit&, WM_SETTEXT, 0, ScrnName$)
End Sub

Sub AIM_Show_AoCom()
    Dim BuddyList As Long, TabGroup As Long, Rounder As Long
    Dim AoCom As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    Rounder& = FindWindowEx(TabGroup&, 0, "WndAte32Class", vbNullString)
    AoCom& = FindWindowEx(Rounder&, 0, "Ate32Class", vbNullString)
    Call ShowWindow(AoCom&, SW_SHOW)
End Sub

Sub Click_SignOn()
    Dim SignOnS2 As Long, Help2 As Long, SetUp2 As Long
    Dim SignOn2 As Long, Klick As Long

    SignOnS2& = FindWindow("#32770", "Sign On")
    Help2& = FindWindowEx(SignOnS2&, 0, "_Oscar_IconBtn", vbNullString)
    SetUp2& = FindWindowEx(SignOnS2&, Help2&, "_Oscar_IconBtn", vbNullString)
    SignOn2& = FindWindowEx(SignOnS2&, SetUp2&, "_Oscar_IconBtn", vbNullString)
    Klick& = SendMessage(SignOn2&, WM_LBUTTONDOWN, 0, 0&)
    Klick& = SendMessage(SignOn2&, WM_LBUTTONUP, 0, 0&)
End Sub

Function List_ToString(TheList As ListBox) As String
' by dos
    Dim DoList As Long, MailString As String
    If TheList.List(0) = "" Then Exit Function
    For DoList& = 0 To TheList.ListCount - 1
        MailString$ = MailString$ & TheList.List(DoList&) & ", "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    List_ToString$ = MailString$
End Function
Sub Mass_Invite(lis As ListBox, say As String, Room As String)
    Dim ChatWindow As Long, Moo As String

    If lis.ListCount = 0 Then
        Exit Sub
    Else
    End If
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call Win_Killwin(ChatWindow&)

    Moo$ = List_ToString(lis)
    Call Send_Invite(Moo$, say, Room)
End Sub
Sub Click_Button2(TheButin As Long)
    Call PostMessage(TheButin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call PostMessage(TheButin&, WM_KEYUP, VK_SPACE, 0&)
End Sub


Sub EnterKey(TheWin As Long)
    Call SendMessageLong(TheWin&, WM_CHAR, ENTA, 0&)
End Sub


Sub Win_Enable(Window&)
    Dim dis
    dis = EnableWindow(Window&, 1)
End Sub

Sub Win_Disable(Window&)
    Dim dis
    dis = EnableWindow(Window&, 0)
End Sub
Sub AIM_SignOff_Close()
' This Signs off and closes it
    Dim BuddyList As Long, SignOnS As Long, KillNow As Integer
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Sign O&ff")

    SignOnS& = FindWindow("#32770", "Sign On")
    KillNow% = SendMessageByNum(SignOnS&, WM_CLOSE, 0, 0)
End Sub
Sub AIM_Close()
    Dim BuddyList As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)

    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call Win_Killwin(BuddyList&)
End Sub


Sub Chat_9liner(SayWhat As String)

    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Pause 0.5
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Pause 0.5
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
End Sub




Sub Chat_6liner(SayWhat As String)

    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Pause 0.5
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
End Sub
Sub AIM_Create_Profile(Text As String)
'Ex: Call AIM_Create_Profile("<b>Wee</b>, <i>my profile</i>")
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim CreateMain As Long, TheBut As Long
    Dim NextButtin As Long, CreatInt As Long, TheBut2 As Long
    Dim NextButtin2 As Long, ProfCre As Long, ProfCWin As Long
    Dim Borderz As Long, ProfText As Long, SetProfStng As Long
    Dim Back As Long, Cancel As Long, Finish As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "My &Profile...")

    CreateMain& = FindWindow("#32770", "Create a Profile - Searchable Directory")
    TheBut& = FindWindowEx(CreateMain&, 0, "Button", vbNullString)
    NextButtin& = FindWindowEx(CreateMain&, TheBut&, "Button", vbNullString)
    Call SendMessage(NextButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextButtin&, WM_KEYUP, VK_SPACE, 0&)
    CreatInt& = FindWindow("#32770", "Create a Profile - Chat Interests")
    TheBut2& = FindWindowEx(CreatInt&, 0, "Button", vbNullString)
    NextButtin2& = FindWindowEx(CreatInt&, TheBut2&, "Button", vbNullString)
    Call SendMessage(NextButtin2&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextButtin2&, WM_KEYUP, VK_SPACE, 0&)
    ProfCre& = FindWindow("#32770", "Create a Profile - More Info")
    ProfCWin& = FindWindowEx(ProfCre&, 0, "#32770", vbNullString)
    Borderz& = FindWindowEx(ProfCWin&, 0, "WndAte32Class", vbNullString)
    ProfText& = FindWindowEx(Borderz&, 0, "Ate32Class", vbNullString)
    SetProfStng& = SendMessageByString(ProfText&, WM_SETTEXT, 0, Text$)
    Back& = FindWindowEx(ProfCre&, 0, "Button", vbNullString)
    Cancel& = FindWindowEx(ProfCre&, Back&, "Button", vbNullString)
    Finish& = FindWindowEx(ProfCre&, Cancel&, "Button", vbNullString)
    Call SendMessage(Finish&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Finish&, WM_KEYUP, VK_SPACE, 0&)

    MsgBox "Your profile was created.", vbInformation, "Profile created."
End Sub
Sub AIM_SignOn()
    Select Case MsgBox("Is your name and password allready set?", vbYesNo + vbQuestion + vbDefaultButton2, "Signing on.")
        Case vbYes
            Dim SignOnS As Long, Help As Long, Setup As Long
            Dim SignOn As Long, Klick As Long

            SignOnS& = FindWindow("#32770", "Sign On")
            Help& = FindWindowEx(SignOnS&, 0, "_Oscar_IconBtn", vbNullString)
            Setup& = FindWindowEx(SignOnS&, Help&, "_Oscar_IconBtn", vbNullString)
            SignOn& = FindWindowEx(SignOnS&, Setup&, "_Oscar_IconBtn", vbNullString)
            Klick& = SendMessage(SignOn&, WM_LBUTTONDOWN, 0, 0&)
            Klick& = SendMessage(SignOn&, WM_LBUTTONUP, 0, 0&)
        Case vbNo
            Dim SignOnS2 As Long, ScrnName As String, Comb As Long
            Dim CombEdit As Long, SetScreenN As Long, PASS As String
            Dim password As Long, SetPassW As Long, Help2 As Long
            Dim SetUp2 As Long, SignOn2 As Long
            Dim BuddyList As Long

            SignOnS2& = FindWindow("#32770", "Sign On")
            ScrnName$ = InputBox("Please enter your screen name.", "Enter screen name for sign on.")
            Comb& = FindWindowEx(SignOnS2&, 0, "ComboBox", vbNullString)
            CombEdit& = FindWindowEx(Comb&, 0, "Edit", vbNullString)
            SetScreenN& = SendMessageByString(CombEdit&, WM_SETTEXT, 0, ScrnName$)

            PASS$ = InputBox("Please enter your password and you will be signed on.", "Enter password for sign on.")
            password& = FindWindowEx(SignOnS2&, 0, "Edit", vbNullString)
            Pause 0.3
            SetPassW& = SendMessageByString(password&, WM_SETTEXT, 0, PASS$)

            Call SendMessageLong(SetPassW&, WM_CHAR, VK_RETURN, 0&)
            Call SendMessageLong(SetPassW&, WM_CHAR, VK_RETURN, 0&)
            Help2& = FindWindowEx(SignOnS2&, 0, "_Oscar_IconBtn", vbNullString)
            SetUp2& = FindWindowEx(SignOnS2&, Help2&, "_Oscar_IconBtn", vbNullString)
            SignOn2& = FindWindowEx(SignOnS2&, SetUp2&, "_Oscar_IconBtn", vbNullString)
            Klick& = SendMessage(SignOn2&, WM_LBUTTONDOWN, 0, 0&)
            Klick& = SendMessage(SignOn2&, WM_LBUTTONUP, 0, 0&)
BuddyOn:

            BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
            If BuddyList& <> 0& Then
                MsgBox "You have been signed on and your screen name and password is now set and you only have to click 'Sign On' next session.", vbInformation, "Your on!"
                Exit Sub
            Else
                GoTo BuddyOn
            End If
    End Select
End Sub

Sub Chat_3liner(SayWhat As String)

    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
    Chat_Send (SayWhat$)
End Sub
Sub Chat_Close()
    Dim ChatWindow As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    If ChatWindow& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If
Start:
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call Win_Killwin(ChatWindow&)
End Sub

Function Chat_GetName2() As String
    Dim ChatWindow As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatWindow& <> 0& Then
    GoTo Start
    Else
        Chat_GetName2 = "[Not in room.]"
        Exit Function
    End If
Start:
    Dim GetsIt As String, Clear As String

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    GetsIt$ = Get_Caption(ChatWindow&)
    Clear$ = ReplaceString(GetsIt, "Chat Room: ", "")
    Chat_GetName2 = Clear$
End Function

Function Chat_Lang() As String
' Gets the Chats language
    Dim RoomInfo As Long, ChStat As Long, BrdStat As Long
    Dim RnStat As Long, lStat As Long, DVDStat As Long, LangStat As Long
    Dim GetIt As String, ChatWindow As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatWindow& <> 0& Then
    GoTo Start
    Else
        Chat_Lang = "[Not in room.]"
        Exit Function
    End If
Start:
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Chat Room Info...")
    RoomInfo& = FindWindow("AIM_CtlGroupWnd", "Chat Room Info")
    ChStat& = FindWindowEx(RoomInfo&, 0, "_Oscar_Static", vbNullString)
    BrdStat& = FindWindowEx(RoomInfo&, ChStat&, "_Oscar_Static", vbNullString)
    RnStat& = FindWindowEx(RoomInfo&, BrdStat&, "_Oscar_Static", vbNullString)
    lStat& = FindWindowEx(RoomInfo&, RnStat&, "_Oscar_Static", vbNullString)
    DVDStat& = FindWindowEx(RoomInfo&, lStat&, "_Oscar_Static", vbNullString)
    LangStat& = FindWindowEx(RoomInfo&, DVDStat&, "_Oscar_Static", vbNullString)
    GetIt$ = Get_Text(LangStat&)
    Chat_Lang = GetIt$

    Call Win_Killwin(RoomInfo&)
End Function
Function Chat_MaxMess() As Integer
' Gets the max message length
    Dim ChatWindow As Long, RoomInfo As Long, GetIt As String
    Dim ScStatik As Long, RnStatik As Long, LSatik As Long
    Dim MMLStatik As Long, DVDStatik As Long, DVD2Statik As Long
    Dim CaStatik As Long, LStatik As Long, MessLegn As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatWindow& <> 0& Then
    GoTo Start
    Else
        Chat_MaxMess = "[Not in room.]"
        Exit Function
    End If

Start:
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Chat Room Info...")

    RoomInfo& = FindWindow("AIM_CtlGroupWnd", "Chat Room Info")
    CaStatik& = FindWindowEx(RoomInfo&, 0, "_Oscar_Static", vbNullString)
    RnStatik& = FindWindowEx(RoomInfo&, CaStatik&, "_Oscar_Static", vbNullString)
    LStatik& = FindWindowEx(RoomInfo&, RnStatik&, "_Oscar_Static", vbNullString)
    MMLStatik& = FindWindowEx(RoomInfo&, LStatik&, "_Oscar_Static", vbNullString)
    DVDStatik& = FindWindowEx(RoomInfo&, MMLStatik&, "_Oscar_Static", vbNullString)
    DVD2Statik& = FindWindowEx(RoomInfo&, DVDStatik&, "_Oscar_Static", vbNullString)
    MessLegn& = FindWindowEx(RoomInfo&, DVD2Statik&, "_Oscar_Static", vbNullString)
    GetIt$ = Get_Text(MessLegn&)
    Chat_MaxMess = GetIt$

    Call Win_Killwin(RoomInfo&)

End Function

Function Chat_GetText() As String
    Dim ChatWindow As Long, BorderThing As Long, GetIt As String

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    BorderThing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    GetIt$ = Get_Text(BorderThing&)
    Chat_GetText = GetIt$
End Function

Function Chat_GetText_NOHTML() As String
    Dim ChatWindow As Long, BorderThing As Long, GetIt As String
    Dim Clear As String

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    BorderThing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    GetIt$ = Get_Text(BorderThing&)
    Clear$ = HTML_Remove(GetIt$)
    Chat_GetText_NOHTML = Clear$
End Function
Sub About()

End Sub
Function Chat_GetName() As String
    Dim ChatWindow As Long, RoomInfo As Long, RoomBox As Long
    Dim GetIt As String

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatWindow& <> 0& Then
    GoTo Start
    Else
        Chat_GetName = "[Not in room.]"
        Exit Function
    End If
Start:
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Chat Room Info...")

    RoomInfo& = FindWindow("AIM_CtlGroupWnd", "Chat Room Info")
    RoomBox& = FindWindowEx(RoomInfo&, 0, "Edit", vbNullString)
    GetIt$ = Get_Text(RoomBox&)
    Chat_GetName = GetIt$

    Call Win_Killwin(RoomInfo&)
End Function

Function Chat_NumUsers() As Integer
    Dim AimChatWin As Long, NumStat As Long, HowMany As String
    Dim Clear As String, Clear2 As Integer

    AimChatWin& = FindWindow("AIM_ChatWnd", vbNullString)
    NumStat& = FindWindowEx(AimChatWin&, 0, "_Oscar_Static", vbNullString)
    HowMany$ = Get_Text(NumStat)
    Clear$ = ReplaceString(HowMany$, " person here", "")
    Clear2% = ReplaceString(HowMany$, " people here", "")
    Chat_NumUsers = Clear2%
End Function

Sub Chat_Scroll_INFO()
    Dim MaxMess As String, NumbUser As Integer, ChatLang As String
    Dim GetNm As String

    MaxMess$ = Chat_MaxMess
    NumbUser% = Chat_NumUsers
    ChatLang$ = Chat_Lang
    GetNm$ = Chat_GetName
    Chat_Send ("[<b><I>Chat Name: </I></b>   " & Chat_GetName & "]")
    Chat_Send ("[<b><I>Chat Language: </I></b>   " & ChatLang$ & "]")
    Chat_Send ("[<b><I>Number of Users: </I></b>   " & Chat_NumUsers & "]")
    Chat_Send ("[<b><I>Max Message: </I></b>   " & MaxMess$ & "]")
End Sub

Sub Find_BuddyBy_Interest(Cbm1 As ComboBox)
' Please put all these in a Combo box for the user to
' select from, you HAVE to have a combo with all these
' in it for this to work

    'Books and Writing
    'Banking
    'Education
    'Engineering
    'Entrepeneurship
    'Finance
    'Law
    'Marketing
    'Medical
    'Performing Arts
    'Small Business
    'Cars
    'African -American
    'College Students
    'Hispanic
    'Seniors
    'Teens
    'Women
    'Computers and Technology
    'Business News
    'International News
    'Politics
    'Sports News
    'Fashion
    'Moms Online
    'Parenting
    'Pregnancy and Birth
    'Separation and Divorce
    'Games
    'Diseases
    'Fitness
    'Medicine
    'Antiques
    'Architecture
    'Astrology
    'Aviation
    'Civil War
    'Bird Watching
    'Coins
    'Crafts
    'Food
    'Gardening
    'Genealogy
    'Martial Arts
    'Pets
    'Photography
    'Science Fiction
    'Sewing and Needlecraft
    'Stamps
    'The Arts
    'Wines and Beer
    'Gardening
    'Home Decorating
    'Home Improvement
    'Bonds
    'Mutual Funds
    'Real Estate
    'Stocks
    'Taxes
    'Chat
    'Marriage
    'Romance
    'Movies
    'Alternative
    'Classical
    'Jazz
    'Rythm and Blues
    'Rock
    'Atheism
    'Buddhism
    'Christianity
    'Hinduism
    'Islam
    'Judaism
    'Auto Racing
    'Baseball
    'Basketball
    'Boating and Sailing
    'Boxing
    'Cycling
    'Fishing
    'Football
    'Golf
    'Hockey
    'Running
    'Scuba
    'Skiing and Boarding
    'Soccer
    'Swimming
    'Tennis
    'Women 's Sports
    'Cartoons and Comics
    'Celebrities
    'Comedy
    'Daytime Soaps
    'Talk Radio
    'Talk Shows
    'Family Travel
    'General Travel
    'International Travel

    Dim BuddyList As Long
    Dim FndBudWin As Long, InsideWin As Long
    Dim Email As Long, NameAdd As Long, ChatWith As Long
    Dim backBut As Long, NextBut As Long, FBBCI As Long, FBBCIwin As Long
    Dim TextBoxz As Long, Moo As String, TextSet As Long
    Dim BackBut2 As Long, NextBut2 As Long, FindFail As Long
    Dim FindFail2 As Long, FindFail3 As Long, KillNow As Integer

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
        Exit Sub
    End If
Start:
    Call RunMenuByString(BuddyList&, "Find a Buddy &Wizard")

    FndBudWin& = FindWindow("#32770", "Find a Buddy")
    InsideWin& = FindWindowEx(FndBudWin&, 0, "#32770", vbNullString)
    Email& = FindWindowEx(InsideWin&, 0, "Button", vbNullString)
    NameAdd& = FindWindowEx(InsideWin&, Email&, "Button", vbNullString)
    ChatWith& = FindWindowEx(InsideWin&, NameAdd&, "Button", vbNullString)
    Call SendMessage(ChatWith&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(ChatWith&, WM_KEYUP, VK_SPACE, 0&)
    backBut& = FindWindowEx(FndBudWin&, 0, "Button", vbNullString)
    NextBut& = FindWindowEx(FndBudWin&, backBut&, "Button", vbNullString)
    Call SendMessage(NextBut&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextBut&, WM_KEYUP, VK_SPACE, 0&)

    FBBCI& = FindWindow("#32770", "Find a Buddy By Common Interest")
    FBBCIwin& = FindWindowEx(FBBCI&, 0, "#32770", vbNullString)
    TextBoxz& = FindWindowEx(FBBCIwin&, 0, "Edit", vbNullString)
    Moo$ = Cbm1.Text
    TextSet& = SendMessageByString(TextBoxz&, WM_SETTEXT, 0, Moo$)
    BackBut2& = FindWindowEx(FBBCI&, 0, "Button", vbNullString)
    NextBut2& = FindWindowEx(FBBCI&, BackBut2&, "Button", vbNullString)
    Call SendMessage(NextBut2&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextBut2&, WM_KEYUP, VK_SPACE, 0&)

    FindFail& = FindWindow("#32770", "Find a Buddy - Failure")
    FindFail2& = FindWindowEx(FindFail&, 0, "#32770", vbNullString)
    FindFail3& = FindWindowEx(FindFail2&, 0, "Static", vbNullString)

    If FindFail& <> 0& And FindFail2& <> 0& And FindFail3& <> 0& Then
        KillNow% = SendMessageByNum(FindFail&, WM_CLOSE, 0, 0)
        MsgBox "Sorry, no one found in AIM database that likes what you like.", vbExclamation, "Interest search."
    Else
        Exit Sub
    End If
End Sub

Function Find_SignOn()
    Dim SignOnS As Long

    SignOnS& = FindWindow("#32770", "Sign On")
    Find_SignOn = SignOnS&
End Function

Sub Get_MemInfo(Who As String)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
    Dim InfoWin As Long, Edgz As Long
    Dim DropDown As Long, SetWho As Long, OKbuttin As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Get Member Inf&o")

    InfoWin& = FindWindow("_Oscar_Locate", "Buddy Info: ")
    Edgz& = FindWindowEx(InfoWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    DropDown& = FindWindowEx(Edgz&, 0, "Edit", vbNullString)
    SetWho& = SendMessageByString(DropDown&, WM_SETTEXT, 0, Who$)
    OKbuttin& = FindWindowEx(InfoWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Function IM_GetSn() As String
' This gets the screenname of the person your talking to

    Dim IMWin As Long, GetIt As String, Clear As String

    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    GetIt$ = Get_Caption(IMWin&)
    Clear$ = ReplaceString(GetIt$, " - Instant Message", "")
    IM_GetSn = Clear$
End Function

Function IM_GetText() As String
    Dim AIMim As Long, THing As Long, GetIt As String
    
    AIMim& = FindWindow("AIM_IMessage", vbNullString)
    THing& = FindWindowEx(AIMim&, 0, "WndAte32Class", vbNullString)
    GetIt$ = Get_Text(THing&)
    IM_GetText = GetIt$
End Function

Function IM_GetText_NOHTML() As String
    Dim AIMim As Long, THing As Long, GetIt As String, Clear As String
    
    AIMim& = FindWindow("AIM_IMessage", vbNullString)
    THing& = FindWindowEx(AIMim&, 0, "WndAte32Class", vbNullString)
    GetIt$ = Get_Text(THing&)
    Clear$ = HTML_Remove(GetIt$)
    IM_GetText_NOHTML = Clear$
End Function

Sub Mass_IM(lis As ListBox, Txt As String)
    If lis.ListCount = 0 Then
        Exit Sub
    Else
    End If

    Dim Moo
    For Moo = 0 To lis.ListCount - 1
        Call IM_Send(lis.List(Moo), Txt, True)
    Next Moo
End Sub

Sub INI_WriteTo(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, (Key$), KeyValue$, Directory$)
End Sub

Function INI_GetFrom(Section As String, Key As String, Directory As String) As String
   Dim Buff As String

   Buff = String(750, Chr(0))
   Key$ = (Key$)
   INI_GetFrom$ = Left(Buff, GetPrivateProfileString(Section$, ByVal Key$, "", Buff, Len(Buff), Directory$))
End Function
Sub Msg_NoAccess(ProgramName As String)
    Dim Rand
    Randomize
     Rand = Int((Rnd * 3) + 1)
      If Rand = 1 Then MsgBox "You do not have access to this information.", vbCritical, "" & ProgramName & " [No access.]"
      If Rand = 2 Then MsgBox "You don't have access to this info.", vbCritical, "" & ProgramName & " [No access.]"
      If Rand = 3 Then MsgBox "You need access to view this information.", vbCritical, "" & ProgramName & " [No access.]"
End Sub

Sub Msg_NotOnError(ProgramName As String)
    Dim Rand
    Randomize
     Rand = Int((Rnd * 4) + 1)
      If Rand = 1 Then MsgBox "Please sign on first.", vbCritical, "" & ProgramName & " [Not online.]"
      If Rand = 2 Then MsgBox "It helps to be online to use this feature.", vbCritical, "" & ProgramName & " [Not online.]"
      If Rand = 3 Then MsgBox "Sign on.", vbCritical, "" & ProgramName & " [Not online.]"
      If Rand = 4 Then MsgBox "You must be online to use this feature", vbCritical, "" & ProgramName & " [Not online.]"
End Sub
Function List_Count(Lst As ListBox)
    Dim Moo As Integer

    Moo% = Lst.ListCount
    List_Count = Moo%
End Function
Sub Load_ComboBox(Path As String, Combo As ComboBox)
'Call Load_ComboBox("c:\windows\desktop\combo.cmb", Combo1)

    Dim What As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Combo.AddItem What$
    Wend
    Close #1
End Sub


Sub Msg_PasswordProtect(password As String, ProgName As String, FormToLoad As Form)
    Dim GetPass As String

    GetPass$ = InputBox("Please enter the password.", "Enter password to enter " & ProgName & " .")

    If GetPass$ = password$ Then
        FormToLoad.Show
    Else
        MsgBox "Wrong, sorry please try again or get proper access.", vbCritical, "" & ProgName & " Password wrong."
    End If
End Sub

Sub Msg_ShureExit(ProgName As String, TheFrm As Form)
    Select Case MsgBox("Are you shure you want to exit " & ProgName & "?", vbYesNo + vbQuestion + vbDefaultButton2, ProgName$ & " [Exit?]")
    Case vbYes
        Unload TheFrm
        End
        End
        Unload TheFrm
    Case vbNo
        Exit Sub
    End Select
End Sub

Sub Save_ComboBox(Path As String, Combo As ComboBox)
'Ex: Call Save_ComboBox("c:\windows\desktop\combo.cmb", combo1)

    Dim Savez As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Savez& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(Savez&)
    Next Savez&
    Close #1
End Sub
Sub Save_ListBox(Path As String, Lst As ListBox)
'Ex: Call Save_ListBox("c:\windows\desktop\list.lst", list1)

    Dim Listz As Long
    On Error Resume Next

    Open Path$ For Output As #1
    For Listz& = 0 To Lst.ListCount - 1
        Print #1, Lst.List(Listz&)
        Next Listz&
    Close #1
End Sub

Sub Load_ListBox(Path As String, Lst As ListBox)
'Ex: Call Load_ListBox("c:\windows\desktop\list.lst", list1)

    Dim What As String
    On Error Resume Next

    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Lst.AddItem What$
    Wend
    Close #1
End Sub
Sub Click_Buttin(DaButtin As Long)

    Call SendMessage(DaButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(DaButtin&, WM_KEYUP, VK_SPACE, 0&)
End Sub


Function Find_BuddyList()
    Dim BuddyList As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Find_BuddyList = BuddyList&
End Function
Sub Find_BuddyBy_NameAddress(First As String, Middle As String, last As String, Maiden As String, NickName As String, Street As String, city As String, State As String, Country As String, Zip As String)
' Ex:
' Call Find_BuddyBy_NameAddress("First", "Middle", "Last", "Maiden", "NickName", "Street", "City", "State", "Country", "Zip")
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim FindaBud As Long, FindBud As Long
    Dim Email As Long, NameAdd As Long, Nbut As Long, NextButtin As Long
    Dim Findname As Long, FindNWin As Long, FirstName As Long, Fndwinz As Long
    Dim TextSet As Long, MiddleName As Long, TextSet2 As Long, NickNames As Long
    Dim LastName As Long, TextSet3 As Long, Maidens As Long
    Dim TextSet33 As Long, NickName2 As Long, TextSet4 As Long
    Dim StreetName As Long, TextSet5 As Long, CityName As Long
    Dim TextSet6 As Long, StateName As Long, TextSet7 As Long
    Dim CountryName As Long, TextSet8 As Long, ZipCode As Long
    Dim TextSet9 As Long, Findwinz As Long, Nxt As Long, NextButtin2 As Long
    Dim Fails As Long, Fails2 As Long, Fails3 As Long, FindBwin As Long
    Dim NoInfo As Long, NoInfo2 As Long, NoInfo3 As Long


    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Find a Buddy &Wizard")

    FindaBud& = FindWindow("#32770", "Find a Buddy")
    FindBwin& = FindWindowEx(FindaBud&, 0, "#32770", vbNullString)
    Email& = FindWindowEx(FindBwin&, 0, "Button", vbNullString)
    NameAdd& = FindWindowEx(FindBwin&, Email&, "Button", vbNullString)
    Call SendMessage(NameAdd&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NameAdd&, WM_KEYUP, VK_SPACE, 0&)

    Nbut& = FindWindowEx(FindaBud&, 0, "Button", vbNullString)
    NextButtin& = FindWindowEx(FindaBud&, Nbut&, "Button", vbNullString)
    Call SendMessage(NextButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextButtin&, WM_KEYUP, VK_SPACE, 0&)

    Findname& = FindWindow("#32770", "Find a Buddy By Name and Address")
    FindNWin& = FindWindowEx(Findname&, 0, "#32770", vbNullString)
    FirstName& = FindWindowEx(FindNWin&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(FirstName&, WM_SETTEXT, 0, First$)
    MiddleName& = FindWindowEx(FindNWin&, FirstName&, "Edit", vbNullString)
    TextSet2& = SendMessageByString(MiddleName&, WM_SETTEXT, 0, Middle$)
    LastName& = FindWindowEx(FindNWin&, MiddleName&, "Edit", vbNullString)
    TextSet3& = SendMessageByString(LastName&, WM_SETTEXT, 0, last$)
    Maidens& = FindWindowEx(FindNWin&, LastName&, "Edit", vbNullString)
    TextSet33& = SendMessageByString(Maidens&, WM_SETTEXT, 0, Maiden$)
    NickNames& = FindWindowEx(FindNWin&, Maidens&, "Edit", vbNullString)
    TextSet4& = SendMessageByString(NickNames&, WM_SETTEXT, 0, NickName$)
    StreetName& = FindWindowEx(FindNWin&, NickNames&, "Edit", vbNullString)
    TextSet5& = SendMessageByString(StreetName&, WM_SETTEXT, 0, Street$)
    CityName& = FindWindowEx(FindNWin&, StreetName&, "Edit", vbNullString)
    TextSet6& = SendMessageByString(CityName&, WM_SETTEXT, 0, city$)
    StateName& = FindWindowEx(FindNWin&, CityName&, "Edit", vbNullString)
    TextSet7& = SendMessageByString(StateName&, WM_SETTEXT, 0, State$)
    CountryName& = FindWindowEx(FindNWin&, StateName&, "Edit", vbNullString)
    TextSet8& = SendMessageByString(CountryName&, WM_SETTEXT, 0, Country$)
    ZipCode& = FindWindowEx(FindNWin&, CountryName&, "Edit", vbNullString)
    TextSet9& = SendMessageByString(ZipCode&, WM_SETTEXT, 0, Zip$)

    Fndwinz& = FindWindow("#32770", "Find a Buddy By Name and Address")
    Nxt& = FindWindowEx(Fndwinz&, 0, "Button", vbNullString)
    NextButtin2& = FindWindowEx(Fndwinz&, Nxt&, "Button", vbNullString)
    Call SendMessage(NextButtin2&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextButtin2&, WM_KEYUP, VK_SPACE, 0&)

    NoInfo& = FindWindow("#32770", "Find a Buddy - Need More Information")
    NoInfo2& = FindWindowEx(NoInfo&, 0, "#32770", vbNullString)
    NoInfo3& = FindWindowEx(NoInfo2&, 0, "Static", vbNullString)

    If NoInfo& And NoInfo2& And NoInfo3& <> 0& Then
        Win_Killwin (NoInfo&)
        MsgBox "Please fill in all the boxes to complet the search.", vbExclamation, "More info needed."
        Exit Sub
    End If

    Fails& = FindWindow("#32770", "Find a Buddy - Failure")
    Fails2& = FindWindowEx(Fails&, 0, "#32770", vbNullString)
    Fails3& = FindWindowEx(Fails2&, 0, "Static", vbNullString)
    If Fails& And Fails2& And Fails3& <> 0& Then
        Win_Killwin (Fails&)
        MsgBox "No one found in AIM database that matches the info you provided.", vbExclamation, "Name and address not found."
    Else
        Exit Sub
    End If
End Sub
Sub Find_BuddyBy_Email(ThereEmail As String)
'Ex: Call Find_BuddyBy_Email("Person@doamin.com")
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim FindBud As Long, TheWin As Long
    Dim Email As Long, Nxbut As Long, NextButtin As Long
    Dim FindEm As Long, Ewin As Long, textbox As Long, TextSet As Long
    Dim NextB As Long, NextButtin2 As Long, EndUpNo As Long
    Dim Resultz As Long, Well As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Find a Buddy &Wizard")

    FindBud& = FindWindow("#32770", "Find a Buddy")
    TheWin& = FindWindowEx(FindBud&, 0, "#32770", vbNullString)
    Email& = FindWindowEx(TheWin&, 0, "Button", vbNullString)
    Call SendMessage(Email&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(Email&, WM_KEYUP, VK_SPACE, 0&)

    Nxbut& = FindWindowEx(FindBud&, 0, "Button", vbNullString)
    NextButtin& = FindWindowEx(FindBud&, Nxbut&, "Button", vbNullString)
    Call SendMessage(NextButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextButtin&, WM_KEYUP, VK_SPACE, 0&)

    FindEm& = FindWindow("#32770", "Find a Buddy by E-mail Address")
    Ewin& = FindWindowEx(FindEm&, 0, "#32770", vbNullString)
    textbox& = FindWindowEx(Ewin&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(textbox&, WM_SETTEXT, 0, ThereEmail$)
    NextB& = FindWindowEx(FindEm&, 0, "Button", vbNullString)
    NextButtin2& = FindWindowEx(FindEm&, NextB&, "Button", vbNullString)
    Call SendMessage(NextButtin2&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NextButtin2&, WM_KEYUP, VK_SPACE, 0&)

    EndUpNo& = FindWindow("#32770", "Find a Buddy by E-mail Address")
    Resultz& = FindWindowEx(EndUpNo&, 0, "#32770", vbNullString)
    Well& = FindWindowEx(Resultz&, 0, "Static", vbNullString)
    If EndUpNo& And Resultz& And Well& <> 0& Then
        Win_Killwin (EndUpNo&)
        MsgBox "No one found in AIM database matching that e-mail address.", vbExclamation, "E-mail not found."
    Else
        Exit Sub
    End If
End Sub

Function HTML_Remove(TheStrg As String) As String
'From Dos32.bas, changed a little by me
    TheStrg$ = ReplaceString(TheStrg$, "#000000", "")
    TheStrg$ = ReplaceString(TheStrg$, "#ff0000", "")
    TheStrg$ = ReplaceString(TheStrg$, "#ffffff", "")
    TheStrg$ = ReplaceString(TheStrg$, "BODY BGCOLOR=", "")
    TheStrg$ = ReplaceString(TheStrg$, "FONT COLOR=", "")
    TheStrg$ = ReplaceString(TheStrg$, "<FONT>", "")
    TheStrg$ = ReplaceString(TheStrg$, "</FONT>", "")
    TheStrg$ = ReplaceString(TheStrg$, "</B>", "")
    TheStrg$ = ReplaceString(TheStrg$, "<B>", "")
    TheStrg$ = ReplaceString(TheStrg$, "<BR>", "" & Chr$(13) + Chr$(10))
    TheStrg$ = ReplaceString(TheStrg$, "<HTML>", "")
    TheStrg$ = ReplaceString(TheStrg$, "HTML", "")
    TheStrg$ = ReplaceString(TheStrg$, "<""", "")
    TheStrg$ = ReplaceString(TheStrg$, """>", "")
    TheStrg$ = ReplaceString(TheStrg$, "</>", "")
    TheStrg$ = ReplaceString(TheStrg$, "#0000ff", "")
    HTML_Remove = TheStrg$
End Function


Function Find_IM()
    Dim IMWin As Long

    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Find_IM = IMWin&
End Function

Sub IM_Close()
    Dim IMWin As Long

    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    If IMWin& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If
Start:
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call Win_Killwin(IMWin&)
End Sub

Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
' By dos, from dos23.bas He gets all the credit for this one
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
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
Sub Pause(interval)
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub
Sub Win_Hide(TheWin As Long)
    Call ShowWindow(TheWin&, SW_HIDE)
End Sub

Sub Win_Killwin(TheWind&)
    Call PostMessage(TheWind&, WM_CLOSE, 0&, 0&)
End Sub
Function Get_Caption(TheWin)
' From Dos32.bas He gets fill credit
    Dim WindowLngth As Integer, WindowTtle As String, Moo As String
    
    WindowLngth% = GetWindowTextLength(TheWin)
    WindowTtle$ = String$(WindowLngth%, 0)
    Moo$ = GetWindowText(TheWin, WindowTtle$, (WindowLngth% + 1))
    Get_Caption = WindowTtle$
End Function
Function Get_Class(TheWin)
' From Dos32.bas He gets fill credit
    Dim Buffzz As String, GetClass As String
    
    Buffzz$ = String$(250, 0)
    GetClass$ = GetClassName(TheWin, Buffzz$, 250)
    Get_Class = Buffzz$
End Function
Function Get_Text(child)
' From Dos32.bas He gets fill credit
    Dim GetTrim As Integer, TrimSpace As String, GetString As String
    
    GetTrim% = SendMessageByNum(child, 14, 0&, 0&)
    TrimSpace$ = Space$(GetTrim)
    GetString$ = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
    Get_Text = TrimSpace$
End Function
Sub Click_Icon(TheIcon&)
    Dim Klick As Long
    
    Klick& = SendMessage(TheIcon&, WM_LBUTTONDOWN, 0, 0&)
    Klick& = SendMessage(TheIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Win_SetText(TheWin As Long, What As String)
    Call SendMessageByString(TheWin&, WM_SETTEXT, 0&, What$)
End Sub

Sub Win_Show(TheWin As Long)

    Call ShowWindow(TheWin&, SW_SHOW)
End Sub

Sub Win_StayOnTop(TheFrm As Form)
    Dim SetOnTop

    SetOnTop = SetWindowPos(TheFrm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub
Sub IM_Send(SendName As String, SayWhat As String, CloseIM As Boolean)
' My send IM comes with a little thing where you can eather close
' it or not close it....
' Ex: Call IM_Send("ThereSn","Sup man",True) <-- that closes the IM
' Put False to not close the IM, All the IM sends have the TRUE FALSE thing
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, SendButtin As Long, Click As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
    Pause 0.1
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
    Pause 0.1
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
    If CloseIM = True Then
        Win_Killwin (IMWin&)
    Else
        Exit Sub
    End If
End Sub
Sub Send_Invite(Who As String, Message As String, Room As String)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabThing As Long, ChtIcon As Long
    Dim ChatIcon As Long, ChatInvite As Long, ToWhoBox As Long, SetWho As Long
    Dim MessageBox As Long, RealBox As Long, SetMessage As Long
    Dim MesRoom As Long, EdBox As Long, RoomBox As Long, SetRoom As Long
    Dim SendIcon1 As Long, SendIcon2 As Long, SendIcon As Long, Click As Long
    Dim MesBox As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabThing& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    ChtIcon& = FindWindowEx(TabThing&, 0, "_Oscar_IconBtn", vbNullString)
    ChatIcon& = FindWindowEx(TabThing&, ChtIcon&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(ChatIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(ChatIcon&, WM_LBUTTONUP, 0, 0&)

    Pause 0.2
    
    ChatInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    ToWhoBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    SetWho& = SendMessageByString(ToWhoBox&, WM_SETTEXT, 0, Who$)
    
    MessageBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    RealBox& = FindWindowEx(ChatInvite&, MessageBox&, "Edit", vbNullString)
    SetMessage& = SendMessageByString(RealBox&, WM_SETTEXT, 0, Message$)
    MesBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    EdBox& = FindWindowEx(ChatInvite&, MesBox&, "Edit", vbNullString)
    RoomBox& = FindWindowEx(ChatInvite&, EdBox&, "Edit", vbNullString)
    SetRoom& = SendMessageByString(RoomBox&, WM_SETTEXT, 0, Room$)
    SendIcon1& = FindWindowEx(ChatInvite&, 0, "_Oscar_IconBtn", vbNullString)
    SendIcon2& = FindWindowEx(ChatInvite&, SendIcon1&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(ChatInvite&, SendIcon2&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendIcon&, WM_LBUTTONUP, 0, 0&)

End Sub
Sub Chat_Send(SayWhat As String)
    Dim ChatWindow As Long, THing As Long, Thing2 As Long
    Dim SetChatText As Long, Buttin As Long, Buttin2 As Long, Buttin3 As Long
    Dim SendButtin As Long, Click As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    THing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Thing2& = FindWindowEx(ChatWindow&, THing&, "WndAte32Class", vbNullString)
    SetChatText& = SendMessageByString(Thing2&, WM_SETTEXT, 0, SayWhat$)
    Buttin& = FindWindowEx(ChatWindow&, 0, "_Oscar_IconBtn", vbNullString)
    Buttin2& = FindWindowEx(ChatWindow&, Buttin&, "_Oscar_IconBtn", vbNullString)
    Buttin3& = FindWindowEx(ChatWindow&, Buttin2&, "_Oscar_IconBtn", vbNullString)
    SendButtin& = FindWindowEx(ChatWindow&, Buttin3&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
End Sub
Function Get_UserSN() As String
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Get_UserSN = "[ not online.]"
      Exit Function
    End If

Start:
    Dim GetIt As String, Clear As String
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    GetIt$ = Get_Caption(BuddyList&)
    Clear$ = ReplaceString(GetIt$, "'s Buddy List", "")
    Get_UserSN = Clear$
End Function

Sub Chat_Clear()
    Dim ChatWindow As Long, BorderThing As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    BorderThing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Call SendMessageByString(BorderThing&, WM_SETTEXT, 0, "")
End Sub
Sub Chat_Attention(SayWhat As String)
    Chat_Send ("       [<B><U>(]A.T.T.E.N.T.I.O.N[)</B></U>]")
    Chat_Send SayWhat$
    Chat_Send ("       [<B><U>(]A.T.T.E.N.T.I.O.N[)</B></U>]")
End Sub
Sub Chat_Link(Address As String, Text As String)
    Chat_Send "<A HREF=""" + Address$ + """>" + Text$ + ""
End Sub
Sub IM_Open()
    Dim BuddyList As Long, TabWin As Long, IMbuttin As Long, Click As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub RunMenuByString(Application, StringSearch)
' From Hix he gets full credit

    Dim ToSearch As Integer, MenuCount As Integer, FindString
    Dim ToSearchSub As Integer, MenuItemCount As Integer, GetString
    Dim SubCount As Integer, MenuString As String, GetStringMenu As Integer
    Dim MenuItem As Integer, RunTheMenu As Integer
    
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
Sub AIM_Exit()
    Dim BuddyList As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "E&xit")
End Sub
Sub AIM_SignOff()
    Dim BuddyList As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Sign O&ff")
End Sub
Sub AIM_Hide()
    Dim BuddyList As Long, X As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    X& = ShowWindow(BuddyList&, SW_HIDE)
End Sub
Sub AIM_Show()
    Dim BuddyList As Long, X As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    X& = ShowWindow(BuddyList&, SW_SHOW)
End Sub
Function Find_ChatRoom()
    Dim ChatWindow As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Find_ChatRoom = ChatWindow&
End Function
Function AIM_Online()
    Dim BuddyList As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    AIM_Online = BuddyList&
End Function
Sub GoTo_WebPage(Address As String)
' Takes you to a webpage threw the GoTo bar
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim STWbox As Long, SetAdd As Long
    Dim GoButtin As Long, Click As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    SetAdd& = SendMessageByString(STWbox&, WM_SETTEXT, 0, Address$)
    Pause 0.1
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(GoButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(GoButtin&, WM_LBUTTONUP, 0, 0&)
End Sub
Function Talker_Elite(strin As String)
'From Chaos232.bas
' Ex: Chat_EliteTalker (Text1.Text)
    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NextChrr As String, NewSent As String, NumSpc As Integer, Crapp As Integer
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    
    Do While NumSpc% <= lenth%
    DoEvents
    Let NumSpc% = NumSpc% + 1
    Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
    Let NextChrr$ = Mid$(inptxt$, NumSpc%, 2)
    If NextChrr$ = "ae" Then Let NextChrr$ = "æ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo send
    If NextChrr$ = "AE" Then Let NextChrr$ = "Æ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo send
    If NextChrr$ = "oe" Then Let NextChrr$ = "œ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo send
    If NextChrr$ = "OE" Then Let NextChrr$ = "Œ": Let NewSent$ = NewSent$ + NextChrr$: Let Crapp% = 2: GoTo send
    If Crapp% > 0 Then GoTo send
    
    If NextChr$ = "A" Then Let NextChr$ = "Å"
    If NextChr$ = "a" Then Let NextChr$ = "ã"
    If NextChr$ = "B" Then Let NextChr$ = "(3"
    If NextChr$ = "C" Then Let NextChr$ = "Ç"
    If NextChr$ = "c" Then Let NextChr$ = "¢"
    If NextChr$ = "D" Then Let NextChr$ = "|)"
    If NextChr$ = "d" Then Let NextChr$ = "ð"
    If NextChr$ = "E" Then Let NextChr$ = "Ê"
    If NextChr$ = "e" Then Let NextChr$ = "è"
    If NextChr$ = "f" Then Let NextChr$ = "ƒ"
    If NextChr$ = "H" Then Let NextChr$ = "h"
    If NextChr$ = "I" Then Let NextChr$ = "‡"
    If NextChr$ = "i" Then Let NextChr$ = "î"
    If NextChr$ = "k" Then Let NextChr$ = "|‹"
    If NextChr$ = "K" Then Let NextChr$ = "(«"
    If NextChr$ = "L" Then Let NextChr$ = "£"
    If NextChr$ = "M" Then Let NextChr$ = "(\/)"
    If NextChr$ = "m" Then Let NextChr$ = "‹v›"
    If NextChr$ = "N" Then Let NextChr$ = "(\)"
    If NextChr$ = "n" Then Let NextChr$ = "ñ"
    If NextChr$ = "O" Then Let NextChr$ = "Ø"
    If NextChr$ = "o" Then Let NextChr$ = "ö"
    If NextChr$ = "P" Then Let NextChr$ = "¶"
    If NextChr$ = "p" Then Let NextChr$ = "Þ"
    If NextChr$ = "r" Then Let NextChr$ = "®"
    If NextChr$ = "S" Then Let NextChr$ = "§"
    If NextChr$ = "s" Then Let NextChr$ = "$"
    If NextChr$ = "t" Then Let NextChr$ = "†"
    If NextChr$ = "U" Then Let NextChr$ = "Ú"
    If NextChr$ = "u" Then Let NextChr$ = "µ"
    If NextChr$ = "V" Then Let NextChr$ = "\/"
    If NextChr$ = "W" Then Let NextChr$ = "w"
    If NextChr$ = "w" Then Let NextChr$ = "w"
    If NextChr$ = "X" Then Let NextChr$ = "><"
    If NextChr$ = "x" Then Let NextChr$ = "×"
    If NextChr$ = "Y" Then Let NextChr$ = "¥"
    If NextChr$ = "y" Then Let NextChr$ = "ý"
    If NextChr$ = "!" Then Let NextChr$ = "¡"
    If NextChr$ = "?" Then Let NextChr$ = "¿"
    If NextChr$ = "." Then Let NextChr$ = "…"
    If NextChr$ = "," Then Let NextChr$ = "‚"
    If NextChr$ = "1" Then Let NextChr$ = "¹"
    If NextChr$ = "%" Then Let NextChr$ = "‰"
    If NextChr$ = "2" Then Let NextChr$ = "²"
    If NextChr$ = "3" Then Let NextChr$ = "³"
    If NextChr$ = "_" Then Let NextChr$ = "¯"
    If NextChr$ = "-" Then Let NextChr$ = "—"
    If NextChr$ = " " Then Let NextChr$ = " "
    If NextChr$ = "<" Then Let NextChr$ = "«"
    If NextChr$ = ">" Then Let NextChr$ = "»"
    If NextChr$ = "*" Then Let NextChr$ = "¤"
    If NextChr$ = "`" Then Let NextChr$ = "“"
    If NextChr$ = "'" Then Let NextChr$ = "”"
    If NextChr$ = "0" Then Let NextChr$ = "º"
    Let NewSent$ = NewSent$ + NextChr$

send:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
Chat_Send NewSent$
End Function
Sub Win_Playwav(FilePath As String)
    Dim SoundName As String, Pla, Flagz As Integer
    
    SoundName$ = FilePath$
    Flagz% = SND_ASYNC Or SND_NODEFAULT
    Pla = sndPlaySound(SoundName$, Flagz%)
End Sub
Sub Win_Center(frmz As Form)

    frmz.Top = (Screen.Height * 0.85) / 2 - frmz.Height / 2
    frmz.Left = Screen.Width / 2 - frmz.Width / 2
End Sub
Sub AIM_Load()
    Dim X As Long, NoFreeze As Integer
    
    X& = Shell("C:\Program Files\AIM95\aim.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Sub GoTo_Room(Room As String)
' This sub takes the user to a room without inviteing people
' Kinda like a enter room

    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
    Dim TabThing As Long, ChtIcon As Long
    Dim ChatIcon As Long, ChatInvite As Long, ToWhoBox As Long, SetWho As Long
    Dim MessageBox As Long, RealBox As Long, SetMessage As Long
    Dim MesRoom As Long, EdBox As Long, RoomBox As Long, SetRoom As Long
    Dim SendIcon1 As Long, SendIcon2 As Long, SendIcon As Long, Who As String
    Dim Click As Long, MesBox As Long
    Who$ = Get_UserSN
    If Who$ = "[Could not retrieve]" Then Exit Sub
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabThing& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    ChtIcon& = FindWindowEx(TabThing&, 0, "_Oscar_IconBtn", vbNullString)
    ChatIcon& = FindWindowEx(TabThing&, ChtIcon&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(ChatIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(ChatIcon&, WM_LBUTTONUP, 0, 0&)
    Pause 0.2
    
    ChatInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    ToWhoBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    SetWho& = SendMessageByString(ToWhoBox&, WM_SETTEXT, 0, Who$)
    
    MessageBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    RealBox& = FindWindowEx(ChatInvite&, MessageBox&, "Edit", vbNullString)
    SetMessage& = SendMessageByString(RealBox&, WM_SETTEXT, 0, "")
    
    MesBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    EdBox& = FindWindowEx(ChatInvite&, MesBox&, "Edit", vbNullString)
    RoomBox& = FindWindowEx(ChatInvite&, EdBox&, "Edit", vbNullString)
    SetRoom& = SendMessageByString(RoomBox&, WM_SETTEXT, 0, Room$)
    
    SendIcon1& = FindWindowEx(ChatInvite&, 0, "_Oscar_IconBtn", vbNullString)
    SendIcon2& = FindWindowEx(ChatInvite&, SendIcon1&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(ChatInvite&, SendIcon2&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Chat_StampsON_OFF()
    Dim ChatWindow As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "Timestamp")
End Sub
Sub IM_GetInfo()
' This gets the Info of the person you are talking to
' Note: only works on the TOP im!

    Dim IMWin As Long
    
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMWin&, "Info")
End Sub
Sub Chat_MacroKill()

    Chat_Send ("<b><u>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ ")
    Pause 0.1
    Chat_Send ("<b><u>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ ")
    Pause 0.1
    Chat_Send ("<b><u>@@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ @@@@@@@@@@ ")
End Sub
Sub IM_StampsON_OFF()
    Dim IMWin As Long
    
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMWin&, "Timestamp")
End Sub
Sub Change_BuddyCaption(NewCap As String)
' If you change caption some stuff might not work.
 
    Dim BuddyWin As Long, SetCap As String
    
    BuddyWin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    SetCap$ = SendMessageByString(BuddyWin&, WM_SETTEXT, 0, NewCap$)
End Sub
Sub Change_IMCaption(NewCap As String)
' If you change caption some stuff might not work.
 
    Dim IMCaption As Long, SetCap As String
    
    IMCaption& = FindWindow("AIM_IMessage", vbNullString)
    SetCap$ = SendMessageByString(IMCaption&, WM_SETTEXT, 0, NewCap$)
End Sub
Sub Change_ChatCaption(NewCap As String)
' If you change caption some stuff might not work.
 
    Dim ChatWindow As Long, SetCap As String
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    SetCap$ = SendMessageByString(ChatWindow&, WM_SETTEXT, 0, NewCap$)
End Sub
Sub Chat_SendBold(SayWhat As String)
    Dim ChatWindow As Long, THing As Long, Thing2 As Long
    Dim SetChatText As Long, Buttin As Long, Buttin2 As Long, Buttin3 As Long
    Dim SendButtin As Long, Click As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    THing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Thing2& = FindWindowEx(ChatWindow&, THing&, "WndAte32Class", vbNullString)
    SetChatText& = SendMessageByString(Thing2&, WM_SETTEXT, 0, "<B>" & SayWhat$)
    Buttin& = FindWindowEx(ChatWindow&, 0, "_Oscar_IconBtn", vbNullString)
    Buttin2& = FindWindowEx(ChatWindow&, Buttin&, "_Oscar_IconBtn", vbNullString)
    Buttin3& = FindWindowEx(ChatWindow&, Buttin2&, "_Oscar_IconBtn", vbNullString)
    SendButtin& = FindWindowEx(ChatWindow&, Buttin3&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Chat_SendUnderlined(SayWhat As String)
    Dim ChatWindow As Long, THing As Long, Thing2 As Long
    Dim SetChatText As Long, Buttin As Long, Buttin2 As Long, Buttin3 As Long
    Dim SendButtin As Long, Click As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    THing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Thing2& = FindWindowEx(ChatWindow&, THing&, "WndAte32Class", vbNullString)
    SetChatText& = SendMessageByString(Thing2&, WM_SETTEXT, 0, "<U>" & SayWhat$)
    Buttin& = FindWindowEx(ChatWindow&, 0, "_Oscar_IconBtn", vbNullString)
    Buttin2& = FindWindowEx(ChatWindow&, Buttin&, "_Oscar_IconBtn", vbNullString)
    Buttin3& = FindWindowEx(ChatWindow&, Buttin2&, "_Oscar_IconBtn", vbNullString)
    SendButtin& = FindWindowEx(ChatWindow&, Buttin3&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Chat_SendItalic(SayWhat As String)
    Dim ChatWindow As Long, THing As Long, Thing2 As Long
    Dim SetChatText As Long, Buttin As Long, Buttin2 As Long, Buttin3 As Long
    Dim SendButtin As Long, Click As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    THing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Thing2& = FindWindowEx(ChatWindow&, THing&, "WndAte32Class", vbNullString)
    SetChatText& = SendMessageByString(Thing2&, WM_SETTEXT, 0, "<I>" & SayWhat$)
    Buttin& = FindWindowEx(ChatWindow&, 0, "_Oscar_IconBtn", vbNullString)
    Buttin2& = FindWindowEx(ChatWindow&, Buttin&, "_Oscar_IconBtn", vbNullString)
    Buttin3& = FindWindowEx(ChatWindow&, Buttin2&, "_Oscar_IconBtn", vbNullString)
    SendButtin& = FindWindowEx(ChatWindow&, Buttin3&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Chat_SendStress(SayWhat As String)
    Dim ChatWindow As Long, THing As Long, Thing2 As Long
    Dim SetChatText As Long, Buttin As Long, Buttin2 As Long, Buttin3 As Long
    Dim SendButtin As Long, Click As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    THing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Thing2& = FindWindowEx(ChatWindow&, THing&, "WndAte32Class", vbNullString)
    SetChatText& = SendMessageByString(Thing2&, WM_SETTEXT, 0, "<B><I><U>" & SayWhat$)
    Buttin& = FindWindowEx(ChatWindow&, 0, "_Oscar_IconBtn", vbNullString)
    Buttin2& = FindWindowEx(ChatWindow&, Buttin&, "_Oscar_IconBtn", vbNullString)
    Buttin3& = FindWindowEx(ChatWindow&, Buttin2&, "_Oscar_IconBtn", vbNullString)
    SendButtin& = FindWindowEx(ChatWindow&, Buttin3&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Win_Maximize(THeWindow As Long)
    Dim max As Long

    max& = ShowWindow(THeWindow&, SW_MAXIMIZE)
End Sub
Sub Win_Minimize(THeWindow As Long)
    Dim Mini As Long

    Mini& = ShowWindow(THeWindow&, SW_MINIMIZE)
End Sub
Sub Load_Text(Txt As textbox, FilePath As String)
'Ex: Call load_Text(list1,"c:\windows\desktop\text.txt")

    Dim mystr As String, FilePath2 As String, textz As String, a As String
    
    Open FilePath2$ For Input As #1
    Do While Not EOF(1)
    Line Input #1, a$
        textz$ = textz$ + a$ + Chr$(13) + Chr$(10)
        Loop
        Txt = textz$
    Close #1
End Sub

Sub Save_Text(Txt As textbox, FilePath As String)
'Ex: Call Save_Text(list1,"c:\windows\desktop\text.txt")
    Dim FilePath3 As String
    
    Open FilePath3$ For Output As #1
        Print #1, Txt
    Close 1
End Sub
Sub Win_StartButtin()
    Dim WinShell As Long, StartButtin As Long, Klick As Long

    WinShell& = FindWindow("Shell_TrayWnd", "")
    StartButtin& = FindWindowEx(WinShell&, 0, "Button", vbNullString)

    Call SendMessage(StartButtin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(StartButtin&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Sub IM_SendLink(Who As String, Address As String, Text As String, Closez As Boolean)
'Call IM_SendLink("There SN", "http://www.hider.com", "kickass VB site GO here!", True)

    If Closez = True Then
        Call IM_Send(Who$, "<A HREF=""" + Address$ + """>" + Text$ + "", True)
    Else
        Call IM_Send(Who$, "<A HREF=""" + Address$ + """>" + Text$ + "", False)
    End If
End Sub
Sub IM_Send_Bold(SendName As String, SayWhat As String, CloseIM As Boolean)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, Click As Long, SendButtin As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
    Pause 0.1
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
    Pause 0.1
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, "<B>" & SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)

    If CloseIM = True Then
        Win_Killwin (IMWin&)
    Else
        Exit Sub
    End If
End Sub
Sub IM_Send_Italic(SendName As String, SayWhat As String, CloseIM As Boolean)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, Click As Long, SendButtin As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)

    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
    Pause 0.1
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
    Pause 0.1
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, "<I>" & SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
    If CloseIM = True Then
        Win_Killwin (IMWin&)
    Else
        Exit Sub
    End If
End Sub
Sub IM_Send_Underlined(SendName As String, SayWhat As String, CloseIM As Boolean)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, Click As Long, SendButtin As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)

    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
    Pause 0.1
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
    Pause 0.1
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, "<U>" & SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
    
    If CloseIM = True Then
        Win_Killwin (IMWin&)
    Else
        Exit Sub
    End If
End Sub
Sub IM_Send_Stress(SendName As String, SayWhat As String, CloseIM As Boolean)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, Click As Long, SendButtin As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
    Pause 0.1
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
    Pause 0.1
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, "<B><I><U>" & SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
    
    If CloseIM = True Then
        Win_Killwin (IMWin&)
    Else
        Exit Sub
    End If
End Sub
Function Talker_Dot(strin As String)
    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NumSpc As Integer, NewSent As String, Dotz As String
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + "•"
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    Dotz$ = NewSent$
    Chat_Send (Dotz$)
End Function
Function Talker_Link(strin As String)
    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NumSpc As Integer, NewSent As String, Link As String
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + "-"
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    Link$ = NewSent$
    Chat_Send (Link$)
End Function
Function Talker_Space(strin As String)
    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NumSpc As Integer, NewSent As String, Spac As String
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + " "
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    Spac$ = NewSent$
    Chat_Send (Spac$)
End Function
Sub IM_Clear()
    Dim AIMim As Long, THing As Long
    
    AIMim& = FindWindow("AIM_IMessage", vbNullString)
    THing& = FindWindowEx(AIMim&, 0, "WndAte32Class", vbNullString)
    Call SendMessageByString(THing&, WM_SETTEXT, 0, "")
End Sub
Function Talker_Slash(strin As String)
    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NumSpc As Integer, NewSent As String, Slah As String
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + "/"
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    Slah$ = NewSent$
    Chat_Send (Slah$)
End Function
Function Talker_Period(strin As String)
    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NumSpc As Integer, NewSent As String, Pero As String
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + "."
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    Pero$ = NewSent$
    Chat_Send (Pero$)
End Function
Sub IM_Send2(SendName As String, SayWhat As String, CloseIM As Boolean)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim IMWin As Long, ComboBox As Long
    Dim TextEditBox As Long, TextSet As Long, EditThing As Long
    Dim SendButtin As Long, TextSet2 As Long, Click As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Send &Instant Message")
    
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
    Pause 0.1
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)

    If CloseIM = True Then
        Win_Killwin (IMWin&)
    Else
        Exit Sub
    End If
End Sub
Sub Send_Invite_2(Who As String, Message As String, Room As String)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
    Dim ChatInvite As Long, ToWhoBox As Long, SetWho As Long
    Dim MessageBox As Long, RealBox As Long, SetMessage As Long
    Dim MesRoom As Long, EdBox As Long, RoomBox As Long, SetRoom As Long
    Dim SendIcon1 As Long, SendIcon2 As Long, SendIcon As Long, MesBox As Long
    Dim Click As Long
       
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Send &Buddy Chat Invitation")
    
    ChatInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    ToWhoBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    SetWho& = SendMessageByString(ToWhoBox&, WM_SETTEXT, 0, Who$)
    
    MessageBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    RealBox& = FindWindowEx(ChatInvite&, MessageBox&, "Edit", vbNullString)
    SetMessage& = SendMessageByString(RealBox&, WM_SETTEXT, 0, Message$)
    MesBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    EdBox& = FindWindowEx(ChatInvite&, MesBox&, "Edit", vbNullString)
    RoomBox& = FindWindowEx(ChatInvite&, EdBox&, "Edit", vbNullString)
    SetRoom& = SendMessageByString(RoomBox&, WM_SETTEXT, 0, Room$)
    SendIcon1& = FindWindowEx(ChatInvite&, 0, "_Oscar_IconBtn", vbNullString)
    SendIcon2& = FindWindowEx(ChatInvite&, SendIcon1&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(ChatInvite&, SendIcon2&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub AIM_Hide_AD()
    Dim BuddyList As Long, TheAdd As Long, X As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TheAdd& = FindWindowEx(BuddyList&, 0, "WndAte32Class", vbNullString)
    X& = ShowWindow(TheAdd&, SW_HIDE)
End Sub
Sub AIM_Show_AD()
    Dim BuddyList As Long, TheAdd As Long, X As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TheAdd& = FindWindowEx(BuddyList&, 0, "WndAte32Class", vbNullString)
    X& = ShowWindow(TheAdd&, SW_SHOW)
End Sub
Sub AIM_Hide_MyNews()
    Dim BuddyList As Long, TabThing As Long, NewsButtin As Long
    Dim X As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabThing& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    NewsButtin& = FindWindowEx(TabThing&, 0, "WndAte32Class", vbNullString)
    X& = ShowWindow(NewsButtin&, SW_HIDE)
End Sub
Sub AIM_Show_MyNews()
    Dim BuddyList As Long, TabThing As Long, NewsButtin As Long
    Dim X As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabThing& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    NewsButtin& = FindWindowEx(TabThing&, 0, "WndAte32Class", vbNullString)
    X& = ShowWindow(NewsButtin&, SW_SHOW)
End Sub
Sub IM_Send_Invite(Message As String, Room As String)
'This sub send an Invite to the person in the TOP im!
    Dim IMWin As Long, ChatInvite As Long, SetMessage As Long
    Dim MessageBox As Long, RealBox As Long
    Dim SetMessafe As Long, MesBox As Long, EdBox As Long
    Dim RoomBox As Long, SetRoom As Long, SendIcon1 As Long
    Dim SendIcon2 As Long, SendIcon As Long, Click As Long
    
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMWin&, "&Send Chat Invitation")
    
    ChatInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    MessageBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    RealBox& = FindWindowEx(ChatInvite&, MessageBox&, "Edit", vbNullString)
    SetMessage& = SendMessageByString(RealBox&, WM_SETTEXT, 0, Message$)
    MesBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    EdBox& = FindWindowEx(ChatInvite&, MesBox&, "Edit", vbNullString)
    RoomBox& = FindWindowEx(ChatInvite&, EdBox&, "Edit", vbNullString)
    SetRoom& = SendMessageByString(RoomBox&, WM_SETTEXT, 0, Room$)
    SendIcon1& = FindWindowEx(ChatInvite&, 0, "_Oscar_IconBtn", vbNullString)
    SendIcon2& = FindWindowEx(ChatInvite&, SendIcon1&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(ChatInvite&, SendIcon2&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendIcon&, WM_LBUTTONUP, 0, 0&)

End Sub
Sub IM_HideButtins()
    Dim TheIM As Long, THing As Long, WarnButtin As Long
    Dim Thing2 As Long, Thing3 As Long, BlockButtin As Long
    Dim X As Long
    
    TheIM& = FindWindow("AIM_IMessage", vbNullString)
    THing& = FindWindowEx(TheIM&, 0, "_Oscar_IconBtn", vbNullString)
    WarnButtin& = FindWindowEx(TheIM&, THing&, "_Oscar_IconBtn", vbNullString)
    Thing2& = FindWindowEx(TheIM&, 0, "_Oscar_IconBtn", vbNullString)
    Thing3& = FindWindowEx(TheIM&, Thing2&, "_Oscar_IconBtn", vbNullString)
    BlockButtin& = FindWindowEx(TheIM&, Thing3&, "_Oscar_IconBtn", vbNullString)
    X = ShowWindow(WarnButtin&, SW_HIDE)
    X = ShowWindow(BlockButtin&, SW_HIDE)
End Sub
Sub IM_ShowButtins()
    Dim TheIM As Long, THing As Long, WarnButtin As Long
    Dim Thing2 As Long, Thing3 As Long, BlockButtin As Long
    Dim X As Long
    
    TheIM& = FindWindow("AIM_IMessage", vbNullString)
    THing& = FindWindowEx(TheIM&, 0, "_Oscar_IconBtn", vbNullString)
    WarnButtin& = FindWindowEx(TheIM&, THing&, "_Oscar_IconBtn", vbNullString)
    Thing2& = FindWindowEx(TheIM&, 0, "_Oscar_IconBtn", vbNullString)
    Thing3& = FindWindowEx(TheIM&, Thing2&, "_Oscar_IconBtn", vbNullString)
    BlockButtin& = FindWindowEx(TheIM&, Thing3&, "_Oscar_IconBtn", vbNullString)
    X& = ShowWindow(WarnButtin&, SW_SHOW)
    X& = ShowWindow(BlockButtin&, SW_SHOW)
End Sub
Sub AIM_Hide_GoToBar()
    Dim BuddyList As Long, STWbox As Long, GoButtin As Long
    Dim X  As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    X& = ShowWindow(STWbox&, SW_HIDE)
    X& = ShowWindow(GoButtin&, SW_HIDE)
End Sub
Sub AIM_Show_GoToBar()
    Dim BuddyList As Long, STWbox As Long, GoButtin As Long
    Dim X  As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    X& = ShowWindow(STWbox&, SW_SHOW)
    X& = ShowWindow(GoButtin&, SW_SHOW)
End Sub
Sub Chat_InviteBuddy(Who As String, Message As String)
    Dim ChatWindow As Long, ChatInvite As Long, ToWhoBox As Long
    Dim SetWho As Long, MessageBox As Long, RealBox As Long
    Dim SetMessage As Long, SendIcon As Long, SendIcon1 As Long
    Dim SendIcon2 As Long, Click As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Invite a Buddy...")

    ChatInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    ToWhoBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    SetWho& = SendMessageByString(ToWhoBox&, WM_SETTEXT, 0, Who$)
    
    MessageBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    RealBox& = FindWindowEx(ChatInvite&, MessageBox&, "Edit", vbNullString)
    SetMessage& = SendMessageByString(RealBox&, WM_SETTEXT, 0, Message$)
    SendIcon1& = FindWindowEx(ChatInvite&, 0, "_Oscar_IconBtn", vbNullString)
    SendIcon2& = FindWindowEx(ChatInvite&, SendIcon1&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(ChatInvite&, SendIcon2&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub IM_Open2()
    Dim BuddyList As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call RunMenuByString(BuddyList&, "Send &Instant Message")
End Sub



Sub WinApp_ShellTo(TheExe As String)
    Dim Shellz As Long, NoFreeze As Long

    Shellz& = Shell(TheExe$, 1): NoFreeze& = DoEvents()
End Sub

Sub WinApp_Unload(TheFrm As Form)
Unload TheFrm
End Sub



Public Sub Addlist()

End Sub
