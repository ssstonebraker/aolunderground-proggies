Attribute VB_Name = "ELEMENT"
'What up? this is BraD, the author of ELEMENT¹
'and ELEMENT²!, thanks for downloading the
'2nd version of ELEMENT.bas!
'This bas is 100% pure API and every sub and function
'works as it should.  It has everything you need
'for a GREAT AOL 4.0 prog!  E-mail me at Bradley084@aol.com
'My AOL screen name is Bradley084 if you got
'questions or comments just IM me or E-mail me!
'                                        -BraD

' Dont worry about that shit below!
Option Explicit

Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "User32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "User32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "User32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "User32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "User32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

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

Public Sub AddRoomToCombobox(TheCombo As ComboBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                TheCombo.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
    If TheCombo.ListCount > 0 Then
        TheCombo.Text = TheCombo.List(0)
    End If
End Sub

'Pre-set 2 color fade combinations begin here


Function BlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlue = Msg
End Function

Function BlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreen = Msg
End Function

Function BlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 220 / a
        F = E * B
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGrey = Msg
End Function

Function BlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurple = Msg
End Function

Function BlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRed = Msg
End Function

Function BlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellow = Msg
End Function

Function BlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlack = Msg
End Function

Function BlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreen = Msg
End Function

Function BluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurple = Msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRed = Msg
End Function

Function BlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellow = Msg
End Function

Function GreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlack = Msg
End Function

Function GreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlue = Msg
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurple = Msg
End Function

Function GreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRed = Msg
End Function

Function GreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellow = Msg
End Function

Function GreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 220 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlack = Msg
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlue = Msg
End Function

Function GreyGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreen = Msg
End Function

Function GreyPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurple = Msg
End Function

Function GreyRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRed = Msg
End Function

Function GreyYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellow = Msg
End Function

Function PurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlack = Msg
End Function

Function PurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlue = Msg
End Function

Function PurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreen = Msg
End Function

Function PurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRed = Msg
End Function

Function PurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellow = Msg
End Function

Function RedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlack = Msg
End Function

Function RedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlue = Msg
End Function

Function RedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreen = Msg
End Function

Function RedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurple = Msg
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellow = Msg
End Function

Function YellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlack = Msg
End Function

Function YellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlue = Msg
End Function

Function YellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreen = Msg
End Function

Function YellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurple = Msg
End Function

Function YellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 255 / a
        F = E * B
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRed = Msg
End Function


'Pre-set 3 Color fade combinations begin here


Function BlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackBlueBlack = Msg
End Function

Function BlackGreenBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreenBlack = Msg
End Function

Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackGreyBlack = Msg
End Function

Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackPurpleBlack = Msg
End Function

Function BlackRedBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackRedBlack = Msg
End Function

Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlackYellowBlack = Msg
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueBlackBlue = Msg
End Function

Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueGreenBlue = Msg
End Function

Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BluePurpleBlue = Msg
End Function

Function BlueRedBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueRedBlue = Msg
End Function

Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    BlueYellowBlue = Msg
End Function

Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlackGreen = Msg
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenBlueGreen = Msg
End Function

Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenPurpleGreen = Msg
End Function

Function GreenRedGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenRedGreen = Msg
End Function

Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreenYellowGreen = Msg
End Function

Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlackGrey = Msg
End Function

Function GreyBlueGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyBlueGrey = Msg
End Function

Function GreyGreenGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyGreenGrey = Msg
End Function

Function GreyPurpleGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyPurpleGrey = Msg
End Function

Function GreyRedGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyRedGrey = Msg
End Function

Function GreyYellowGrey(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 490 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    GreyYellowGrey = Msg
End Function

Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBlackPurple = Msg
End Function

Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleBluePurple = Msg
End Function

Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleGreenPurple = Msg
End Function

Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleRedPurple = Msg
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    PurpleYellowPurple = Msg
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlackRed = Msg
End Function

Function RedBlueRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedBlueRed = Msg
End Function

Function RedGreenRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedGreenRed = Msg
End Function

Function RedPurpleRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedPurpleRed = Msg
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    RedYellowRed = Msg
End Function

Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlackYellow = Msg
End Function

Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowBlueYellow = Msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowGreenYellow = Msg
End Function

Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowPurpleYellow = Msg
End Function

Function YellowRedYellow(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        D = Right(c, 1)
        E = 510 / a
        F = E * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & D
    Next B
    YellowRedYellow = Msg
End Function


'Preset 2-3 color fade hexcode generator


Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function



Sub FadeFormGreen(vForm As Form)
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

Sub FadeFormGrey(vForm As Form)
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

Sub FadeFormPurple(vForm As Form)
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

Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub


'Variable color fade functions begin here


Function TwoColors(Text, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
    C1BAK = c1
    C2BAK = c2
    C3BAK = c3
    C4BAK = c4
    c = 0
    o = 0
    o2 = 0
    Q = 1
    Q2 = 1
    For X = 1 To Len(Text)
        BVAL1 = Red2 - Red1
        BVAL2 = Green2 - Green1
        BVAL3 = Blue2 - Blue1
        
        VAL1 = (BVAL1 / Len(Text) * X) + Red1
        VAL2 = (BVAL2 / Len(Text) * X) + Green1
        VAL3 = (BVAL3 / Len(Text) * X) + Blue1
        
        c1 = RGB2HEX(VAL1, VAL2, VAL3)
        c2 = RGB2HEX(VAL1, VAL2, VAL3)
        c3 = RGB2HEX(VAL1, VAL2, VAL3)
        c4 = RGB2HEX(VAL1, VAL2, VAL3)
        
        If c1 = c2 And c2 = c3 And c3 = c4 And c4 = c1 Then c = 1: Msg = Msg & "<FONT COLOR=#" + c1 + ">"
        If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
        
        If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        
        If Wavy = True Then
            If o2 = 1 Then Msg = Msg + "<SUB>"
            If o2 = 3 Then Msg = Msg + "<SUP>"
            Msg = Msg + Mid$(Text, X, 1)
            If o2 = 1 Then Msg = Msg + "</SUB>"
            If o2 = 3 Then Msg = Msg + "</SUP>"
            If Q2 = 2 Then
                Q = 1
                Q2 = 1
                If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
                If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
                If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
                If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
            End If
        ElseIf Wavy = False Then
            Msg = Msg + Mid$(Text, X, 1)
            If Q2 = 2 Then
            Q = 1
            Q2 = 1
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + c1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + c2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + c3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + c4 + ">"
        End If
        End If
nc:     Next X
    c1 = C1BAK
    c2 = C2BAK
    c3 = C3BAK
    c4 = C4BAK
    TwoColors = Msg
End Function

Function ThreeColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Red3, Green3, Blue3, Wavy As Boolean)

'This code is still buggy, use at your own risk

    D = Len(Text)
        If D = 0 Then GoTo TheEnd
        If D = 1 Then Fade1 = Text
    For X = 2 To 500 Step 2
        If D = X Then GoTo Evens
    Next X
    For X = 3 To 501 Step 2
        If D = X Then GoTo Odds
    Next X
Evens:
    c = D \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c)
    GoTo TheEnd
Odds:
    c = D \ 2
    Fade1 = Left(Text, c)
    Fade2 = Right(Text, c + 1)
TheEnd:
    LA1 = Fade1
    LA2 = Fade2
        If Wavy = True Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, True) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, True)
        If Wavy = False Then FadeA = TwoColors(Left(LA1, Len(LA1) - 1), Red1, Green1, Blue1, Red2, Green2, Blue2, False) + TwoColors(Right(LA1, 1), Red2, Green2, Blue2, Red2, Green2, Blue2, False)
        If Wavy = True Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, True) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, True)
        If Wavy = False Then FadeB = TwoColors(Left(LA2, Len(LA2) - 1), Red2, Green2, Blue2, Red3, Green3, Blue3, False) + TwoColors(Right(LA2, 1), Red3, Green3, Blue3, Red3, Green3, Blue3, False)
    Msg = FadeA + FadeB
    ThreeColors = Msg
End Function

Function RGB2HEX(R, G, B)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = B
        If X& = 2 Then Color& = G
        If X& = 3 Then Color& = R
        For xx& = 1 To 2
            Divide = Color& / 16
            answer& = Int(Divide)
            Remainder& = (10000 * (Divide - answer&)) / 625
            If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
            If Remainder& = 10 Then Configuring$ = "A" + Configuring$
            If Remainder& = 11 Then Configuring$ = "B" + Configuring$
            If Remainder& = 12 Then Configuring$ = "C" + Configuring$
            If Remainder& = 13 Then Configuring$ = "D" + Configuring$
            If Remainder& = 14 Then Configuring$ = "E" + Configuring$
            If Remainder& = 15 Then Configuring$ = "F" + Configuring$
            Color& = answer&
        Next xx&
    Next X&
    Configuring$ = TrimSpaces(Configuring$)
    RGB2HEX = Configuring$
End Function















Sub KillWait()

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
AOedit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOedit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub

Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub

Sub EliteTalker(word$)
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
ChatSend (Made$)
End Sub
Sub CenterForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Private Sub BoldChatSend(Boldchat)
ChatSend ("<b>" & Boldchat & "</b>")
End Sub

Public Sub MailOpenNew()
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
Public Sub MailOpenOld()
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
Public Sub MailOpenSent()
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
Public Function MailCountOld() As Long
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
Public Function MailCountSent() As Long
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

Sub Buddy_Invite(Person)
' This will send an Invite to a buddy chat to someone
' works good for a pinter if u use a timer
FreeProcess
On Error GoTo errhandler
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
bud% = FindChildByTitle(MDI%, "Buddy List Window")
E = FindChildByClass(bud%, "_AOL_Icon")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
ClickIcon (E)
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
errhandler:
Exit Sub
End Sub
Public Sub Stop_Button()
Do
DoEvents
Loop
End Sub
Sub Anti45MinTimer()
' Put this in a timer, put the interval to 100
' This presses ok really fast when the AOL 45 minute
' timer pops up!
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function

Function FindChildByTitle(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs%
FindChildByTitle = Room%
End Function

Public Sub Button(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub ChatIgnoreByIndex(index As Long)
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

Public Sub ChatIgnoreByName(name As String)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lIndex As Long
    Room& = FindRoom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ And LCase(ScreenName$) = LCase(name$) Then
                lIndex& = index&
                Call ChatIgnoreByIndex(lIndex&)
                DoEvents
                Exit Sub
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub



Public Function ChatLineMsg(TheChatLine As String) As String
    If InStr(TheChatLine, Chr(9)) = 0 Then
        ChatLineMsg = ""
        Exit Function
    End If
    ChatLineMsg = Right(TheChatLine, Len(TheChatLine) - InStr(TheChatLine, Chr(9)))
End Function
Public Function ChatLineSN(TheChatLine As String) As String
    If InStr(TheChatLine, ":") = 0 Then
        ChatLineSN = ""
        Exit Function
    End If
    ChatLineSN = Left(TheChatLine, InStr(TheChatLine, ":") - 1)
End Function
Public Sub ChatSend(Chat As String)
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Function CheckAlive(ScreenName As String) As Boolean
    ' if you dont know what this is then it checks
    ' to see if someones acct. is active or not
    Dim AOL As Long, MDI As Long, ErrorWindow As Long
    Dim ErrorTextWindow As Long, ErrorString As String
    Dim MailWindow As Long, NoWindow As Long, NoButton As Long
    Call SendMail("*, " & ScreenName$, "You alive?", "=)")
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Do
        DoEvents
        ErrorWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
        ErrorTextWindow& = FindWindowEx(ErrorWindow&, 0&, "_AOL_View", vbNullString)
        ErrorString$ = GetText(ErrorTextWindow&)
    Loop Until ErrorWindow& <> 0 And ErrorTextWindow& <> 0 And ErrorString$ <> ""
    If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(ScreenName$, " ", ""))) > 0 Then
        CheckAlive = False
    Else
        CheckAlive = True
    End If
    MailWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
    Call PostMessage(ErrorWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(MailWindow&, WM_CLOSE, 0&, 0&)
    DoEvents
    Do
        DoEvents
        NoWindow& = FindWindow("#32770", "America Online")
        NoButton& = FindWindowEx(NoWindow&, 0&, "Button", "&No")
    Loop Until NoWindow& <> 0& And NoButton& <> 0
    Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function CheckIfMaster() As Boolean
    ' This checks to see if someone is the
    ' master SN
    Dim AOL As Long, MDI As Long, pWindow As Long
    Dim pButton As Long, Modal As Long, mStatic As Long
    Dim mString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    Call Keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Parental Controls")
        pButton& = FindWindowEx(pWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pWindow& <> 0& And pButton& <> 0&
    Pause 0.3
    Do
        DoEvents
        Call PostMessage(pButton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(pButton&, WM_LBUTTONUP, 0&, 0&)
        Pause 0.8
        Modal& = FindWindow("_AOL_Modal", vbNullString)
        mStatic& = FindWindowEx(Modal&, 0&, "_AOL_Static", vbNullString)
        mString$ = GetText(mStatic&)
    Loop Until Modal& <> 0 And mStatic& <> 0& And mString$ <> ""
    mString$ = ReplaceString(mString$, Chr(10), "")
    mString$ = ReplaceString(mString$, Chr(13), "")
    If mString$ = "Set Parental Controls" Then
        CheckIfMaster = True
    Else
        CheckIfMaster = False
    End If
    Call PostMessage(Modal&, WM_CLOSE, 0&, 0&)
    DoEvents
    Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
End Function
Public Function CheckIMs(Person As String) As Boolean
    ' Checks someone's IMs, to see if they're
    ' on or off
    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim Available As Long, Available1 As Long, Available2 As Long
    Dim Available3 As Long, oWindow As Long, oButton As Long
    Dim oStatic As Long, oString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword("aol://9293:" & Person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
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
        oWindow& = FindWindow("#32770", "America Online")
        oButton& = FindWindowEx(oWindow&, 0&, "Button", "OK")
    Loop Until oWindow& <> 0& And oButton& <> 0&
    Do
        DoEvents
        oStatic& = FindWindowEx(oWindow&, 0&, "Static", vbNullString)
        oStatic& = FindWindowEx(oWindow&, oStatic&, "Static", vbNullString)
        oString$ = GetText(oStatic)
    Loop Until oStatic& <> 0& And Len(oString$) > 15
    If InStr(oString$, "is online and able to receive") <> 0 Then
        CheckIMs = True
    Else
        CheckIMs = False
    End If
    Call SendMessage(oButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(oButton&, WM_KEYUP, VK_SPACE, 0&)
    Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End Function

Public Sub ClickIcon(aIcon As Long)
    ' Dont worry about what this means!
    Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function DoubleText(MyString As String) As String
    ' Dont worry about this either!
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
Public Function FileExists(sFileName As String) As Boolean
    ' Checks to see if a certain file exists or not
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Function GetFileAttributes(TheFile As String) As Integer
    ' Gets the properties of a certain file
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function
Public Sub SetFileHidden(TheFile As String)
    ' Hides a file so its not visible, like say
    ' you made a prog with intro music and you
    ' didnt want that person to take ur .wavfile
    ' just hide it! LoL
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Public Sub SetFileNormal(TheFile As String)
    ' Dont wory about this one
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub
Public Sub SetFileReadOnly(TheFile As String)
    ' Opens a file's attributes and sets it
    ' as a read-only file
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub
Public Function IMwords() As String
    Dim Rich As Long
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    IMText$ = GetText(Rich&)
End Function

Public Sub IMUnIgnore(Person As String)
    ' =)
    Call InstantMessage("$IM_ON, " & Person$, "=)")
End Sub
Public Sub IM(Person As String, Message As String)
    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim SendButton As Long, OK As Long, Button As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
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
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Message$)
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or IM& = 0&
    If OK& <> 0& Then
        Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Sub AOLKeyword(KW As String)
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
Public Function LineChar(thetext As String, CharNum As Long) As String
    Dim TextLength As Long, NewText As String
    TextLength& = Len(thetext$)
    If CharNum& > TextLength& Then
        Exit Function
    End If
    NewText$ = Left(thetext$, CharNum&)
    NewText$ = Right(NewText$, 1)
    LineChar$ = NewText$
End Function
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
Public Function LineFromString(MyString As String, Line As Long) As String
    Dim theline As String, Count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    Count& = LineCount(MyString$)
    If Line& > Count& Then
        Exit Function
    End If
    If Line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then
            FSpot = Len(MyString$)
        End If
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function
Public Function ListToMailString(TheList As ListBox) As String
    Dim DoList As Long, MailString As String
    If TheList.List(0) = "" Then Exit Function
    For DoList& = 0 To TheList.ListCount - 1
        MailString$ = MailString$ & "(" & TheList.List(DoList&) & "), "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToMailString$ = MailString$
End Function
Public Sub Load2listboxes(Directory As String, ListA As ListBox, ListB As ListBox)
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
Public Sub LoadComboBox(ByVal Directory As String, Combo As ComboBox)
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
Public Sub Loadlistbox(Directory As String, TheList As ListBox)
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
Sub LoadText(txtLoad As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.Text = TextString$
End Sub
Public Sub MemberRoom(Room As String)
    ' Goes to a Town Square room
    Call Keyword("aol://2719:61-2-" & Room$)
End Sub
Public Sub TimeOut(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub PlayWavFile(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub
Public Sub PrivateRoom(Room As String)
    ' Think About it...
    Call Keyword("aol://2719:2-2-" & Room$)
End Sub
Public Function ProfileGet(ScreenName As String) As String
    ' Gets someones profile
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim ToolIcon As Long, DoThis As Long, sMod As Long
    Dim MDI As Long, pgWindow As Long, pgEdit As Long, pgButton As Long
    Dim pWindow As Long, pTextWindow As Long, pString As String
    Dim NoWindow As Long, OKButton As Long, CurPos As POINTAPI
    Dim WinVis As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
    ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
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
    Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(CurPos.X, CurPos.Y)
    Do
        DoEvents
        pgWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Get a Member's Profile")
        pgEdit& = FindWindowEx(pgWindow&, 0&, "_AOL_Edit", vbNullString)
        pgButton& = FindWindowEx(pgWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until pgWindow& <> 0& And pgEdit& <> 0& And pgButton& <> 0&
    Call SendMessageByString(pgEdit&, WM_SETTEXT, 0&, ScreenName$)
    Call SendMessage(pgButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(pgButton&, WM_LBUTTONUP, 0&, 0&)
    DoEvents
    Do
        DoEvents
        pWindow& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
        pTextWindow& = FindWindowEx(pWindow&, 0&, "_AOL_View", vbNullString)
        pString$ = GetText(pTextWindow&)
        NoWindow& = FindWindow("#32770", "America Online")
    Loop Until pWindow& <> 0& And pTextWindow& <> 0& Or NoWindow& <> 0&
    DoEvents
    If NoWindow& <> 0& Then
        OKButton& = FindWindowEx(NoWindow&, 0&, "Button", "OK")
        Call SendMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call SendMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = "< No Profile >"
    Else
        Call PostMessage(pWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(pgWindow&, WM_CLOSE, 0&, 0&)
        ProfileGet$ = pString$
    End If
End Function
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
    Dim TempString As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While Count& <= StringLength&
        Count& = Count& + 1
        NextChr$ = Mid$(TempString$, Count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    ReverseString$ = NewString$
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
Public Sub RunMenu(TopMenu As Long, SubMenu As Long)
    Dim AOL As Long, aMenu As Long, sMenu As Long, mnID As Long
    Dim mVal As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    sMenu& = GetSubMenu(aMenu&, TopMenu&)
    mnID& = GetMenuItemID(sMenu&, SubMenu&)
    Call SendMessageLong(AOL&, WM_COMMAND, mnID&, 0&)
End Sub
Public Sub RunMenuByString(SearchString As String)
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
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Public Sub Save2ListBoxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.List(SaveLists&) & "*" & ListB.List(SaveLists)
    Next SaveLists&
    Close #1
End Sub

Public Sub SaveComboBox(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub
Public Sub SaveListBox(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Public Sub Scroll(ScrollString As String)
    Dim CurLine As String, Count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindRoom& = 0 Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    Count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To Count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = Left(CurLine$, 92)
            End If
            Call ChatSend(CurLine$)
            Pause 0.7
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Sub

Public Sub MailSend(Person As String, subject As String, Message As String)
    ' Sends Mail!
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
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Message$)
    DoEvents
    Pause 0.2
    Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub SetText(Window As Long, Text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub
Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Function SwitchStrings(MyString As String, String1 As String, String2 As String) As String
    Dim TempString As String, Spot1 As Long, Spot2 As Long
    Dim Spot As Long, ToFind As String, ReplaceWith As String
    Dim NewSpot As Long, LeftString As String, RightString As String
    Dim NewString As String
    If Len(String2) > Len(String1) Then
        TempString$ = String1$
        String1$ = String2$
        String2$ = TempString$
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

Public Sub WaitForOk(Room As String)
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    Room$ = LCase(ReplaceString(Room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindRoom&)
        RoomTitle$ = LCase(ReplaceString(Room$, " ", ""))
        FullWindow& = FindWindow("#32770", "America Online")
        FullButton& = FindWindowEx(FullWindow&, 0&, "Button", "OK")
    Loop Until (FullWindow& <> 0& And FullButton& <> 0&) Or Room$ = RoomTitle$
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
Public Sub WindowHide(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub
Public Sub WindowShow(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub

Public Sub AddToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub


Sub AntiIdle()
' Put this in a timer and set interval to 100
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Public Function FindInfoWindow() As Long
    'I took this from another .bas for a sub [atomic.bas]
    Dim AOL As Long, MDI As Long, child As Long
    Dim AOLCheck As Long, AOLIcon As Long, AOLStatic As Long
    Dim AOLIcon2 As Long, AOLGlyph As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
        FindInfoWindow& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            AOLCheck& = FindWindowEx(child&, 0&, "_AOL_Checkbox", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            AOLGlyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLIcon2& = FindWindowEx(child&, AOLIcon&, "_AOL_Icon", vbNullString)
            If AOLCheck& <> 0& And AOLStatic& <> 0& And AOLGlyph& <> 0& And AOLIcon& <> 0& And AOLIcon2& <> 0& Then
                FindInfoWindow& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindInfoWindow& = child&
End Function
Public Function FindRoom() As Long
    ' Finds a chat room
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
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
Public Sub FormMove(TheForm As Form)
    ' Drags the form when there is no border
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub FormExitDown(TheForm As Form)
    ' I took this from movin.bas, pretty cool on an exit!
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) + 25))
    Loop Until TheForm.Top > 7200
End Sub
Public Sub ProgNotOnTop(FormName As Form)
    'Sets it so the form desnt stayontop
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub ProgOnTop(FormName As Form)
    ' Sets it so the form stays on top of all
    ' other windows
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Function GetCaption(WindowHandle As Long) As String
    ' Gets the caption of an object
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   ' Gets data from an .INI file
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Function GetList(WindowHandle As Long) As String
    ' Gets the list text
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, LB_GETTEXTLEN, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, LB_GETTEXT, TextLength& + 1, buffer$)
    GetListText$ = buffer$
End Function
Public Function GetText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetText$ = buffer$
End Function
Public Function ProgUser() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        ProgUser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                ProgUser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    ProgUser$ = ""
End Function
Public Sub IMIgnore(Person As String)
    ' You can do this on ur own, but itf you want
    ' it in ur prog go ahead, this is the only thing
    ' I could think of sorry =(
    Call InstantMessage("$IM_OFF, " & Person$, "=)")
End Sub
Public Function IMLastMsg() As String
    Dim Rich As Long, MsgString As String, Spot As Long
    Dim NewSpot As Long
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
    MsgString$ = GetText(Rich&)
    NewSpot& = InStr(MsgString$, Chr(9))
    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
    MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
    IMLastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function
Public Sub IManswer(Msg As String)
    Dim IM As Long, Rich As Long, Icon As Long
    IM& = FindIM&
    If IM& = 0& Then Exit Sub
    Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(IM&, Rich&, "RICHCNTL", vbNullString)
    Icon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Icon& = FindWindowEx(IM&, Icon&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, Msg$)
    DoEvents
    Call SendMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function imsender() As String
    Dim IM As Long, Caption As String
    Caption$ = GetCaption(FindIM&)
    If InStr(Caption$, ":") = 0& Then
        imsender$ = ""
        Exit Function
    Else
        imsender$ = Right(Caption$, Len(Caption$) - InStr(Caption$, ":") - 1)
    End If
End Function
Public Function IMFind() As Long
    'Looks for an Instant Message Window on AOL4
    Dim AOL As Long, MDI As Long, child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(child&)
    If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
        FindIM& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(child&)
            If InStr(Caption$, "Instant Message") = 1 Or InStr(Caption$, "Instant Message") = 2 Or InStr(Caption$, "Instant Message") = 3 Then
                FindIM& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindIM& = child&
End Function

