Attribute VB_Name = "Chaos232"
'Wuz Up niggie  I was gonna Quit making Bas
'files then all the sudden i saw decompiled
'Progs with my bas So I made another
'well my handle is not Chaos any more it is
'Slice
'But i made Total Chaos so i'll keep the bas
'Chaos
'I have so much more in here everfade color u
'can think of from ByteFade made By my Boy
'and Cryofade umm i got some weird stuff a Bot
'alot of stuff from my Progs Look at KNK's site
'for some codes like save text box's and scroll
'textbox's Please as soon as you use this Mail
'Me at Outletmag@hotmail or ProgerxVB@hotmail.com
'Peace


Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function Sendmessege Lib "user32" Alias "SendMessegeA" (ByValwMsg As Long, ByVal wParam As Long, param As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
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


Private Declare Function PutFocus Lib "user32" Alias "SetFocus" _
       (ByVal hwnd As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, _
       ByVal wMsg As Long, _
       ByVal wParam As Integer, _
       ByVal lParam As Long) As Long
       Private Const EM_LINESCROLL = &HB6

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



Const EM_UNDO = &HC7
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





Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Sub addroom(Lst As ListBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
namez$ = String$(256, " ")
Ret = AOLGetList(Index, namez$)
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)))
ADD_AOL_LB namez$, Lst
Next Index
end_addr:
Lst.RemoveItem Lst.ListCount - 1
i = GetListIndex(Lst, AOLGetUser())
If i <> -2 Then Lst.RemoveItem i
End Sub




Sub AOLSNReset(SN$, aoldir$, Replace$)
l0036 = Len(SN$)
Select Case l0036
Case 3
i = SN$ + "       "
Case 4
i = SN$ + "      "
Case 5
i = SN$ + "     "
Case 6
i = SN$ + "    "
Case 7
i = SN$ + "   "
Case 8
i = SN$ + "  "
Case 9
i = SN$ + " "
Case 10
i = SN$
End Select
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
End Select
X = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, X, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
Put #2, X + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
X = X + 32000
LF2 = LOF(2)
Close #2
If X > LF2 Then GoTo 301
Loop
301:
End Sub



Sub AOLIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub TB4(Number As Integer)
Aol% = FindWindow("AOL Frame25", vbNullString)
TB% = FindChildByClass(Aol%, "AOL Toolbar")
tc% = FindChildByClass(TB%, "_AOL_Toolbar")
td% = FindChildByClass(tc%, "_AOL_Icon")

If Number = 1 Then
    Call AOLIcon(td%)
    Exit Sub
End If

For T = 0 To Number - 2
td% = GetWindow(td%, 2)
Next T

Call AOLIcon(td%)

End Sub


Function AOLMDI()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(Aol%, "MDIClient")
End Function


Sub killwin(hwnd%)
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |Closes a chosen window                              | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim KillNow%
KillNow% = SendMessageByNum(hwnd%, WM_CLOSE, 0, 0)
End Sub


Function fader(thetext$)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 8
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    V$ = Mid$(G$, W + 4, 1)
    Q$ = Mid$(G$, W + 5, 1)
    X$ = Mid$(G$, W + 6, 1)
    Y$ = Mid$(G$, W + 7, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & T$ & "<FONT COLOR=" & Chr$(34) & "#DCDCDC" & Chr$(34) & ">" & V$ & "<FONT COLOR=" & Chr$(34) & "#C0C0C0" & Chr$(34) & ">" & Q$ & "<FONT COLOR=" & Chr$(34) & "#808080" & Chr$(34) & ">" & X$ & "<FONT COLOR=" & Chr$(34) & "#696969" & Chr$(34) & ">" & Y$
Next W
SendChat p$
End Function

Public Function AOLGetNewMail(Index) As String
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
mail% = FindChildByTitle(MDI%, AOLGetUser & "'s Online Mailbox")
tabd% = FindChildByClass(mail%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
AOLTree% = FindChildByClass(tabp%, "_AOL_Tree")

'de = sendmessage(aoltree%, LB_GETCOUNT, 0, 0)
txtlen% = SendMessageByNum(AOLTree%, LB_GETTEXTLEN, Index, 0&)
txt$ = String(txtlen% + 1, 0&)
X = SendMessageByString(AOLTree%, LB_GETTEXT, Index, txt$)
AOLGetNewMail = txt$
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

Function AOLGetUser()
On Error Resume Next
Aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function


Sub ADD_AOL_LB(itm As String, Lst As ListBox)
If Lst.ListCount = 0 Then
Lst.AddItem itm
Exit Sub
End If
Do Until XX = (Lst.ListCount)
Let diss_itm$ = Lst.List(XX)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub
Sub AOLversion()

Aol% = FindWindow("AOL Frame25", 0&)
Wel% = FindChildByTitle(Aol%, "Welcome, " + UserSN())
aol3% = FindChildByClass(Wel%, "RICHCNTL")
If aol3% = 0 Then AC_AOLVersion = 25: Exit Sub
If aol3% <> 0 Then
    If GetCaption(Aol%) <> "America Online" Then AC_AOLVersion = 3 Else AC_AOLVersion = 4
    End If
    End Sub
Function fadeBlackBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackBlue = Msg
End Function

Function fadeBlackGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
fadeBlackGreen = Msg
End Function

Function fadeBlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 220 / a
        f = E * b
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackGrey = Msg
End Function

Function fadeBlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackPurple = Msg
End Function

Function fadeBlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackRed = Msg
End Function

Function fadeBlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackYellow = Msg
End Function

Function fadeBlueBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueBlack = Msg
End Function

Function fadeBlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueGreen = Msg
End Function

Function fadeBluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBluePurple = Msg
End Function

Function fadeBlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueRed = Msg
End Function

Function fadeBlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueYellow = Msg
End Function

Function fadeGreenBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlack = Msg
End Function

Function fadeGreenBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlue = Msg
End Function

Function fadeGreenPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenPurple = Msg
End Function

Function fadeGreenRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenRed = Msg
End Function

Function fadeGreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenYellow = Msg
End Function

Function fadeGreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 220 / a
        f = E * b
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlack = Msg
End Function

Function fadeGreyBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlue = Msg
End Function

Function fadeGreyGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyGreen = Msg
End Function

Function fadeGreyPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyPurple = Msg
End Function

Function fadeGreyRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyRed = Msg
End Function

Function fadeGreyYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyYellow = Msg
End Function

Function fadePurpleBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBlack = Msg
End Function

Function fadePurpleBlue(Text1)
    
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255, 0, 255 - f)
        H = RGB2HEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBlue = Msg
End Function

Function fadePurpleGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleGreen = Msg
End Function

Function fadePurpleRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleRed = Msg
End Function

Function fadePurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(255 - f, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleYellow = Msg
End Function

Function fadeRedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlack = Msg
End Function

Function fadeRedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlue = Msg
End Function

Function fadeRedGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedGreen = Msg
End Function

Function fadeRedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedPurple = Msg
End Function

Function fadeRedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedYellow = Msg
End Function

Function fadeYellowBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlack = Msg
End Function

Function fadeYellowBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlue = Msg
End Function

Function fadeYellowGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowGreen = Msg
End Function

Function fadeYellowPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowPurple = Msg
End Function

Function fadeYellowRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 255 / a
        f = E * b
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowRed = Msg
End Function


'Pre-set 3 Color fade combinations begin here


Function fadeBlackBlueBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackBlueBlack = Msg
End Function

Function fadeBlackGreenBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackGreenBlack = Msg
End Function

Function fadeBlackGreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackGreyBlack = Msg
End Function

Function fadeBlackPurpleBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackPurpleBlack = Msg
End Function

Function fadeBlackRedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackRedBlack = Msg
End Function

Function fadeBlackYellowBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlackYellowBlack = Msg
End Function

Function fadeBlueBlackBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueBlackBlue = Msg
End Function

Function fadeBlueGreenBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueGreenBlue = Msg
End Function

Function fadeBluePurpleBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBluePurpleBlue = Msg
End Function

Function fadeBlueRedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueRedBlue = Msg
End Function

Function fadeBlueYellowBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeBlueYellowBlue = Msg
End Function

Function fadeGreenBlackGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlackGreen = Msg
End Function

Function fadeGreenBlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenBlueGreen = Msg
End Function

Function fadeGreenPurpleGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenPurpleGreen = Msg
End Function

Function fadeGreenRedGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenRedGreen = Msg
End Function

Function fadeGreenYellowGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreenYellowGreen = Msg
End Function

Function fadeGreyBlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlackGrey = Msg
End Function

Function fadeGreyBlueGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyBlueGrey = Msg
End Function

Function fadeGreyGreenGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyGreenGrey = Msg
End Function

Function fadeGreyPurpleGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyPurpleGrey = Msg
End Function

Function fadeGreyRedGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyRedGrey = Msg
End Function

Function fadeGreyYellowGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 490 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 255, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeGreyYellowGrey = Msg
End Function

Function fadePurpleBlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBlackPurple = Msg
End Function

Function fadePurpleBluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleBluePurple = Msg
End Function

Function fadePurpleGreenPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleGreenPurple = Msg
End Function

Function fadePurpleRedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleRedPurple = Msg
End Function

Function fadePurpleYellowPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(255 - f, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadePurpleYellowPurple = Msg
End Function

Function fadeRedBlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlackRed = Msg
End Function

Function fadeRedBlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedBlueRed = Msg
End Function

Function fadeRedGreenRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedGreenRed = Msg
End Function

Function fadeRedPurpleRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 0, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedPurpleRed = Msg
End Function

Function fadeRedYellowRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeRedYellowRed = Msg
End Function

Function fadeYellowBlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlackYellow = Msg
End Function

Function fadeYellowBlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowBlueYellow = Msg
End Function

Function fadeYellowGreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255, 255 - f)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowGreenYellow = Msg
End Function

Function fadeYellowPurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(f, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowPurpleYellow = Msg
End Function

Function fadeYellowRedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        E = 510 / a
        f = E * b
        If f > 255 Then f = (255 - (f - 255))
        G = RGB(0, 255 - f, 255)
        H = RGBtoHEX(G)
        Msg = Msg & "" & D
    Next b
    fadeYellowRedYellow = Msg
End Function


'Preset 2-3 color fade hexcode generator


Function fadeRGBtoHEX(RGB)
    a = Hex(RGB)
    b = Len(a)
    If b = 5 Then a = "0" & a
    If b = 4 Then a = "00" & a
    If b = 3 Then a = "000" & a
    If b = 2 Then a = "0000" & a
    If b = 1 Then a = "00000" & a
    fadeRGBtoHEX = a
End Function


'Form back color fade codes begin here
'Works best when used in the Form_Paint() sub


Sub FadeFormBlue(vForm As Form)
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

Sub addroomtotext(TheList As ListBox, Text As TextBox)
' addroomtotext list1, text1
Dim Y
Call addroom(TheList)
For Y = 0 To TheList.ListCount - 1
tt$ = tt$ + TheList.List(Y) + ","
Next Y
Timeout (0.01)
Text.Text = tt$

End Sub


Sub aol4_macroScroll(Text As String)
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    SendChat Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If Counter = 4 Then
        Timeout (2.9)
        Counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub

Sub aol4_SpiralScroll(txt As TextBox)
X = txt.Text
thastar:
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = MYSTR
SendChat "•[" + X + "]•"
If txt.Text = X Then
Exit Sub
End If
GoTo thastar

End Sub


Function ScrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Scrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)
'Full bas by eLeSsDee == eLeSsDee@mindless.com
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe

'adds the scrambled text to the full scrambled element
cityz:
Scrambled$ = Scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
Scrambled$ = Scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = Scrambled$

Exit Function
End Function

Function HyperLink(txt As String, URL As String)
HyperLink = ("<A HREF=" & Chr$(34) & Text2 & Chr$(34) & ">" & Text1 & "</A>")
End Function
Public Function AOLGetList(Index As Long, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = person$
End Function


Public Function AOLSupRoom()
IsUserOnline
If AOLIsOnline = 0 Then GoTo last
FindChatRoom
If AOLFindRoom = 0 Then GoTo last

On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = AOLFindRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call SendChat("SuP 2  " & person$)
Timeout (1)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function


Public Sub AOLClearChat()
getpar% = FindChatRoom()
child = FindChildByClass(getpar%, "RICHCNTL")
End Sub

Sub AOL40_Keyword(Keyword)
'This will send a keyword through AOL 4.o
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Tool2%, "_AOL_Icon")
For GetIcon = 1 To 20
icon% = GetWindow(icon%, 2)
Next GetIcon
Call Pause(0.05)
Call ClickIcon(icon%)
Do: DoEvents
MDI% = FindChildByClass(AOLWindow(), "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
Edit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
Icon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And Edit% <> 0 And Icon2% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, Keyword)
Call Timeout(0.05)
Call ClickIcon(Icon2%)
Call ClickIcon(Icon2%)
End Sub

Function AOLWindow()
'This sets focus on the AOL window
AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function


Function Chat_RoomName()
Call GetCaption(AOLFindChatRoom)
End Function

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

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function FindChatRoom()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Function UserSN()
On Error Resume Next
Aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub killwait()

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call Timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function IsUserOnline()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome,")
If Welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub



Sub ToChat(Chat)
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



Sub StayOnTop(theform As Form)
setwinontop = SetWindowPos(theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Public Function AOLFindRoom()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone
AOLFindRoom = 0
GoTo 50
firs% = GetWindow(MDI%, GW_CHILD)
While firs%
firs% = GetWindow(firs%, 2)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = GetWindow(listers%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
listerc% = FindChildByClass(firs%, "_AOL_Combobox")
If listers% And listere% And listerb% And listerc% Then GoTo bone

Wend

bone:
Room% = firs%
AOLFindRoom = Room%
50
End Function

Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub SendMail(Recipiants, Subject, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub MailMe(Recipiants, Subject, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, messege)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub MailPunt(Recipiants, Subject, Message)
Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Text1.Text)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Text2.Text)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(MDI%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function


Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getwintext% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

Sub IMBuddy(Recipiant, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
Buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If Buddy% = 0 Then
    AOL40_Keyword ("BuddyView")
    Do: DoEvents
    Loop Until Buddy% <> 0
End If

AOIcon% = FindChildByClass(Buddy%, "_AOL_Icon")

For l = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next l

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

Call AOL40_Keyword("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call Timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetChatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
ChatText = GetText(AORich%)
GetChatText = ChatText
End Function

Function LastChatLineWithSN()
ChatText$ = GetChatText

For FindChar = 1 To Len(ChatText$)

thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function

Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For z = 1 To 11
    If Mid$(ChatTrim$, z, 1) = ":" Then
        SN = Left$(ChatTrim$, z - 1)
    End If
Next z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
If person$ = UserSN Then GoTo Na
List1.AddItem person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub

Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub



Sub FormDance(M As Form)

'  This makes a form dance across the screen
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 5000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000

End Sub
Private Sub InitializeTextBoxSlow()

        
       'This routine assigns the string to the textbox text propert
       '     y
       '     'as the string is being built. This is the method that
       '     'the MS VBKB detailed. I named it InitializeTextBoxSlow.
       Dim i As Integer
       Dim j As Integer
       Text1.Text = ""
       lblStatus = "Performing slow load..."
        
       '     'just a pause to let the textbox and label update

              DoEvents

                            For i% = 1 To 100
                                   Text1.Text = Text1.Text + "This is line " + Str$(i%)
                                    
                                   '     'Add 10 words to a line of text.

                                          For j% = 1 To 10
                                                 Text1.Text = Text1.Text + " ...Word " + Str$(j%)
                                          Next j%

                                    
                                   '     'Force a carriage return and linefeed
                                   '     'VB3 users need to use chr$(13) & chr$(10)
                                   Text1.Text = Text1.Text + vbCrLf
                            Next i%

                     Text1.Text = Text1.Text
              End Sub


Private Sub InitializeTextBoxFast()

        
       'This routine assigns the string to temporary string variabl
       '     e
       '     'as the string is being built.
       Dim tmp As String
       Dim i As Integer
       Dim j As Integer
       Text1.Text = ""
       lblStatus = "Performing fast load..."
        
       '     'just a pause to let the textbox and label update

              DoEvents

                            For i% = 1 To 100
                                   tmp$ = tmp$ + "This is line " + Str$(i%)
                                    
                                   '     'Add 10 words to a line of text

                                          For j% = 1 To 10
                                                 tmp$ = tmp$ + " ...Word " + Str$(j%)
                                          Next j%

                                    
                                   '     'Force a carriage return and linefeed
                                   '     'VB3 users need to use chr$(13) & chr$(10)
                                   tmp$ = tmp$ + vbCrLf
                            Next i%

                      
                     '     'Now it's time to assign it to the text property.
                     Text1.Text = tmp$
                      
              End Sub


Function ScrollText&(TextBox As Control, vLines As Integer)

       Dim Success As Long
       Dim SavedWnd As Long
       Dim R As Long
       Dim Lines As Long
       'save the window handle of the control that currently has fo
       '     cus
       SavedWnd = Screen.ActiveControl.hwnd
       Lines& = vLines
        
       '     'Set the focus to the passed control (text control)
       TextBox.SetFocus
        
       '     'Scroll the lines.
       Success = SendMessageLong(TextBox.hwnd, EM_LINESCROLL, 0, Lines&)
        
       '     'Restore the focus to the original control
       R = PutFocus(SavedWnd)
        
       '     'Return the number of lines actually scrolled
       ScrollText& = Success
End Function

Function RemoveSpace(thetext$)
Dim Text$
Dim theloop%
Text$ = thetext$
For theloop% = 1 To Len(thetext$)
If Mid(Text$, theloop%, 1) = " " Then
Text$ = Left$(Text$, theloop% - 1) + Right$(Text$, Len(Text$) - theloop%)
theloop% = theloop% - 1
End If
Next
RemoveSpace = Text$
End Function


Function RGB2HEX(R, G, b)
Dim X%
Dim XX%
Dim Color%
Dim Divide
Dim Answer%
Dim Remainder%
Dim Configuring$
For X% = 1 To 3
If X% = 1 Then Color% = b
If X% = 2 Then Color% = G
If X% = 3 Then Color% = R
For XX% = 1 To 2
Divide = Color% / 16
Answer% = Int(Divide)
Remainder% = (10000 * (Divide - Answer%)) / 625

If Remainder% < 10 Then Configuring$ = Str(Remainder%) + Configuring$
If Remainder% = 10 Then Configuring$ = "A" + Configuring$
If Remainder% = 11 Then Configuring$ = "B" + Configuring$
If Remainder% = 12 Then Configuring$ = "C" + Configuring$
If Remainder% = 13 Then Configuring$ = "D" + Configuring$
If Remainder% = 14 Then Configuring$ = "E" + Configuring$
If Remainder% = 15 Then Configuring$ = "F" + Configuring$
Color% = Answer%
Next XX%
Next X%
Configuring$ = RemoveSpace(Configuring$)
RGB2HEX = Configuring$
End Function


Sub AOLSetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub DoubleClick(Button%)
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |This double clicks a button of your choice          | |                                                   | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim DoubleClickNow%
DoubleClickNow% = SendMessageByNum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub
Sub Answerbot()
'steps...
'1. in Timer1 tye Call FortuneBot
'2. make 2 command buttons
'3. in command button 1 type-
'Timer1.enbled = True
'AOLChatSend "Type /fortune to get your fortune"
'4. in the command button 2 type-
'Timer1.enabled = false
'AOLChatSend "Fortune Bot Off!"
FreeProcess
Timer1.interval = 1
On Error Resume Next
Dim last As String
Dim name As String
Dim a As String
Dim n As Integer
Dim X As Integer
DoEvents
a = LastChatLine
last = Len(a)
For X = 1 To last
name = Mid(a, X, 1)
Final = Final & name
If name = ":" Then Exit For
Next X
Final = Left(Final, Len(Final) - 1)
If Final = AOLGetUser Then
Exit Sub
Else
If InStr(a, "/Vv KoBe vV") Then
 SendChat (" Don't Waste Time on a Server")
Call Timeout(0.6)
End If
End If
End Sub

Sub ResetNew(SN As String, pth As String)
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
tru_sn = "NewUser" + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub



Function r_elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If nextChr$ = "A" Then Let nextChr$ = "Å"
If nextChr$ = "a" Then Let nextChr$ = "å"
If nextChr$ = "B" Then Let nextChr$ = "ß"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "Ð"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "h"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "K" Then Let nextChr$ = "(«"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "/\/\"
If nextChr$ = "m" Then Let nextChr$ = "‹v›"
If nextChr$ = "N" Then Let nextChr$ = "/\/"
If nextChr$ = "n" Then Let nextChr$ = "ñ"
If nextChr$ = "O" Then Let nextChr$ = "Ø"
If nextChr$ = "o" Then Let nextChr$ = "ö"
If nextChr$ = "P" Then Let nextChr$ = "¶"
If nextChr$ = "p" Then Let nextChr$ = "Þ"
If nextChr$ = "r" Then Let nextChr$ = "®"
If nextChr$ = "S" Then Let nextChr$ = "§"
If nextChr$ = "s" Then Let nextChr$ = "$"
If nextChr$ = "t" Then Let nextChr$ = "†"
If nextChr$ = "U" Then Let nextChr$ = "Ú"
If nextChr$ = "u" Then Let nextChr$ = "µ"
If nextChr$ = "V" Then Let nextChr$ = "\/"
If nextChr$ = "W" Then Let nextChr$ = "\\'"
If nextChr$ = "w" Then Let nextChr$ = "vv"
If nextChr$ = "X" Then Let nextChr$ = "><"
If nextChr$ = "x" Then Let nextChr$ = "×"
If nextChr$ = "Y" Then Let nextChr$ = "¥"
If nextChr$ = "y" Then Let nextChr$ = "ý"
If nextChr$ = "!" Then Let nextChr$ = "¡"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = "." Then Let nextChr$ = "…"
If nextChr$ = "," Then Let nextChr$ = "‚"
If nextChr$ = "1" Then Let nextChr$ = "¹"
If nextChr$ = "%" Then Let nextChr$ = "‰"
If nextChr$ = "2" Then Let nextChr$ = "²"
If nextChr$ = "3" Then Let nextChr$ = "³"
If nextChr$ = "_" Then Let nextChr$ = "¯"
If nextChr$ = "-" Then Let nextChr$ = "—"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "<" Then Let nextChr$ = "«"
If nextChr$ = ">" Then Let nextChr$ = "»"
If nextChr$ = "*" Then Let nextChr$ = "¤"
If nextChr$ = "`" Then Let nextChr$ = "“"
If nextChr$ = "'" Then Let nextChr$ = "”"
If nextChr$ = "0" Then Let nextChr$ = "º"
Let newsent$ = newsent$ + nextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
r_elite = newsent$

End Function
Function r_elite2(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If nextChr$ = "A" Then Let nextChr$ = "Å"
If nextChr$ = "a" Then Let nextChr$ = "ã"
If nextChr$ = "B" Then Let nextChr$ = "(3"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "|)"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "h"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "K" Then Let nextChr$ = "(«"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "/\/\"
If nextChr$ = "m" Then Let nextChr$ = "‹v›"
If nextChr$ = "N" Then Let nextChr$ = "/\/"
If nextChr$ = "n" Then Let nextChr$ = "ñ"
If nextChr$ = "O" Then Let nextChr$ = "Ø"
If nextChr$ = "o" Then Let nextChr$ = "ö"
If nextChr$ = "P" Then Let nextChr$ = "¶"
If nextChr$ = "p" Then Let nextChr$ = "Þ"
If nextChr$ = "r" Then Let nextChr$ = "®"
If nextChr$ = "S" Then Let nextChr$ = "§"
If nextChr$ = "s" Then Let nextChr$ = "$"
If nextChr$ = "t" Then Let nextChr$ = "†"
If nextChr$ = "U" Then Let nextChr$ = "Ú"
If nextChr$ = "u" Then Let nextChr$ = "µ"
If nextChr$ = "V" Then Let nextChr$ = "\/"
If nextChr$ = "W" Then Let nextChr$ = "\\'"
If nextChr$ = "w" Then Let nextChr$ = ""
If nextChr$ = "X" Then Let nextChr$ = "><"
If nextChr$ = "x" Then Let nextChr$ = "×"
If nextChr$ = "Y" Then Let nextChr$ = "¥"
If nextChr$ = "y" Then Let nextChr$ = "ý"
If nextChr$ = "!" Then Let nextChr$ = "¡"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = "." Then Let nextChr$ = "…"
If nextChr$ = "," Then Let nextChr$ = "‚"
If nextChr$ = "1" Then Let nextChr$ = "¹"
If nextChr$ = "%" Then Let nextChr$ = "‰"
If nextChr$ = "2" Then Let nextChr$ = "²"
If nextChr$ = "3" Then Let nextChr$ = "³"
If nextChr$ = "_" Then Let nextChr$ = "¯"
If nextChr$ = "-" Then Let nextChr$ = "—"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "<" Then Let nextChr$ = "«"
If nextChr$ = ">" Then Let nextChr$ = "»"
If nextChr$ = "*" Then Let nextChr$ = "¤"
If nextChr$ = "`" Then Let nextChr$ = "“"
If nextChr$ = "'" Then Let nextChr$ = "”"
If nextChr$ = "0" Then Let nextChr$ = "º"
Let newsent$ = newsent$ + nextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
SendChat newsent$

End Function


Function r_dots(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "•"
Let newsent$ = newsent$ + nextChr$
Loop
r_dots = newsent$

End Function


Function r_backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = Text3
Let lenth% = Len(Text3)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(Text3, numspc%, 1)
Let newsent$ = nextChr$ & newsent$
Loop
Text2.AddItem newsent$

End Function

Function r_hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
If nextChr$ = "A" Then Let nextChr$ = "a"
If nextChr$ = "E" Then Let nextChr$ = "e"
If nextChr$ = "I" Then Let nextChr$ = "i"
If nextChr$ = "O" Then Let nextChr$ = "o"
If nextChr$ = "U" Then Let nextChr$ = "u"
If nextChr$ = "b" Then Let nextChr$ = "B"
If nextChr$ = "c" Then Let nextChr$ = "C"
If nextChr$ = "d" Then Let nextChr$ = "D"
If nextChr$ = "z" Then Let nextChr$ = "Z"
If nextChr$ = "f" Then Let nextChr$ = "F"
If nextChr$ = "g" Then Let nextChr$ = "G"
If nextChr$ = "h" Then Let nextChr$ = "H"
If nextChr$ = "y" Then Let nextChr$ = "Y"
If nextChr$ = "j" Then Let nextChr$ = "J"
If nextChr$ = "k" Then Let nextChr$ = "K"
If nextChr$ = "l" Then Let nextChr$ = "L"
If nextChr$ = "m" Then Let nextChr$ = "M"
If nextChr$ = "n" Then Let nextChr$ = "N"
If nextChr$ = "x" Then Let nextChr$ = "X"
If nextChr$ = "p" Then Let nextChr$ = "P"
If nextChr$ = "q" Then Let nextChr$ = "Q"
If nextChr$ = "r" Then Let nextChr$ = "R"
If nextChr$ = "s" Then Let nextChr$ = "S"
If nextChr$ = "t" Then Let nextChr$ = "T"
If nextChr$ = "w" Then Let nextChr$ = "W"
If nextChr$ = "v" Then Let nextChr$ = "V"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "]" Then Let nextChr$ = "]"
If nextChr$ = "[" Then Let nextChr$ = "["
Let newsent$ = newsent$ + nextChr$
Loop
r_hacker = newsent$

End Function
Sub r_kahn()
Dim Firstletter, LastLetter, Middle
txtlen = Len(txt)
Firstletter = Left$(txt, 1)
LastLetter = Right$(txt, 1)
Middle = NotSure
withnofirst = Right$(txt, txtlen - 1)
nofirstlen = Len(withnofirst)
Withnofirstorlast = Left$(withnofirst, nofirstlen - 1)
Text_Encode = LastLetter & Withnofirstorlast & Firstletter
End Sub
Function r_link(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "—"
Let newsent$ = newsent$ + nextChr$
Loop
r_link = newsent$

End Function

Function r_html(strin As String)
'Returns the strin lagged
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "<html>"
Let newsent$ = newsent$ + nextChr$
Loop
r_html = newsent$

End Function

Function r_spaced(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + " "
Let newsent$ = newsent$ + nextChr$
Loop
r_spaced = newsent$

End Function
Public Sub AOLScrollList(Lst As ListBox)
For X% = 0 To List1.ListCount - 1
SendChat ("Scrolling Name [" & X% & "]" & List1.List(X%))
Timeout (0.75)
Next X%
End Sub
Sub WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
SendChat (p$)
End Sub

Sub EliteTalker(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "d" Then leet$ = "d"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If letter$ = "`" Then leet$ = "´"
    If letter$ = "!" Then leet$ = "¡"
    If letter$ = "?" Then leet$ = "¿"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q
SendChat (Made$)
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", " ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", " ")
End Sub

Sub MyASCII(PPP$)
G$ = WavY("ChAoS'§ Quick Lagger 4 AOL4 ")
l$ = WavY(" by ChAoS")
lo$ = WavY(PPP$ & "Loaded")
b$ = WavY("User: " & UserSN)
TI$ = CoLoRChaTBlueBlack(TrimTime)
V$ = CoLoRChaTBlueBlack("²·º")
FONTTT$ = "<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">"
SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & V$ & l$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & lo$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •")
Call Timeout(0.15)
SendChat (FONTTT$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & b$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> •" & TI$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Function WavYChaTRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
WavYChaTRG = p$
End Function
Function WavYChaTRedBlue(thetext As String)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & s$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
WavYChaTRB = p$
End Function

Sub Attention(thetext As String)
G$ = WavY("Nike Toolz for AOL4 ")
l$ = WavY(" by VB4 & Nike")
aa$ = WavY("Attention")
SendChat ("$AOLame4$ ATTENTION $AOLame4$")
Call Timeout(0.15)
SendChat (Text1.Text)
Call Timeout(0.15)
SendChat ("$AOLame4$ ATTENTION $AOLame4$")
Call Timeout(0.15)
SendChat ("<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & ">" & "‹›·´¯`·._.·• " & G$ & "v¹·¹" & l$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • " & aa$ & "<FONT COLOR=" & Chr$(34) & "#0" & Chr$(34) & "> • ")
End Sub

Sub Bot_EightBall()
Dim Lst As String
Dim Text As String
Dim cht As Integer
Dim txt As String
Dim nws As String
Dim who As String
Dim wht As String
Dim R As Integer
Dim E As Integer
Dim M As Integer
Dim X
Dim Y
Geno:
Y = UserSN()
Aol% = FindWindow("AOL Frame25", 0&)
cht = FindChildByClass(Aol%, "_AOL_View")
txt = WinCaption(cht)
If Lst = "" Then Lst = txt
If txt = Lst Then Exit Sub
Lst = txt
nws = LastChatLine(txt)
who = Mid(nws, 2, InStr(nws, ":") - 2)
wht = Mid(nws, Len(who) + 4, Len(nws) - Len(who))
If LCase(Trim(Trim(Y))) = LCase(Trim(Trim(who))) Then GoTo Geno
R = GetParent(cht)
E = FindChildByClass(R, "_AOL_Edit")
tixt = RandomNumber(11)
If tixt = "1" Then
tixt = "Looks doubtful."
ElseIf tixt = "2" Then: tixt = "Definately YES!"
ElseIf tixt = "3" Then: tixt = "Definately No!"
ElseIf tixt = "4" Then: tixt = "Not a  chance"
ElseIf tixt = "5" Then: tixt = "nO"
ElseIf tixt = "6" Then: tixt = "gen yeA!"
ElseIf tixt = "7" Then: tixt = "Response HaZey try again."
ElseIf tixt = "8" Then: tixt = "ProbabLee"
ElseIf tixt = "9" Then: tixt = "yep yep"
ElseIf tixt = "10" Then: tixt = "I'm not suRe"
ElseIf tixt = "11" Then: tixt = "AbsolootLee yeZ"

End If
Text = wht$
W = InStr(LCase$(Text), LCase$("if"))
If W <> 0 Then
SendChat "•––•^v^•{‡ " & who & ", The 8-ball say: " & tixt
Timeout 0.5
GoTo Geno
End If

End Sub


Sub KillGlyph()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub


Function CoLoRChaTBlueBlack(thetext As String)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
CoLoRChaT = p$
End Function
Function ColorChatRedGreen(thetext)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
ColorChatRedGreen = p$

End Function
Function ColorChatRedBlue(thetext)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & s$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
ColorChatRedBlue = p$

End Function

Function TrimTime()
b$ = Left$(Time$, 5)
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(b$, 3) & " " & Ap$
End Function
Function TrimTime2()
b$ = Time$
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(b$, 5) & " " & Ap$
End Function

Function EliteText(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    letter$ = ""
    letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If letter$ = "b" Then leet$ = "b"
    If letter$ = "c" Then leet$ = "ç"
    If letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If letter$ = "j" Then leet$ = ",j"
    If letter$ = "n" Then leet$ = "ñ"
    If letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If letter$ = "s" Then leet$ = "š"
    If letter$ = "t" Then leet$ = "†"
    If letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If letter$ = "w" Then leet$ = "vv"
    If letter$ = "y" Then leet$ = "ÿ"
    If letter$ = "0" Then leet$ = "Ø"
    If letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If letter$ = "B" Then leet$ = "ß"
    If letter$ = "C" Then leet$ = "Ç"
    If letter$ = "D" Then leet$ = "Ð"
    If letter$ = "E" Then leet$ = "Ë"
    If letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If letter$ = "N" Then leet$ = "Ñ"
    If letter$ = "O" Then leet$ = "Õ"
    If letter$ = "S" Then leet$ = "Š"
    If letter$ = "U" Then leet$ = "Û"
    If letter$ = "W" Then leet$ = "VV"
    If letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = letter$
    Made$ = Made$ & leet$
Next Q

EliteText = Made$

End Function

Sub MyName()
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::               :::       ::::::::::: ")
Call Timeout(0.15)
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::    :::::::    :::           :::")
Call Timeout(0.15)
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>:::   :::   :::   :::   :::           :::")
Call Timeout(0.15)
SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B> :::::::     :::::::    :::::::::     :::")
End Sub

Sub IMIgnore(TheList As ListBox)
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To TheList.ListCount
        If LCase$(TheList.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Function SNfromIM()

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient") '

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

Sub Playwav(file)
SoundName$ = file
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub KillModal()
MODAL% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(MODAL%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
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

Function WavY(thetext As String)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    R$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    s$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & R$ & "</sup>" & U$ & "<sub>" & s$ & "</sub>" & T$
Next W
WavY = p$

End Function

Sub Pause(interval)
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub

Sub centerform(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Sub RespondIM(Message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(IM%, "RICHCNTL")

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
List1.AddItem SNfromIM
List1.AddItem MessageFromIM
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
e2 = GetWindow(E, GW_HWNDNEXT) 'Send Text
E = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, Text1)
ClickIcon (E)
Call Timeout(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
E = FindChildByClass(IM%, "RICHCNTL")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (E)
End Sub

Function MessageFromIM()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

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


Sub Upchat()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(Aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Aol%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Sub UnUpchat()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(Aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(Aol%, 0)
End Sub

Sub HideAOL()
Aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(Aol%, 0)
End Sub

Sub ShowAOL()
Aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(Aol%, 5)
End Sub

