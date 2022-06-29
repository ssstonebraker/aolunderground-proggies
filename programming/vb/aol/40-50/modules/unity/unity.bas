Attribute VB_Name = "ScR"

' America Online 4.0 Bas File Named ScR for the Club ScR
' This file will be used to make Scribe 2.0
' Unity made this file

Declare Function sndPlaySoundA Lib "c:\WINDOWS\SYSTEM\WINMM.DLL" (ByVal lpszSoundName$, ByVal ValueFlags As Long) As Long

   Global Const SND_SYNC = &H0
   Global Const SND_ASYNC = &H1
   Global Const SND_NODEFAULT = &H2
   Global Const SND_LOOP = &H8
   Global Const SND_NOSTOP = &H10
' End of the Cool Sound Stuff
'All that Fun Declaration stuff

'Global Const CB_GETCOUNT = (WM_USER + 6)
Declare Sub releaseCapture Lib "User" ()
Declare Function getnextwindow1 Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer

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

Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, GW_CHILD)
   

While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function



Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Sub AntiIdle()

AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)


End Sub

Sub Ao4Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub

Sub Ao4Kill45min()
Do
a% = FindWindow("_Aol_Palette", 0&)
B% = FindChildByClass(a%, "_Aol_Icon")
Call Pause(0.001)
Loop Until B% <> 0
Ao4Click (B%)
End Sub

Sub Ao4KW(where$)
B% = FindWindow("Aol Frame25", 0&)
a% = FindChildByClass(B%, "_Aol_Toolbar")
c% = FindChildByClass(a%, "_Aol_Icon")
D% = GetWindow(c%, GW_HWNDNEXT)
e% = GetWindow(D%, GW_HWNDNEXT)
F% = GetWindow(e%, GW_HWNDNEXT)
G% = GetWindow(F%, GW_HWNDNEXT)
h% = GetWindow(G%, GW_HWNDNEXT)
i% = GetWindow(h%, GW_HWNDNEXT)
J% = GetWindow(i%, GW_HWNDNEXT)
K% = GetWindow(J%, GW_HWNDNEXT)
L% = GetWindow(K%, GW_HWNDNEXT)
M% = GetWindow(L%, GW_HWNDNEXT)
N% = GetWindow(M%, GW_HWNDNEXT)
O% = GetWindow(N%, GW_HWNDNEXT)
P% = GetWindow(O%, GW_HWNDNEXT)
Q% = GetWindow(P%, GW_HWNDNEXT)
R% = GetWindow(Q%, GW_HWNDNEXT)
S% = GetWindow(R%, GW_HWNDNEXT)
T% = GetWindow(S%, GW_HWNDNEXT)
u% = GetWindow(T%, GW_HWNDNEXT)
V% = GetWindow(u%, GW_HWNDNEXT)
w% = GetWindow(u%, GW_HWNDNEXT)
Y% = GetWindow(w%, GW_HWNDNEXT)
Ao4Click Y%
Do
aol = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(aol, "Keyword")
daedit = FindChildByClass(bah, "_AOL_Edit")
Pause (0.001)
Loop Until daedit <> 0
daedit = FindChildByClass(bah, "_AOL_Edit")
send daedit, where$
ico% = FindChildByClass(bah, "_AOL_Icon")
Ao4Click ico%
End Sub

Sub Ao4mailcenter()
aol% = FindWindow("AOL Frame25", 0&)
Toolbar% = FindChildByClass(aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
MCenter% = GetWindow(TooLBaRB%, 2)
Ao4Click MCenter%


End Sub

Sub Ao4openmail()
aol% = FindWindow("AOL Frame25", 0&)
Toolbar% = FindChildByClass(aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
Ao4Click TooLBaRB%


End Sub

Sub Ao4Sendtext(TextToSend$)

aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
firs% = GetWindow(MDI%, 5)
listers% = FindChildByClass(firs%, "RICHCNTL")
listere% = getnextwindow1(listers%, 2)
NXT% = getnextwindow1(listere%, 2)
NXT1% = getnextwindow1(NXT%, 2)
NXT2% = getnextwindow1(NXT1%, 2)
NXT3% = getnextwindow1(NXT2%, 2)
NXT4% = getnextwindow1(NXT3%, 2)
listerb% = FindChildByClass(firs%, "_AOL_Listbox")
LISTER1% = FindChildByClass(firs%, "_AOL_Combobox")
DoEvents
DoEvents
sndtext% = SendMessageByString(NXT4%, WM_SETTEXT, 0, TextToSend$)
DoEvents
DoEvents
SendNow% = SendMessageByNum(NXT4%, WM_CHAR, &HD, 0)
DoEvents
End Sub

Sub Ao4Title(NewTitle$)
aol% = FindWindow("AOL Frame25", 0&)
AOLSetText aol%, NewTitle$
End Sub

Sub Ao4writemail()
aol% = FindWindow("AOL Frame25", 0&)
Toolbar% = FindChildByClass(aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
Ao4Click TooLBaRB%


End Sub

Sub AOLIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLSetText(win, Txt)
'thetext% = SendMessageByString("AOL FRAME25", WM_SETTEXT, 0, "Team Paradise Online")


End Sub

Sub CenterForm(FRM As Form)
FRM.Left = Screen.Width / 2 - FRM.Width / 2
FRM.Top = Screen.Height / 2 - FRM.Height / 2
End Sub

Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function CoLoRChaTBlueBlack(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
CoLoRChaT = P$
End Function

Function ColorChatRedBlue(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
ColorChatRedBlue = P$

End Function

Function ColorChatRedGreen(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & R$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
ColorChatRedGreen = P$

End Function

Sub EliteTalker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
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
Ao4Sendtext (Made$)
End Sub

Sub findchat()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   findchatroom = Room%
Else:
   findchatroom = 0
End If

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

Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
'frees process of freezes in your program
'and other stuff that makes your program
'slow down.  Works great.

End Function

Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function GetLineCount(text)

theview$ = text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(text, Len(text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Function GetWindowDir()
buffer$ = String$(255, 0)
X = GetWindowsDirectory(buffer$, 255)
If Right$(buffer$, 1) <> "\" Then buffer$ = buffer$ + "\"
GetWindowDir = buffer$
End Function

Sub hideAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 0)
End Sub

Sub im(Person, message)

AA% = FindWindow("Aol Frame25", 0&)
a% = FindChildByTitle(AA%, "Buddy List Window")
B% = FindChildByClass(a%, "_Aol_Icon")
If a% = 0 Then Ao4KW "Buddy View"
Do
a% = FindChildByTitle(AA%, "Buddy List Window")
B% = FindChildByClass(a%, "_Aol_Icon")
Call Pause(0.001)
Loop Until a% <> 0
c% = GetWindow(B%, GW_HWNDNEXT)
D% = GetWindow(c%, GW_HWNDNEXT)
e% = GetWindow(D%, GW_HWNDNEXT)
Ao4Click D%
Do

'instant message part
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
im2 = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(im2, "_AOL_Edit")
aolrich% = FindChildByClass(im2, "RICHCNTL")
imsend% = FindChildByClass(im2, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(im2, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next sends

AOLIcon (imsend%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
im2 = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im2, WM_CLOSE, 0, 0): Exit Do
If im2 = 0 Then Exit Do
Loop
End Sub

Sub IMBuddy(Recipiant, message)

aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Sub IMIgnore(thelist As ListBox)
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
im2 = FindChildByTitle(MDI%, ">Instant Message From:")
If im2 <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(SNfromIM) Then
            Badim2 = im2
            IMRICH% = FindChildByClass(Badim2, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(Badim2, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub

Sub IMKeyword(Recipiant, message)

End Sub

Sub IMsOff2()

End Sub

Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Sub KeyWord(TheKeyWord As String)
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below

'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub

Sub killmodal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub



Sub Killspinner()
' Kills the annoying spinning AOL logo in the toobar
' on AOL 4.0
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub
Function LastChatLineWithSN()
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

Function sendimkw()


aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")

Call KeyWord("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call timeout(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Function

Sub killwait()
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)

End Sub

Sub killwin(windo)
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub



Sub mail(SendTo$, subject$, text$)
aol% = FindWindow("AOL Frame25", 0&)
Toolbar% = FindChildByClass(aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(Toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
Ao4Click TooLBaRB%

Do
X% = DoEvents()
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
hideit = ShowWindow(chatlist%, SW_HIDE)
Loop Until chatedit% <> 0
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
hideit = ShowWindow(chatlist%, SW_HIDE)
chatwin% = GetParent(chatlist%)
Button% = FindChildByClass(chatlist%, "_AOL_Icon")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, SendTo$)
blah% = GetWindow(chatedit%, GW_HWNDNEXT)
Good% = GetWindow(blah%, GW_HWNDNEXT)
bad% = GetWindow(Good%, GW_HWNDNEXT)
Sad% = GetWindow(bad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Sad%, WM_SETTEXT, 0, subject$)
rich = FindChildByClass(chatlist%, "RICHCNTL")
sndtext% = SendMessageByString(rich, WM_SETTEXT, 0, text$ & " ")
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
Button% = FindChildByClass(chatlist%, "_AOL_Icon")
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
Pause 0.2

End Sub

Function MessageFromIM()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")

im2 = FindChildByTitle(MDI%, ">Instant Message From:")
If im2 Then GoTo Greed
im2 = FindChildByTitle(MDI%, "  Instant Message From:")
If im2 Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(im2, "RICHCNTL")
'IMmessage = GetText(imtext%)
sn = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, sn) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

Sub ParentChange(Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub

Sub Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub PLAYsound(File)
'if you havent guessed this plays a wav file
'call playwave (app.path + "\yourwave.wav")

SoundName$ = File
   ValueFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySoundA(SoundName$, ValueFlags%)

End Sub

Function ReplaceText(text, charfind, charchange)
If InStr(text, charfind) = 0 Then
ReplaceText = text
Exit Function
End If

For Replace = 1 To Len(text)
thechar$ = Mid(text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function

Sub RespondIM(message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")

im2 = FindChildByTitle(MDI%, ">Instant Message From:")
If im2 Then GoTo Greed2
im2 = FindChildByTitle(MDI%, "  Instant Message From:")
If im2 Then GoTo Greed2
Exit Sub
Greed2:
e = FindChildByClass(im2, "RICHCNTL")

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e2 = GetWindow(e, GW_HWNDNEXT) 'Send Text
e = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
ClickIcon (e)
Call timeout(0.8)
im2 = FindChildByTitle(MDI%, "  Instant Message From:")
e = FindChildByClass(im2, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (e)
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

Sub send(chatedit, sill$)
sndtext = SendMessageByString(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub SendMail(Recipiants, subject, message)

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
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
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

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

Sub showaol()
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 5)
End Sub

Function SNfromIM()

aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient") '

im2 = FindChildByTitle(MDI%, ">Instant Message From:")
If im2 Then GoTo greed3
im2 = FindChildByTitle(MDI%, "  Instant Message From:")
If im2 Then GoTo greed3
Exit Function
greed3:
IMCap$ = GetCaption(im2)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function

        

Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function

Sub timeout(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub

Function TrimCharacter(thetext, chars)
'TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces(text)
If InStr(text, " ") = 0 Then
TrimSpaces = text
Exit Function
End If

For TrimSpace = 1 To Len(text)
thechar$ = Mid(text, TrimSpace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next TrimSpace

TrimSpaces = thechars$
End Function

Function UntilWindowClass(parentw, childhand)
GoBack:
DoEvents
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
GoTo GoBack
FindClassLike = 0

bone:
Room% = firs%
UntilWindowClass = Room%
End Function

Function UntilWindowTitle(parentw, childhand)
GoBac:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBac
FindWindowLike = 0

bone:
Room% = firs%
UntilWindowTitle = Room%

End Function

Sub upchat4()
aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
'Call EnableWindow(aol%, 1)
'Call EnableWindow(Upp%, 0)

End Sub

Sub usersn()
On Error Resume Next
aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
'usersn = User

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

Sub WaitWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
topmdi% = GetWindow(MDI%, 5)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
topmdi2% = GetWindow(MDI%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop

End Sub

Sub WavyChatBlueBlack(thetext)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
Ao4Sendtext (P$)
End Sub

Function WavYChaTRedBlue(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
WavYChaTRB = P$
End Function

Function WavYChaTRedGreen(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    R$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & R$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
WavYChaTRG = P$
End Function

Function WinCaption(win)
WinCaption = "Team Paradise Online"

WinTextLength% = GetWindowTextLength("AOL Frame25")
WinTitle$ = String$(hwndLength%, 0)
getWinText% = GetWindowText("AOL Frame25", WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

Function IFileExists(ByVal sFileName As String) As Integer
' |¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯|\
' |Checks if a file you chose exists                   | |
' |____________________________________________________| |
'  \_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\_\__\|
Dim TheFileLength As Integer
On Error Resume Next
TheFileLength = Len(Dir$(sFileName))
If Err Or TheFileLength = 0 Then
IFileExists = False
Else
IFileExists = True
End If
End Function
Sub stayontop2()
Call SetWindowPos(FRM.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub


Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Function IsUserOnline()
aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Sub KeyWord4(TheKeyWord As String)
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below

'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call timeout(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call timeout(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub

