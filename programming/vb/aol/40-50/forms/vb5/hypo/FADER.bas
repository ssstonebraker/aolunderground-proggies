Attribute VB_Name = "Fader"
Rem ############################################
Rem #                         GrEeTs                                      #
Rem ############################################

Rem This is the first fader.bas ever made I would like to
Rem thank GeM3,JoLt, and Insanity for the filez that helped
Rem put this together alot of things I took from
Rem them and edited and some I made myself
Rem Email me if you have found an error you can reach me
Rem at: oO_JMR_Oo@hotmail.com
Rem check out my grewps page at:
Rem http://members.xoom.com/JMR/
Rem This is made for 32 bit programmers for AOL 4.0 I hope
Rem you like it I spent a good amount of time on it

Rem T(-)l§ '//Á§ [V]Á[)È ß¥ <-----(-JMR-)-----<<

Rem                               []
Rem                               []
Rem                               []  £00//<
Rem                               []  [)0'//(\)
Rem                               []  (-)Èl2È
Rem                               []  //=0l2
Rem                               []  Ç0[)È
Rem                               []
Rem                           \¯¯¯¯¯/
Rem                             \     /
Rem                              \   /
Rem                                \/
Rem
Public stopper
Public textstring
Public quitflagmassim As Boolean
Public quitflagbust As Boolean

Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
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
Function sn()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
UsER = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
sn = UsER
End Function
Function UserOn()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Sub ClickIcon(icon%)
Klick% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Klick% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub IconClick(icon%)
Qlick% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Qlick% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Sub IMOff()
Call im("$im_off", " ")
End Sub
Sub IMOn()
Call im("$im_on", " ")
End Sub
Sub im(ScreenName, DaText)
SendKeys "^i"
Do: DoEvents
IMs% = FindChildByTitle(AOLMDI(), "Send Instant Message")
Edit% = FindChildByClass(IMs%, "_AOL_Edit")
rich% = FindChildByClass(IMs%, "RICHCNTL")
icon% = FindChildByClass(IMs%, "_AOL_Icon")
Loop Until Edit% <> 0 And rich% <> 0 And icon% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, ScreenName)
Call SendMessageByString(rich%, WM_SETTEXT, 0, DaText)
For X = 1 To 9
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next X
Call timeout(0.01)
IconClick (icon%)
End Sub
Sub ChangeCaption(newcaption)
'This will change AOL's caption
Call SetText(AOLsWindow(), newcaption)
End Sub
Sub FadeBluGrn(thetext As String)
a = Len(thetext)
For W = 1 To a Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#0000FF>" & ab$ & "<FONT COLOR=#001199>" & u$ & "<FONT COLOR=#002288>" & S$ & "<FONT COLOR=#00003377>" & T$ & "<FONT COLOR=#004466>" & Y$ & "<FONT COLOR=#005555>" & l$ & "<FONT COLOR=#006644>" & F$ & "<FONT COLOR=#007733>" & B$ & "<FONT COLOR=#008822>" & c$ & "<FONT COLOR=#009911>" & D$ & "<FONT COLOR=#008822>" & H$ & "<FONT COLOR=#007733>" & j$ & "<FONT COLOR=#006644>" & k$ & "<FONT COLOR=#005555>" & m$ & "<FONT COLOR=#004466>" & n$ & "<FONT COLOR=#003377>" & Q$ & "<FONT COLOR=#002288>" & V$ & "<FONT COLOR=#001199>" & Z$
Next W
SendIt (PC$)
End Sub
Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(mdi%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Public Function GetList(Index As Long, Buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If
Buffer$ = Person$
End Function
Sub ADDLB(itm As String, lst As ListBox)
'Add a list of names to a VB ListBox
If lst.ListCount = 0 Then
lst.AddItem itm
Exit Sub
End If
Do Until xx = (lst.ListCount)
Let diss_itm$ = lst.List(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
lst.AddItem itm
End Sub
Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
mdi% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
UsER = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLUserSN = UsER
End Function
Sub AddRoomList(lst As ListBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
    names$ = String$(256, " ")
    Ret = GetList(Index, names$)
    names$ = Left$(Trim$(names$), Len(Trim(names$)))
    ADDLB names$, lst
Next Index
endaddroom:
lst.RemoveItem lst.ListCount - 1
sns% = UserSN()
If i <> -2 Then lst.RemoveItem sns%
End Sub
Function FreeProcess()
'This will make it so your program doesn't freeze and clears
'many errors
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
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
Sub AntiIdle()
Modal% = FindWindow("_AOL_Modal", vbNullString)
icon% = FindChildByClass(Modal%, "_AOL_Icon")
IconClick (AOIcon%)
End Sub
Sub CenterForm(frm As Form)
frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub
Sub KillGlyph()
'Ya know that annoying spinning thing on AOL well this
'fixes that little pest
tool% = FindChildByClass(AOLsWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Glyph% = FindChildByClass(Toolbar%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub
Sub FadeCustom(thetext As String, Hx1 As Integer, Hx2 As Integer, Hx3 As Integer, Hx4 As Integer, Hx5 As Integer, Hx6 As Integer, Hx7 As Integer, Hx8 As Integer, Hx9 As Integer, Hx10 As Integer)
'Dont worry this is 18 hexes that can
'Be entered but I made it 10 CuZ
'it goes: 1,2,3,4,5,6,7,8,9,10,9,8,7,6,5,4,3,2
'this is so when it fades it will loop
'I entered this so you wouldnt have to delete and
'myne and edit your own or figure out how to
'a new sub
a = Len(thetext)
For W = 1 To a Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#" & Hex1 & ">" & ab$ & "<FONT COLOR=#" & Hx2 & ">" & u$ & "<FONT COLOR=#" & Hx3 & ">" & S$ & "<FONT COLOR=#" & Hx4 & ">" & T$ & "<FONT COLOR=#" & Hx5 & ">" & Y$ & "<FONT COLOR=#" & Hx6 & ">" & l$ & "<FONT COLOR=#" & Hx7 & ">" & F$ & "<FONT COLOR=#" & Hx8 & ">" & B$ & "<FONT COLOR=#" & Hx9 & ">" & c$ & "<FONT COLOR=#" & Hx10 & ">" & D$ & "<FONT COLOR=#" & Hx9 & ">" & H$ & "<FONT COLOR=#" & Hx8 & ">" & j$ & "<FONT COLOR=#" & Hx7 & ">" & k$ & "<FONT COLOR=#" & Hx6 & ">" & m$ & "<FONT COLOR=#" & Hx5 & ">" & n$ & "<FONT COLOR=#" & Hx4 & ">" & Q$ & "<FONT COLOR=#" & Hx3 & ">" & V$ & "<FONT COLOR=#" & Hx2 & ">" & Z$
Next W
SendIt (PC$)

End Sub
Sub SignOff()
 DORFAC = FindWindow("AOL Frame25", vbNullString)
SendKeys "%ss"
End Sub
Sub LocateMemba(theSN As String)
Call keyword("aol://3548:" & theSN)
End Sub
Sub FadeBlue(thetext As String)
a = Len(thetext)
For W = 1 To a Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000019>" & ab$ & "<FONT COLOR=#000026>" & u$ & "<FONT COLOR=#00003F>" & S$ & "<FONT COLOR=#000058>" & T$ & "<FONT COLOR=#000072>" & Y$ & "<FONT COLOR=#00008B>" & l$ & "<FONT COLOR=#0000A5>" & F$ & "<FONT COLOR=#0000BE>" & B$ & "<FONT COLOR=#0000D7>" & c$ & "<FONT COLOR=#0000F1>" & D$ & "<FONT COLOR=#0000D7>" & H$ & "<FONT COLOR=#0000BE>" & j$ & "<FONT COLOR=#0000A5>" & k$ & "<FONT COLOR=#00008B>" & m$ & "<FONT COLOR=#000072>" & n$ & "<FONT COLOR=#000058>" & Q$ & "<FONT COLOR=#00003F>" & V$ & "<FONT COLOR=#000026>" & Z$
Next W
SendIt (PC$)

End Sub
Sub SendIt(chat As String)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub FadeRed(thetext As String)
a = Len(thetext)
For W = 1 To a Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FF0000>" & ab$ & "<FONT COLOR=#990000>" & u$ & "<FONT COLOR=#880000>" & S$ & "<FONT COLOR=#770000>" & T$ & "<FONT COLOR=#660000>" & Y$ & "<FONT COLOR=#550000>" & l$ & "<FONT COLOR=#440000>" & F$ & "<FONT COLOR=#330000>" & B$ & "<FONT COLOR=#220000>" & c$ & "<FONT COLOR=#110000>" & D$ & "<FONT COLOR=#220000>" & H$ & "<FONT COLOR=#330000>" & j$ & "<FONT COLOR=#440000>" & k$ & "<FONT COLOR=#550000>" & m$ & "<FONT COLOR=#660000>" & n$ & "<FONT COLOR=#770000>" & Q$ & "<FONT COLOR=#880000>" & V$ & "<FONT COLOR=#990000>" & Z$
Next W
SendIt (PC$)


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
Sub StopWav()
Call Playwav(" ")
End Sub
Sub RunMenuByStringAOL(stringer As String)
Call RunMenuByString(AOLsWindow(), stringer)
End Sub
Function AOLsWindow()
AOLWindow = FindWindow("AOL Frame25", vbNullString)
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
Sub SendCharNum(win, chars)
e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub
Function GetChatText()
'Room% = FindChatRoom
'AORich% = FindChildByClass(Room%, "RICHCNTL")
'chattext = GetWindow(AORich%)
'GetChatText = chattext
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function
Function Enter(win)
'This will simply press enter
Call SendCharNum(win, 13)
End Function
Sub ShowAOL()
X = FindWindow("AOL Frame25", 0&)
ShoWindow (X)
End Sub
Sub HideWindow(hwnd)
X = ShowWindow(hwnd, SW_HIDE)
End Sub
Sub HideAOL()
X = FindWindow("AOL Frame25", 0&)
HideWindow (X)
End Sub
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
Function GetWinText(hwnd As Integer) As String
TextLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
Buffer$ = Space$(TextLength)
GetTheText = SendMessageByString(hwnd, WM_GETTEXT, LengthOfText + 1, Buffer$)
GetWinText = Buffer$
End Function
Sub keyword(keyword As String)
tool% = FindChildByClass(AOLsWindow(), "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Tool2%, "_AOL_Icon")
For GetIcon = 1 To 20
icon% = GetWindow(icon%, 2)
Next GetIcon
Call timeout(0.05)
Call IconClick(icon%)
Do: DoEvents
mdi% = FindChildByClass(AOLsWindow(), "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
Edit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
icon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And Edit% <> 0 And icon2% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, keyword)
Call timeout(0.05)
Call IconClick(icon2%)
Call IconClick(icon2%)
End Sub
Sub ShoWindow(hwnd)
X = ShowWindow(hwnd, SW_SHOW)
End Sub
Sub FadeBlack(thetext As String)
a = Len(thetext)
For W = 1 To a Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#000000>" & ab$ & "<FONT COLOR=#111111>" & u$ & "<FONT COLOR=#222222>" & S$ & "<FONT COLOR=#333333>" & T$ & "<FONT COLOR=#444444>" & Y$ & "<FONT COLOR=#555555>" & l$ & "<FONT COLOR=#666666>" & F$ & "<FONT COLOR=#777777>" & B$ & "<FONT COLOR=#888888>" & c$ & "<FONT COLOR=#999999>" & D$ & "<FONT COLOR=#888888>" & H$ & "<FONT COLOR=#777777>" & j$ & "<FONT COLOR=#666666>" & k$ & "<FONT COLOR=#555555>" & m$ & "<FONT COLOR=#444444>" & n$ & "<FONT COLOR=#333333>" & Q$ & "<FONT COLOR=#222222>" & V$ & "<FONT COLOR=#111111>" & Z$
Next W
SendIt (PC$)

End Sub
Sub FadeGreen(thetext As String)
a = Len(thetext)
For W = 1 To a Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#001100>" & ab$ & "<FONT COLOR=#002200>" & u$ & "<FONT COLOR=#003300>" & S$ & "<FONT COLOR=#004400>" & T$ & "<FONT COLOR=#005500>" & Y$ & "<FONT COLOR=#006600>" & l$ & "<FONT COLOR=#007700>" & F$ & "<FONT COLOR=#008800>" & B$ & "<FONT COLOR=#009900>" & c$ & "<FONT COLOR=#00FF00>" & D$ & "<FONT COLOR=#009900>" & H$ & "<FONT COLOR=#008800>" & j$ & "<FONT COLOR=#007700>" & k$ & "<FONT COLOR=#006600>" & m$ & "<FONT COLOR=#005500>" & n$ & "<FONT COLOR=#004400>" & Q$ & "<FONT COLOR=#003300>" & V$ & "<FONT COLOR=#002200>" & Z$
Next W
SendIt (PC$)
End Sub
Sub FadeYellow(thetext As String)
a = Len(thetext)
For W = 1 To a Step 18
    ab$ = Mid$(thetext, W, 1)
    u$ = Mid$(thetext, W + 1, 1)
    S$ = Mid$(thetext, W + 2, 1)
    T$ = Mid$(thetext, W + 3, 1)
    Y$ = Mid$(thetext, W + 4, 1)
    l$ = Mid$(thetext, W + 5, 1)
    F$ = Mid$(thetext, W + 6, 1)
    B$ = Mid$(thetext, W + 7, 1)
    c$ = Mid$(thetext, W + 8, 1)
    D$ = Mid$(thetext, W + 9, 1)
    H$ = Mid$(thetext, W + 10, 1)
    j$ = Mid$(thetext, W + 11, 1)
    k$ = Mid$(thetext, W + 12, 1)
    m$ = Mid$(thetext, W + 13, 1)
    n$ = Mid$(thetext, W + 14, 1)
    Q$ = Mid$(thetext, W + 15, 1)
    V$ = Mid$(thetext, W + 16, 1)
    Z$ = Mid$(thetext, W + 17, 1)
    PC$ = PC$ & "<b><FONT COLOR=#FFFF00>" & ab$ & "<FONT COLOR=#999900>" & u$ & "<FONT COLOR=#888800>" & S$ & "<FONT COLOR=#777700>" & T$ & "<FONT COLOR=#666600>" & Y$ & "<FONT COLOR=#555500>" & l$ & "<FONT COLOR=#444400>" & F$ & "<FONT COLOR=#333300>" & B$ & "<FONT COLOR=#222200>" & c$ & "<FONT COLOR=#111100>" & D$ & "<FONT COLOR=#222200>" & H$ & "<FONT COLOR=#333300>" & j$ & "<FONT COLOR=#444400>" & k$ & "<FONT COLOR=#555500>" & m$ & "<FONT COLOR=#666600>" & n$ & "<FONT COLOR=#777700>" & Q$ & "<FONT COLOR=#888800>" & V$ & "<FONT COLOR=#999900>" & Z$
Next W
SendIt (PC$)

End Sub
Sub SetText(win, thetxt)
thetxxt% = SendMessageByString(win, WM_SETTEXT, 0, thetxt)
End Sub
Sub Playwav(File)
WAVName$ = File
daFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(WAVName$, daFlags%)
End Sub
Sub SendMail(sn, subject, message)
tool% = FindChildByClass(AOLsWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
icon% = GetWindow(icon%, GW_HWNDNEXT)
Call IconClick(icon%)
Do: DoEvents
mail% = FindChildByTitle(AOLMDI(), "Write Mail")
Edit% = FindChildByClass(mail%, "_AOL_Edit")
rich% = FindChildByClass(mail%, "RICHCNTL")
icon% = FindChildByClass(mail%, "_AOL_ICON")
Loop Until mail% <> 0 And Edit% <> 0 And rich% <> 0 And icon% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, sn)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Call SendMessageByString(Edit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(rich%, WM_SETTEXT, 0, message)
For GetIcon = 1 To 18
icon% = GetWindow(icon%, GW_HWNDNEXT)
Next GetIcon
Call IconClick(icon%)
End Sub
Sub timeout(duration)
starttime = Timer
Do While Timer - starttime < duration
DoEvents
Loop
End Sub
