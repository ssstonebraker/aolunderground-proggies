Global RoomBust
Global PromptValue
Global DlogOK
Global DlogStop
Global MailStat
Global MChatBot
Global IgnoreBot
Global ServerBot
Global mComm$
Option Compare Text
Global Const WM_USER = &H400
Global Const EW_ExitWindows = &H0
Global Const SW_HIDE = 0
Global Const SW_NORMAL = 1
Global Const SW_MINIMIZED = 2
Global Const SW_MAXIMIZED = 3
Global Const SW_NOACTIVATE = 4
Global Const SW_SHOW = 5
Global Const SW_MINIMIZE = 6
Global Const SW_MINNOACTIVE = 7
Global Const SW_NA = 8
Global Const SW_RESTORE = 9
Global Const BM_SETCHECK = WM_USER + 1
Global Const LB_SETCURSEL = (WM_USER + 7)
Global Const LB_GETTEXT = (WM_USER + 10)
Global Const LB_GETTEXTLEN = (WM_USER + 11)
Type Rect
  Left As Integer
  Top As Integer
  Right As Integer
  Bottom As Integer
End Type
Type HelpWinInfo
  wStructSize As Integer
  x As Integer
  y As Integer
  dx As Integer
  dy As Integer
  wMax As Integer
  rgChMember As String * 2
End Type

'                   API Subs and Functions
'                   ----------------------

'Subs and Functions for "User"
Declare Sub ReleaseCapture Lib "User" ()
Declare Sub closewindow Lib "User" ()
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, LPRect As Rect)
Declare Sub SetWindowPos Lib "User" (ByVal h%, ByVal hb%, ByVal x%, ByVal y%, ByVal cx%, ByVal cy%, ByVal f%)
Declare Sub DrawMenuBar Lib "User" (ByVal hWnd As Integer)
Declare Sub GetScrollRange Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer, Lpminpos As Integer, lpmaxpos As Integer)
Declare Sub SetCursorPos Lib "User" (ByVal x As Integer, ByVal y As Integer)
Declare Sub UpdateWindow Lib "User" (ByVal hWnd%)
Declare Sub ShowOwnedPopups Lib "User" (ByVal hWnd%, ByVal fShow%)
Declare Function ExitWindows Lib "User" (ByVal dwReturnCode As Long, ByVal wReserved As Integer) As Integer
Declare Function showwindow Lib "User" (ByVal hWnd As Integer, ByVal nCmdShow As Integer) As Integer
Declare Function getfocus Lib "User" () As Integer
Declare Function getnextwindow Lib "User" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource%) As Integer
Declare Function GetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%) As Integer
Declare Function SetClassWord Lib "User" (ByVal hWnd%, ByVal nIndex%, ByVal wNewWord%) As Integer
Declare Function setfocusapi% Lib "User" Alias "SetFocus" (ByVal hWnd As Integer)
Declare Function getwindow% Lib "User" (ByVal hWnd%, ByVal wCmd%)
Declare Function findwindow% Lib "User" (ByVal lpClassName As Any, ByVal lpWindowName As Any)
Declare Function FindWindowByNum% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function FindWindowByString% Lib "User" Alias "FindWindow" (ByVal lpClassName&, ByVal lpWindowName&)
Declare Function ExitWindow% Lib "User" (ByVal dwReturnCode&, ByVal wReserved%)
Declare Function getparent% Lib "User" (ByVal hWnd As Integer)
Declare Function SetParent% Lib "User" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer)
Declare Function GetMessage% Lib "User" (lpMsg As String, ByVal hWnd As Integer, ByVal wMsgFilterMin As Integer, ByVal wMsgFilterMax As Integer)
Declare Function GetMenuString% Lib "User" (ByVal hMenu%, ByVal wIDItem%, ByVal lpString$, ByVal nMaxCount%, ByVal wFlag%)
Declare Function sendmessage Lib "User" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lparam As Any) As Long
Declare Function sendmessagebystring& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lparam$)
Declare Function sendmessagebynum& Lib "User" Alias "SendMessage" (ByVal hWnd%, ByVal wMsg%, ByVal wParam%, ByVal lparam&)
Declare Function CreateMenu% Lib "User" ()
Declare Function AppendMenu Lib "User" (ByVal hMenu As Integer, ByVal wflags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any) As Integer
Declare Function AppendMenuByString% Lib "User" Alias "AppendMenu" (ByVal hMenu%, ByVal wFlag%, ByVal wIDNewItem%, ByVal lpNewItem$)
Declare Function InsertMenu% Lib "User" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wflags As Integer, ByVal wIDNewItem As Integer, ByVal lpNewItem As Any)
Declare Function WinHelp% Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As Any)
Declare Function WinHelpByString% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$)
Declare Function WinHelpByNum% Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&)
Declare Function getwindow% Lib "User" (ByVal hWnd%, ByVal wCmd%)
Declare Function getwindowtext% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String, ByVal aint As Integer)
Declare Function GetWindowWord Lib "User" (ByVal hWnd As Integer, ByVal nIndex As Integer) As Integer
Declare Function setwindowtext% Lib "User" (ByVal hWnd As Integer, ByVal lpString As String)
Declare Function GetActiveWindow% Lib "User" ()
Declare Function setactivewindow% Lib "User" (ByVal hWnd%)
Declare Function GetSysModalWindow% Lib "User" ()
Declare Function SetSysModalWindow% Lib "User" (ByVal hWnd As Integer)
Declare Function iswindowvisible% Lib "User" (ByVal hWnd%)
Declare Function getcurrenttime& Lib "User" ()
Declare Function GetScrollPos Lib "User" (ByVal hWnd As Integer, ByVal nBar As Integer) As Integer
Declare Function getcursor% Lib "User" ()
Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
Declare Function GetSystemMenu Lib "User" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Declare Function GetNextDlgTabItem Lib "User" (ByVal hDlg As Integer, ByVal hctl As Integer, ByVal bPrevious As Integer) As Integer
Declare Function GetWindowTextLength Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function gettopwindow Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ArrangeIconicWindow% Lib "User" (ByVal hWnd%)
Declare Function GetMenu% Lib "User" (ByVal hWnd%)
Declare Function GetMenuItemID% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetMenuItemCount% Lib "User" (ByVal hMenu%)
Declare Function GetMenuState% Lib "User" (ByVal hMenu%, ByVal wId%, ByVal wflags%)
Declare Function GetSubMenu% Lib "User" (ByVal hMenu%, ByVal nPos%)
Declare Function GetSystemMetrics Lib "User" (ByVal nIndex%) As Integer
Declare Function GetDesktopWindow Lib "User" () As Integer
Declare Function GetDC Lib "User" (ByVal hWnd%) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd%, ByVal hdc%) As Integer
Declare Function GetWindowDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function SwapMouseButton% Lib "User" (ByVal bSwap%)
Declare Function ENumChildWindow% Lib "User" (ByVal hwndparent%, ByVal lpenumfunc&, ByVal lparam&)

'Subs and Functions for "Kernel"
Declare Function lStrlenAPI Lib "Kernel" Alias "lStrln" (ByVal lp As Long) As Integer
Declare Function GetWindowDirectory% Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%)
Declare Function GetWinFlags Lib "Kernel" () As Long
Declare Function GetVersion Lib "Kernel" () As Long
Declare Function GetFreeSpace Lib "Kernel" (ByVal wflags%) As Long
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer$, ByVal nSize%) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal NewString As String, ByVal filename As String) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal AppName As String, ByVal KeyName As String, ByVal default As String, ByVal ReturnedString As String, ByVal MAXSIZE As Integer, ByVal filename As String) As Integer
Declare Function GetProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%) As Integer
Declare Function GetProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpDefaulT, ByVal lpReturnedString$, ByVal nSize%) As Integer
Declare Function WriteProfileString Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal lpString$) As Integer
Declare Function GetPrivateProfileInt Lib "Kernel" (ByVal lpAppName$, ByVal lpKeyName$, ByVal nDefault%, ByVal lpFileName$) As Integer
Declare Function agGetStringFromLPSTR$ Lib "APIGuide.Dll" (ByVal lpString&)

'Subs and Functions for "MMSystem"
Declare Function sndplaysound Lib "MMSystem" (ByVal lpWavName$, ByVal Flags%) As Integer '

'Subs and Functions for "VBWFind.Dll"
Declare Function FindChild% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal title$)
Declare Function findchildbytitle% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal title$)
Declare Function findchildbyclass% Lib "VBWFind.Dll" (ByVal hWnd%, ByVal title$)

'Subs and Functions for Other DLL's and VBX's
Declare Function AOLGetList% Lib "green.dll" (ByVal Index%, ByVal Buf$)
Declare Function VarPtr& Lib "VBRun300.Dll" (Param As Any)

'                   Sound Constants
'                   ---------------

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

'                   Global Constants
'                   ----------------
'Window Messages
Global Const WM_NULL = &H0
Global Const WM_CREATE = &H1
Global Const WM_DESTROY = &H2
Global Const WM_MOVE = &H3
Global Const WM_SIZE = &H5
Global Const WM_ACTIVATE = &H6
Global Const WM_SETFOCUS = &H7
Global Const WM_KILLFOCUS = &H8
Global Const WM_ENABLE = &HA
Global Const WM_SETREDRAW = &HB
Global Const WM_SETTEXT = &HC
Global Const WM_GETTEXT = &HD
Global Const WM_GETTEXTLENGTH = &HE
Global Const WM_PAINT = &HF
Global Const WM_CLOSE = &H10
Global Const WM_QUERYENDSESSION = &H11
Global Const WM_QUIT = &H12
Global Const WM_QUERYOPEN = &H13
Global Const WM_ERASEBKGND = &H14
Global Const WM_SYSCOLORCHANGE = &H15
Global Const WM_ENDSESSION = &H16
Global Const WM_SYSTEMERROR = &H17
Global Const WM_SHOWWINDOW = &H18
Global Const WM_CTLCOLOR = &H19
Global Const WM_WININICHANGE = &H1A
Global Const WM_DEVMODECHANGE = &H1B
Global Const WM_ACTIVATEAPP = &H1C
Global Const WM_FONTCHANGE = &H1D
Global Const WM_TIMECHANGE = &H1E
Global Const WM_CANCELMODE = &H1F
Global Const WM_SETCURSOR = &H20
Global Const WM_MOUSEACTIVATE = &H21
Global Const WM_CHILDACTIVATE = &H22
Global Const WM_QUEUESYNC = &H23
Global Const WM_GETMINMAXINFO = &H24
Global Const WM_PAINTICON = &H26
Global Const WM_ICONERASEBKGND = &H27
Global Const WM_NEXTDLGCTL = &H28
Global Const WM_SPOOLERSTATUS = &H2A
Global Const WM_DRAWITEM = &H2B
Global Const WM_MEASUREITEM = &H2C
Global Const WM_DELETEITEM = &H2D
Global Const WM_VKEYTOITEM = &H2E
Global Const WM_CHARTOITEM = &H2F
Global Const WM_SETFONT = &H30
Global Const WM_GETFONT = &H31
Global Const WM_COMMNOTIFY = &H44
Global Const WM_QUERYDRAGICON = &H37
Global Const WM_COMPAREITEM = &H39
Global Const WM_COMPACTING = &H41
Global Const WM_WINDOWPOSCHANGING = &H46
Global Const WM_WINDOWPOSCHANGED = &H47
Global Const WM_POWER = &H48
Global Const WM_NCCREATE = &H81
Global Const WM_NCDESTROY = &H82
Global Const WM_NCCALCSIZE = &H83
Global Const WM_NCHITTEST = &H84
Global Const WM_NCPAINT = &H85
Global Const WM_NCACTIVATE = &H86
Global Const WM_GETDLGCODE = &H87
Global Const WM_NCMOUSEMOVE = &HA0
Global Const WM_NCLBUTTONDOWN = &HA1
Global Const WM_NCLBUTTONUP = &HA2
Global Const WM_NCLBUTTONDBLCLK = &HA3
Global Const WM_NCRBUTTONDOWN = &HA4
Global Const WM_NCRBUTTONUP = &HA5
Global Const WM_NCRBUTTONDBLCLK = &HA6
Global Const WM_NCMBUTTONDOWN = &HA7
Global Const WM_NCMBUTTONUP = &HA8
Global Const WM_NCMBUTTONDBLCLK = &HA9
Global Const WM_KEYFIRST = &H100
Global Const WM_KEYDOWN = &H100
Global Const WM_KEYUP = &H101
Global Const WM_CHAR = &H102
Global Const WM_DEADCHAR = &H103
Global Const WM_SYSKEYDOWN = &H104
Global Const WM_SYSKEYUP = &H105
Global Const WM_SYSCHAR = &H106
Global Const WM_SYSDEADCHAR = &H107
Global Const WM_KEYLAST = &H108
Global Const WM_INITDIALOG = &H110
Global Const WM_COMMAND = &H111
Global Const WM_SYSCOMMAND = &H112
Global Const WM_TIMER = &H113
Global Const WM_HSCROLL = &H114
Global Const WM_VSCROLL = &H115
Global Const WM_INITMENU = &H116
Global Const WM_INITMENUPOPUP = &H117
Global Const WM_MENUSELECT = &H11F
Global Const WM_MENUCHAR = &H120
Global Const WM_ENTERIDLE = &H121
Global Const WM_MOUSEFIRST = &H200
Global Const WM_MOUSEMOVE = &H200
Global Const WM_LBUTTONDOWN = &H201
Global Const WM_LBUTTONUP = &H202
Global Const WM_LBUTTONDBLCLK = &H203
Global Const WM_RBUTTONDOWN = &H204
Global Const WM_RBUTTONUP = &H205
Global Const WM_RBUTTONDBLCLK = &H206
Global Const WM_MBUTTONDOWN = &H207
Global Const WM_MBUTTONUP = &H208
Global Const WM_MBUTTONDBLCLK = &H209
Global Const WM_MOUSELAST = &H209
Global Const WM_PARENTNOTIFY = &H210
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
Global Const WM_DROPFILES = &H233
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
Global Const WM_CLEAR = &H303
Global Const WM_UNDO = &H304
Global Const WM_RENDERFORMAT = &H305
Global Const WM_RENDERALLFORMATS = &H306
Global Const WM_DESTROYCLIPBOARD = &H307
Global Const WM_DRAWCLIPBOARD = &H308
Global Const WM_PAINTCLIPBOARD = &H309
Global Const WM_VSCROLLCLIPBOARD = &H30A
Global Const WM_SIZECLIPBOARD = &H30B
Global Const WM_ASKCBFORMATNAME = &H30C
Global Const WM_CHANGECBCHAIN = &H30D
Global Const WM_HSCROLLCLIPBOARD = &H30E
Global Const WM_QUERYNEWPALETTE = &H30F
Global Const WM_PALETTEISCHANGING = &H310
Global Const WM_PALETTECHANGED = &H311

'Menu flags for Add/Check/EnableMenuItem()
Global Const MF_INSERT = &H0
Global Const MF_CHANGE = &H80
Global Const MF_APPEND = &H100
Global Const MF_DELETE = &H200
Global Const MF_REMOVE = &H1000
Global Const MF_BYCOMMAND = &H0
Global Const MF_BYPOSITION = &H400
Global Const MF_SEPARATOR = &H800
Global Const MF_ENABLED = &H0
Global Const MF_GRAYED = &H1
Global Const MF_DISABLED = &H2
Global Const MF_UNCHECKED = &H0
Global Const MF_CHECKED = &H8
Global Const MF_USECHECKBITMAPS = &H200
Global Const MF_STRING = &H0
Global Const MF_BITMAP = &H4
Global Const MF_OWNERDRAW = &H100
Global Const MF_POPUP = &H10
Global Const MF_MENUBARBREAK = &H20
Global Const MF_MENUBREAK = &H40
Global Const MF_UNHILITE = &H0
Global Const MF_HILITE = &H80
Global Const MF_SYSMENU = &H2000
Global Const MF_HELP = &H4000
Global Const MF_MOUSESELECT = &H8000
Global Const MF_END = &H80

'SetWindowPos Flags
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOZORDER = &H4
Global Const SWP_NOREDRAW = &H8
Global Const SWP_NOACTIVATE = &H10
Global Const SWP_DRAWFRAME = &H20
Global Const SWP_SHOWWINDOW = &H40
Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_NOCOPYBITS = &H100
Global Const SWP_NOREPOSITION = &H200
Global Const SWP_NOSENDCHANGING = &H400
Global Const SWP_DEFERERASE = &H2000

'GetWindow() Constants
Global Const GW_HWNDFIRST = 0
Global Const GW_HWNDLAST = 1
Global Const GW_HWNDNEXT = 2
Global Const GW_HWNDPREV = 3
Global Const GW_OWNER = 4
Global Const GW_CHILD = 5

'One More
Global Const LB_GETCOUNT = (WM_USER + 12)

Sub addroom (lst As ListBox)
For Index% = 0 To 22
    DoEvents
    StringSpace$ = String(255, 0)
    g% = AOLGetList(Index%, StringSpace$)
    StringSpace$ = Left(StringSpace$, InStr(1, StringSpace$, Chr$(0)) - 1)
    For z% = 0 To lst.ListCount - 1
        DoEvents
        If lst.List(z%) = StringSpace$ Then DontAdd = True
    Next z%
    If StringSpace$ = "" Then Exit Sub
    If DontAdd = False Then lst.AddItem (StringSpace$)
Next Index%
End Sub

Sub aokeyword (keyword As String)
AOL% = findwindow("AOL Frame25", 0&)
Call runmenubystring("Keyword...", "&Go To")
waitforchild "Keyword", True
x% = findchildbytitle(AOL%, "Keyword")
e% = findchildbyclass(x%, "_AOL_EDIT")
k% = sendmessagebystring(e%, WM_SETTEXT, o, keyword)
g% = getnextwindow(e%, 2)
aolclick (g%)
End Sub

Sub aolclick (E1 As Integer)
timeout (.1)
g% = sendmessagebynum(E1, WM_LBUTTONDOWN, 0, 0&)
g% = sendmessagebynum(E1, WM_LBUTTONUP, 0, 0&)
timeout (.1)
End Sub

Sub center (Frm As Form)
x% = (Screen.Width - Frm.Width) / 2
y% = (Screen.Height - Frm.Height) / 2
Frm.Move x%, y%
End Sub

Sub dlog (Werd)
DlogOK = False
Message.Show
Message.Label1.Caption = Werd
Do
    DoEvents
Loop Until DlogOK
End Sub

Function elroy () As String
elroy = Chr$(69) & Chr$(108) & Chr$(114) & Chr$(111) & Chr$(121)
End Function

Function getsn () As String
On Error Resume Next
AOL = findwindow("aol frame25", 0&)
Wel = findchildbytitle(AOL, "Welcome,")
If Wel = 0 Then getsn = "Not Online": Exit Function
namelen = sendmessage(Wel, WM_GETTEXTLENGTH, 0, 0)
buffer$ = String$(namelen, 0)
x = sendmessagebystring(Wel, WM_GETTEXT, namelen, buffer$)
a = InStr(buffer$, ",")
sn$ = Mid$(buffer$, a + 2, (Len(buffer$) - (a + 1)))
sn$ = Trimnull(sn$)
getsn = sn$
End Function

Sub logline (l As String)
If MenuForm.itemLog.Checked = False Then Exit Sub
f% = FreeFile
Open "c:\green.log" For Binary Access Write As f%
p$ = l & Chr$(13) & Chr$(10)
Put #1, LOF(1) + 1, p$
Close f%
End Sub

Sub mailpref ()
AOL% = findwindow("AOL Frame25", 0&)
ssa% = whichversion()
If ssa% = 25 Then wh$ = "Set Preferences" + Chr$(9) + "Ctrl+="
If ssa% = 3 Then wh$ = "Preferences"
Call runmenubystring(wh$, "Mem&bers")
Do
    DoEvents
    c% = findchildbytitle(AOL%, "Preferences")
Loop Until c%
d% = getnextwindow(findchildbytitle(c%, "Mail"), 2)
Do
    DoEvents
    aolclick (d%)
    e% = findwindow("_AOL_MODAL", "Mail Preferences")
Loop Until e%
f% = findchildbytitle(e%, "Confirm mail after it has been sent")
h% = sendmessagebynum(f%, BM_SETCHECK, False, 0)
x% = findchildbytitle(e%, "Close mail after it has been sent")
y% = sendmessagebynum(x%, BM_SETCHECK, True, 0)
j% = findchildbytitle(e%, "OK")
Do
    DoEvents
    aolclick (j%)
    e% = findwindow("_AOL_MODAL", "Mail Preferences")
Loop While e%
k% = sendmessagebynum(c%, WM_CLOSE, 0, 0&)
End Sub

Sub plawav (file)
soundname$ = file
wflags% = SND_ASYNC Or SND_NODEFAULT
x% = sndplaysound(soundname$, wflags%)
End Sub

Function ReadINI (AppName, KeyName, filename As String) As String
Dim sRet As String
    sRet = String(255, 0)
    ReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Sub runmenubystring (Menu_String As String, Top_Position As String)
Top_Position_Num = -1
AOL% = findwindow("AOL Frame25", 0&)
Menu_Handle = GetMenu(AOL%)
Do
    DoEvents
    Top_Position_Num = Top_Position_Num + 1
    buffer$ = String$(255, 0)
    Look_For_Menu_String% = GetMenuString(Menu_Handle, Top_Position_Num, buffer$, Len(Top_Position) + 1, MF_BYPOSITION)
    Trim_Buffer = TrimNull2(buffer$)
Loop Until Trim_Buffer = Top_Position
Sub_Menu_Handle = GetSubMenu(Menu_Handle, Top_Position_Num)
BY_POSITION = -1
Do
    DoEvents
    BY_POSITION = BY_POSITION + 1
    buffer$ = String(255, 0)
    Look_For_Menu_String% = GetMenuString(Sub_Menu_Handle, BY_POSITION, buffer$, Len(Menu_String) + 1, MF_BYPOSITION)
    Trim_Buffer = TrimNull2(buffer$)
Loop Until Trim_Buffer = Menu_String
DoEvents
Get_ID% = GetMenuItemID(Sub_Menu_Handle, BY_POSITION)
Click_Menu_Item = sendmessagebynum(AOL%, WM_COMMAND, Get_ID%, 0)
End Sub

Sub send (chatstring$)
AOL% = findwindow("AOL Frame25", 0&)
List% = findchildbyclass(AOL%, "_AOL_LISTBOX")
room = getparent(List%)
TALK% = findchildbyclass(room, "_AOL_EDIT")
send2% = getwindow(TALK%, 2)
z = sendmessagebynum(send2%, WM_LBUTTONUP, 0, 0)
x = sendmessagebystring(TALK%, WM_SETTEXT, 0, chatstring$)
z = sendmessagebynum(send2%, WM_LBUTTONDOWN, 0, 0)
z = sendmessagebynum(send2%, WM_LBUTTONUP, 0, 0)
If MChatBot = True Then g% = setfocusapi(Chat.hWnd)
timeout (.02)
End Sub

Sub sendim (sn$, msg$)
AOL% = findwindow("AOL Frame25", 0&)
v% = findchildbytitle(AOL%, "Send Instant Message")
If v% = 0 Then
    Call runmenubystring("Send an Instant Message", "Mem&bers")
    Do
        DoEvents
        v% = findchildbytitle(AOL%, "Send Instant Message")
    Loop Until v%
End If
g% = setfocusapi(v%)
imsn% = getnextwindow(findchildbytitle(v%, "To:"), 2)
g% = sendmessagebystring(imsn%, WM_SETTEXT, 0, sn$)
Select Case whichversion()
    Case 25
        t% = getnextwindow(imsn%, 2)
    Case Else
        t% = findchildbyclass(v%, "RICHCNTL")
End Select
g% = sendmessagebystring(t%, WM_SETTEXT, 0, "")
g% = sendmessagebystring(t%, WM_SETTEXT, 0, msg$)
w% = getnextwindow(t%, 2)
aolclick (w%)
g% = sendmessagebynum(v%, WM_CLOSE, 0, 0)
End Sub

Function sg () As String
sg = Chr$(83) & Chr$(111) & Chr$(121) & Chr$(108) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(32) & Chr$(71) & Chr$(114) & Chr$(101) & Chr$(101) & Chr$(110) & Chr$(32) & Chr$(87) & Chr$(97) & Chr$(114) & Chr$(101) & Chr$(122) & Chr$(32) & Chr$(83) & Chr$(101) & Chr$(114) & Chr$(118) & Chr$(101) & Chr$(114)
End Function

Function sg2 () As String
sg2 = Chr$(83) & Chr$(111) & Chr$(121) & Chr$(108) & Chr$(101) & Chr$(110) & Chr$(116) & Chr$(32) & Chr$(71) & Chr$(114) & Chr$(101) & Chr$(101) & Chr$(110)
End Function

Sub stayontop (f As Form)
SetWindowPos f.hWnd, -1, 0, 0, 0, 0, &H50
End Sub

Sub timeout (Duration)
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop
End Sub

Function Trimnull (in) As String
Dim x, total$
For x = 1 To Len(in)
    If (Mid$(in, x, 1) <> Chr$(0)) Then
        total$ = total$ + Mid$(in, x, 1)
    Else
        GoTo NullDetect
    End If
Next
NullDetect:
Trimnull = total$
End Function

Function TrimNull2 (in$) As String
For x = 1 To Len(in$)
    If (Mid$(in$, x, 1) <> Chr$(0)) Then
        total$ = total$ + Mid$(in$, x, 1)
    Else
        GoTo NullDetect2
    End If
Next
NullDetect2:
TrimNull2 = total$
End Function

Function vers ()
vers = Chr$(70) & Chr$(105) & Chr$(110) & Chr$(97) & Chr$(108)
End Function

Sub waitforchild (childname$, tf%)
AOL% = findwindow("AOL Frame25", 0&)
WaitForChildLoop:
DoEvents
x% = findchildbytitle(AOL%, childname$)
If tf% <> False Then
    If x% = 0 Then GoTo WaitForChildLoop
Else
    If x% <> 0 Then GoTo WaitForChildLoop
End If
End Sub

Function whichversion ()
AOL% = findwindow("AOL Frame25", 0&)
Wel% = findchildbytitle(AOL%, "Welcome")
aol3% = findchildbyclass(Wel%, "RICHCNTL")
If aol3% = 0 Then whichversion = 25: Exit Function
If aol3% <> 0 Then whichversion = 3: Exit Function
End Function

Sub WriteINI (sAppname, sKeyName, sNewString, sFileName As String)
'Example: WriteINI("DaProggy", "Lamers Name", text3.text, app.path + "\Prog.ini")
Dim r As Integer
    r = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Sub

