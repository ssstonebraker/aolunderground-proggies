Attribute VB_Name = "Jaguar32"
'Jaguar32.Bas Silver Edition *FIXED*
'(For Visual Basic Versions 4, 5, 6)
'For use with Aol (95 and 4.0)
'Release Date: Christmas of '98
'Please note disclaimer at bottom along with Fader notes from monk-e-god
'Do NOT copy bas file under any conditions. See disclaimer.

' *Use Jaguar32 at your own risk. We are not responsible for anything
'made using or while using jaguar32.*

'Creators contact addresses: (if known)
'Jaguar             Jaguar32X@Juno.com (Main Email)
'Jaguar             Jaguar32X@iname.com (Fowarding Service)
'Jaguar (AIM)    PuRRBaLL11 [That has an eleven at the end of it if you cant read or understand that]
'Flux                dr_flux@hotmail.com
'Monk-e-god     monkegod@hotmail.com
'Genghis           LordGenghis@Juno.Com (cant get it to work)
'Dolan               xxdolanxx@hotmail.com
'Beav                phishme7@juno.com
'JMR                 XxJMR@aol.com

'Creators: (9)
'Jaguar
'Flux
'VSTD COORD
'Baron
'Genghis
'monk-e-god
'Beav
'JMR
'Dolan

'Jaguar32 News and Updates:
'Well once again, a new version. Woohoo! I made about 6 new subs or
'functions and we now have faders thanks to monk-e-god who kicks
'ye royal ass. Also, once Dolan gets to it, he is making a server so all
'the lazy people can have servers. I am getting a new computer this
'Christmas and its kicking even more royal ass. I am now opening
'Jaguar32 up to new suggestions to subs and functions I can make.
'Im running out of ideas. Accually, I have enough to update it for about
'3 years, but I thought it would be nice to open it up to your suggestions.
'Plus, you get recognized for your work since people will use it.
'Im also thinking about a programming group over the internet, which
'would be both AOL and non AOL related which if you think its a good
'idea email me. That about wraps it up.

'A note from Jaguar about disclaimer:
'Programmers of the Jaguar32 team have decided that since we make
'and distribute Jaguar32 as freeware and that we do not have a charge
'for Jaguar32 that we strictly enforce the disclaimer. In other words,
'we spent a hell of a lot of time bringing you Jaguar32 and its a aid
'to you in your programming so we ask that you follow the disclaimer.

'Dislcaimer:

' *Use Jaguar32 at your own risk. We are not responsible for anything
'made using or while using jaguar32.*

'Thank you for choosing Jaguar32.Bas. Before you start using this bas
'file there are a couple of things we have all agreed on. First off, we do
'not want anyone to add on to this bas file and call it theirs. Make your
'own fucking bas file from scratch like we did! Second, please do not
'tamper with the code that is in the bas file unless you have emailed
'Jaguar and I have said it was ok. Third, if someone wants Jaguar32
'please tell them to email Jaguar at Jaguar32X@Juno.com. Fourth,
'unless you have the expressed written permission by one of the
'creators, we do not want to see this bas file on servers or mms
'because it is constantly being updated, and we have the latest
'update. Basically cause we have to send the thing out again to
'people that we have sent it to a million times.

'PS. The Addroom (not the one for AOL4) works only on AOL95
'not AOL3.0

'Special Thanks Given To:
'All of Jaguar, Genghis, Baron, VSTD, and Flux's friends (you know who you are)
'All of Teel (including Rj2 who is totally awesome)
'NIKON <~He is way cool
'Puma & Panther & Leopard
'People who dont decompile and steal codes.
'People who prog their ass off and dont get any credit for what they do.
'The holy hand-grenade and the man-eating bunny.

'Some notes on Fader Functions by monk-e-god.
'Some subs in this bas may not be self-explanatory at first because
'they require you to type in the red, green, and blue values of each color.
'Some of you might not know the RGB values of certain colors so here are
'a few:

'Red = R: 255, G: 0, B:0
'Green = R: 0, G: 255, B:0
'Blue = R: 0, G: 0, B: 255
'Yellow = R: 255, G: 255, B: 0
'White = R: 255, G: 255, B: 255
'Black = R: 0, G: 0, B: 0

'So to fade from Blue to Black to Blue you would do:
'FadedText$ = XFader_FadeThreeColor(0, 0, 255, 0, 0, 0, 0, 0, 255, Text2Fade$, False)

'Or you could use the easier subs by doing:
'FadedText$ = XFader_FadeByColor3(FADE_BLUE, FADE_BLACK, FADE_BLUE, Text2Fade$, False)

'To make the text wavy all you have
'to do is set the last parameter(Wavy)
'to True.

'Multifading 101 by Monk-e-god
'To use this you need to declare an array and fill it with the colors to fade.

'Example:
'Dim ColorArray(4)
'ColorArray(1) = FADE_RED
'ColorArray(2) = FADE_BLACK
'ColorArray(3) = FADE_BLUE
'ColorArray(4) = FADE_BLACK
'FadedText$ = MultiFade(4, ColorArray, "The Text You Want To Fade", False)

'Enjoy Jaguar32!
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppname As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
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
Declare Function SendMessage2 Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function mciSendString& Lib "Winmm" Alias "mciSendStringA" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen&, ByVal hCallBack&)

Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_ALIAS = &H10000
Public Const SND_FILENAME = &H20000
Public Const SND_RESOURCE = &H40004
Public Const SND_ALIAS_ID = &H110000
Public Const SND_ALIAS_START = 0
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const SND_VALID = &H1F
Public Const SND_NOWAIT = &H2000
Public Const SND_VALIDFLAGS = &H17201F

Public Const SND_RESERVED = &HFF000000
Public Const SND_TYPE_MASK = &H170007


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
Public Const EM_GETLINE = &HC4

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

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2

Public Const HTERROR = (-2)
Public Const HTTRANSPARENT = (-1)
Public Const HTNOWHERE = 0
Public Const HTCLIENT = 1
Public Const HTCAPTION = 2
Public Const HTSYSMENU = 3
Public Const HTGROWBOX = 4
Public Const HTSIZE = HTGROWBOX
Public Const HTMENU = 5
Public Const HTHSCROLL = 6
Public Const HTVSCROLL = 7
Public Const HTMINBUTTON = 8
Public Const HTMAXBUTTON = 9
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17
Public Const HTBORDER = 18
Public Const HTREDUCE = HTMINBUTTON
Public Const HTZOOM = HTMAXBUTTON
Public Const HTSIZEFIRST = HTLEFT
Public Const HTSIZELAST = HTBOTTOMRIGHT

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
Global Title
Global Buff3
Global buff2
Global Buff
Global ct
Global RoomHits
Global Log
Global TheList
Global r&
Global entry$
Global iniPath$
Global mmlastline

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
  Red As Long
  Green As Long
  Blue As Long
End Type

Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Global fore%
Global username As String
Function AOLActivate()
X = GetCaption(AOLWindow)
AppActivate X
End Function

Public Sub AOLChatSend2(Txt As textbox)
'This scrolls a multilined textbox adding pauses where needed
'This is basically for macro shops and things like that.
AOLChatSend "· ···•(\›• INCOMMING TEXT"
Pause 4
Dim onelinetxt$, X$, Start%, i%
Start% = 1
fa = 1
For i% = Start% To Len(Txt.Text)
X$ = Mid(Txt.Text, i%, 1)
onelinetxt$ = onelinetxt$ + X$
If Asc(X$) = 13 Then
AOLChatSend ": " + onelinetxt$
Pause (0.5)
J% = J% + 1
i% = InStr(Start%, Txt.Text, X$)
If i% >= Len(Txt.Text) Then Exit For
Start% = i% + 1
onelinetxt$ = ""
End If
Next i%
AOLChatSend ":" + onelinetxt$
End Sub


Public Sub AddRoom_SNs(Listboxes As ListBox)
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
Listboxes.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub



Function CutStringDown(ByVal pws As String) As String
On Error Resume Next
CutStringDown = Trim(Left$(pws, InStr(pws, Chr$(0)) - 1))
End Function

Function AOLGotoPrivateRoom(Room As String)
Theroomcode = "aol://2719:2-2-" & Room
AOLKeyword (Theroomcode)
End Function

Function AOLFindIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
AOLFindIM = im%
End Function











Function AOLSupRoom()
AOLIsOnline
If AOLIsOnline = 0 Then GoTo last
AOLFindRoom
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
Call AOLChatSend("~Genghis~ Sup, " & person$)
Pause (0.4)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function




Function FindAOL()
aol% = FindWindow("AOL Frame25", vbNullString)
FindAOL = aol%
End Function

Function FindKeyword()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(mdi%, "Keyword")
Kedit% = FindChildByClass(keyw%, "_AOL_Edit")
FindKeyword = Kedit%
End Function

Function FindNewIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
 End Function

Function FindWelcome()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
FindWelcome = FindChildByTitle(mdi%, "Welcome, ")
End Function

Function AOLIMRoomIMer(Mess As String)
AOLIsOnline
If AOLIsOnline = 0 Then GoTo last


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
Call AOLInstantMessage(person$, Mess)
Next Index
Call CloseHandle(AOLProcessThread)
End If
last:
End Function



Function Mail_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
killwin (A3000%)
End Function

Function Mail_DeleteSent()
Call AOLRunMenuByString("Check Mail You've &Sent")

aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
again:
Pause (1)
A3000% = FindChildByTitle(A2000%, "Outgoing Mail")
If A3000% = 0 Then GoTo again
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
Pause (6)
AOLButton (Delete%)
killwin (A3000%)
End Function

Function Mail_KeepAsNew()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Keepasnew% = FindChildByTitle(A3000%, "Keep As New")
AOLButton (Keepasnew%)
End Function


Function Mail_DeleteSingle()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
AOLButton (Delete%)
End Function


Function Mail_FindComposed()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi = FindChildByClass(aol%, "MDIClient")
Mail_FindComposed = FindChildByTitle(mdi, "Compose Mail")
End Function

Function Mail_ForwardMail(SN As String, Message As String)
FindForwardWindow
person = SN
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Fwd: ")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, person)
a = SendMessageByString(Mess%, WM_SETTEXT, 0, Message)
AOLIcon (icone%)
End Function

Function GetFromINI(AppName$, KeyName$, Filename$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), Filename$))
End Function



Function Mail_ClickForward()
X = FindOpenMail
If X = 0 Then GoTo last
AOLActivate
SendKeys "{TAB}"
AG:
Pause (0.2)
SendKeys " "
X = FindSendWin(2)
If X = 0 Then GoTo AG
last:
End Function

Function Mail_KillComposed()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi = FindChildByClass(aol%, "MDIClient")
Composed = FindChildByTitle(mdi, "Compose Mail")
killwin (Composed)
End Function
Function MAil_BuildList(Lst As ListBox)
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
Pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Lst.AddItem Buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo Start
last:

End Function
Function Mail_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
Pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "New Mail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_CloseMail()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
killwin (A3000%)
End Function

Function Mail_Out_CursorSet(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function






Function Mail_Out_ListMail(Box As ListBox)
Box.Clear
AOLMDI
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then
AOLRunMenuByString ("Read &New Mail")
Justamin:
mailwin = FindChildByTitle(AOLMDI, "New Mail")
If mailwin = 0 Then GoTo Justamin
Pause (7)
End If

mailwin = FindChildByTitle(AOLMDI, "Outgoing FlashMail")
AOLCountMail
Start:
If Counter = AOLCountMail Then GoTo last
MailTree = FindChildByClass(mailwin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, Counter, 0)
    Buffer$ = String$(namelen, 0)
    X = SendMessageByString(MailTree, LB_GETTEXT, Counter, Buffer$)
    TabPos = InStr(Buffer$, Chr$(9))
    Buffer$ = right$(Buffer$, (Len(Buffer$) - (TabPos)))
    Box.AddItem Buffer$
 Pause (0.001)
Counter = Counter + 1
GoTo Start
last:
End Function

Function Mail_Out_MailCaption()
End Function

Function Mail_Out_MailCount()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function Mail_Out_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "Outgoing FlashMail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function


Function Mail_PressEnter()
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
X = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
X = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function



Function Mail_SetCursor(mailIndex As String)
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
A6000% = SendMessageByString(MailTree%, LB_SETCURSEL, mailIndex, 0)
End Function

Function Mail_MailCaption()
FindOpenMail
Mail_MailCaption = GetCaption(FindOpenMail)
End Function


Function ReadFile(Where As String)
filenum = FreeFile
Open (Where) For Input As filenum
info = Input(LOF(filenum), filenum)
info = ReadFile
End Function




Function Text_backwards(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
Text_backwards = newsent$
End Function
Function Text_Elite(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed
If nextchr$ = "A" Then Let nextchr$ = "/\"
If nextchr$ = "a" Then Let nextchr$ = "å"
If nextchr$ = "B" Then Let nextchr$ = "ß"
If nextchr$ = "C" Then Let nextchr$ = "Ç"
If nextchr$ = "c" Then Let nextchr$ = "¢"
If nextchr$ = "D" Then Let nextchr$ = "Ð"
If nextchr$ = "d" Then Let nextchr$ = "ð"
If nextchr$ = "E" Then Let nextchr$ = "Ê"
If nextchr$ = "e" Then Let nextchr$ = "è"
If nextchr$ = "f" Then Let nextchr$ = "ƒ"
If nextchr$ = "H" Then Let nextchr$ = "|-|"
If nextchr$ = "I" Then Let nextchr$ = "‡"
If nextchr$ = "i" Then Let nextchr$ = "î"
If nextchr$ = "k" Then Let nextchr$ = "|‹"
If nextchr$ = "L" Then Let nextchr$ = "£"
If nextchr$ = "M" Then Let nextchr$ = "]V["
If nextchr$ = "m" Then Let nextchr$ = "^^"
If nextchr$ = "N" Then Let nextchr$ = "/\/"
If nextchr$ = "n" Then Let nextchr$ = "ñ"
If nextchr$ = "O" Then Let nextchr$ = "Ø"
If nextchr$ = "o" Then Let nextchr$ = "ö"
If nextchr$ = "P" Then Let nextchr$ = "¶"
If nextchr$ = "p" Then Let nextchr$ = "Þ"
If nextchr$ = "r" Then Let nextchr$ = "®"
If nextchr$ = "S" Then Let nextchr$ = "§"
If nextchr$ = "s" Then Let nextchr$ = "$"
If nextchr$ = "t" Then Let nextchr$ = "†"
If nextchr$ = "U" Then Let nextchr$ = "Ú"
If nextchr$ = "u" Then Let nextchr$ = "µ"
If nextchr$ = "V" Then Let nextchr$ = "\/"
If nextchr$ = "W" Then Let nextchr$ = "VV"
If nextchr$ = "w" Then Let nextchr$ = "vv"
If nextchr$ = "X" Then Let nextchr$ = "X"
If nextchr$ = "x" Then Let nextchr$ = "×"
If nextchr$ = "Y" Then Let nextchr$ = "¥"
If nextchr$ = "y" Then Let nextchr$ = "ý"
If nextchr$ = "!" Then Let nextchr$ = "¡"
If nextchr$ = "?" Then Let nextchr$ = "¿"
If nextchr$ = "." Then Let nextchr$ = "…"
If nextchr$ = "," Then Let nextchr$ = "‚"
If nextchr$ = "1" Then Let nextchr$ = "¹"
If nextchr$ = "%" Then Let nextchr$ = "‰"
If nextchr$ = "2" Then Let nextchr$ = "²"
If nextchr$ = "3" Then Let nextchr$ = "³"
If nextchr$ = "_" Then Let nextchr$ = "¯"
If nextchr$ = "-" Then Let nextchr$ = "—"
If nextchr$ = " " Then Let nextchr$ = " "
If nextchr$ = "<" Then Let nextchr$ = "«"
If nextchr$ = ">" Then Let nextchr$ = "»"
If nextchr$ = "*" Then Let nextchr$ = "¤"
If nextchr$ = "`" Then Let nextchr$ = "“"
If nextchr$ = "'" Then Let nextchr$ = "”"
If nextchr$ = "0" Then Let nextchr$ = "º"
Let newsent$ = newsent$ + nextchr$
Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
Text_Elite = newsent$
End Function
Function Text_Hacker(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "A" Then Let nextchr$ = "a"
If nextchr$ = "E" Then Let nextchr$ = "e"
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "O" Then Let nextchr$ = "o"
If nextchr$ = "U" Then Let nextchr$ = "u"
If nextchr$ = "b" Then Let nextchr$ = "B"
If nextchr$ = "c" Then Let nextchr$ = "C"
If nextchr$ = "d" Then Let nextchr$ = "D"
If nextchr$ = "z" Then Let nextchr$ = "Z"
If nextchr$ = "f" Then Let nextchr$ = "F"
If nextchr$ = "g" Then Let nextchr$ = "G"
If nextchr$ = "h" Then Let nextchr$ = "H"
If nextchr$ = "y" Then Let nextchr$ = "Y"
If nextchr$ = "j" Then Let nextchr$ = "J"
If nextchr$ = "k" Then Let nextchr$ = "K"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = "m" Then Let nextchr$ = "M"
If nextchr$ = "n" Then Let nextchr$ = "N"
If nextchr$ = "x" Then Let nextchr$ = "X"
If nextchr$ = "p" Then Let nextchr$ = "P"
If nextchr$ = "q" Then Let nextchr$ = "Q"
If nextchr$ = "r" Then Let nextchr$ = "R"
If nextchr$ = "s" Then Let nextchr$ = "S"
If nextchr$ = "t" Then Let nextchr$ = "T"
If nextchr$ = "w" Then Let nextchr$ = "W"
If nextchr$ = "v" Then Let nextchr$ = "V"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
Text_Hacker = newsent$
End Function
Function Text_Decode(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
If nextchr$ = "I" Then Let nextchr$ = "i"
If nextchr$ = "l" Then Let nextchr$ = "L"
If nextchr$ = " " Then Let nextchr$ = " "
Let newsent$ = newsent$ + nextchr$
Loop
Text_Decode = newsent$
End Function
Function Text_Spaced(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchr$ = nextchr$ + " "
Let newsent$ = newsent$ + nextchr$
Loop
Text_Spaced = newsent$
End Function
Function Text_Khangolian(Txt As String)
'This translats text into my lang(Khangolian)
Dim Firstletter, LastLetter, Middle
txtlen = Len(Txt)
Firstletter = Left$(Txt, 1)
LastLetter = right$(Txt, 1)
Middle = NotSure
withnofirst = right$(Txt, txtlen - 1)
nofirstlen = Len(withnofirst)
Withnofirstorlast = Left$(withnofirst, nofirstlen - 1)
Text_Encode = LastLetter & Withnofirstorlast & Firstletter
End Function

Function Text_Looping(Txt As String)
Dim thecaption, Captionlen, middlelen, Firstletter, Middle
If Txt = "" Then GoTo dead
thecaption = Txt
Captionlen = Len(thecaption)
middlelen = Captionlen - 1
Firstletter = Left$(thecaption, 1)
Middle = right(thecaption, middlelen)
Text_Looping = Middle & Firstletter
GoTo last
dead:
Text_Looping = ""
last:
End Function

Function Text_StripLetter(Txt As String, Which As String)
'This takes out a certain letter
'Which is the letter you take out(its in number value)
'For example..in the work Khan if I wanted to
'take out the H I would use
'Text_StripLetter("Khan", 2)
txtlen = Len(Txt)
before = Left$(Txt, Which - 1)
MsgBox before
beforelen = Len(before)
afterthat = txtlen - beforelen - 1
After = right$(Txt, afterthat)
MsgBox After
Text_StripLetter = before & After
End Function

Public Sub TextColor_Blue(Txt As textbox)
Txt.ForeColor = &HFFFF00
Pause 0.1
Txt.ForeColor = &HFF0000
Pause 0.1
Txt.ForeColor = &HC00000
Pause 0.1
Txt.ForeColor = &H800000
Pause 0.1
Txt.ForeColor = &H400000
Pause 0.1
End Sub

Public Sub TextColor_Teal(Txt As textbox)
Txt.ForeColor = &HFFFF00
Pause 0.1
Txt.ForeColor = &HC0C000
Pause 0.1
Txt.ForeColor = &H808000
Pause 0.1
Txt.ForeColor = &H404000
Pause 0.1
End Sub

Public Sub TextColor_Green(Txt As textbox)
Txt.ForeColor = &HFF00&
Pause 0.1
Txt.ForeColor = &HC000&
Pause 0.1
Txt.ForeColor = &H8000&
Pause 0.1
Txt.ForeColor = &H4000&
Pause 0.1
End Sub

Public Sub TextColor_Yellow(Txt As textbox)
Txt.ForeColor = &HFFFF&
Pause 0.1
Txt.ForeColor = &HC0C0&
Pause 0.1
Txt.ForeColor = &H8080&
Pause 0.1
Txt.ForeColor = &H4040&
Pause 0.1
End Sub


Public Sub TextColor_Red(Txt As textbox)
Txt.ForeColor = &HFF&
Pause 0.1
Txt.ForeColor = &HC0&
Pause 0.1
Txt.ForeColor = &H80&
Pause 0.1
Txt.ForeColor = &H40&
Pause 0.1
End Sub

Function Text_TurnToUpperCase(Txt As String)
Text_TurntoUCase = UCase(Txt)
End Function

Function Text_TurnToLowerCase(Txt As String)
Text_TurntoLCase = LCase(Txt)
End Function
Sub ULAlign(frm As Form)
    Dim X, Y                    ' New top, left for the form
    Y = 0
    X = 0
    frm.Move X, Y             ' Change location of the form

End Sub

Sub PlayWav(file)
SoundName$ = file
SoundFlags& = &H20000 Or &H1
Snd& = sndPlaySound(SoundName$, SoundFlags&)
End Sub


Sub AOLChangeCaption(newcaption As String)
Call AOLSetText(AOLWindow(), newcaption)
End Sub

Sub AOLBuddyBLOCK(SN As textbox)
BUDLIST% = FindChildByTitle(AOLMDI(), "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_HWNDNEXT)
setup% = GetWindow(IM1%, GW_HWNDNEXT)
AOLIcon (setup%)
Pause (2)
STUPSCRN% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(Edit%, GW_HWNDNEXT)
view% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(view%, GW_HWNDNEXT)
AOLIcon PRCYPREF%
Pause (1.8)
Call AOLKillWindow(STUPSCRN%)
Pause (2)
PRYVCY% = FindChildByTitle(AOLMDI(), "Privacy Preferences")
DABUT% = FindChildByTitle(PRYVCY%, "Block only those people whose screen names I list")
AOLButton (DABUT%)
DaPERSON% = FindChildByClass(PRYVCY%, "_AOL_EDIT")
Call AOLSetText(DaPERSON%, SN)
Creat% = FindChildByClass(PRYVCY%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
AOLIcon Edit%
Pause (1)
Save% = GetWindow(Edit%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
AOLIcon Save%
End Sub
Public Sub AOLKillWindow(windo)
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub
Sub XAOL4_15Liner(Txt As String)
'Max of 14 chr or else u get Msg is too long
Call XAOL4_SetFocus
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.8
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.8
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.8
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.8
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.3
XAOL4_ChatSend "" + Txt + "" & c$ & "" + Txt + ""
Pause 0.8
End Sub
Sub XAOL4_AntiIdle()
'Sub contributed and written by ieet xero.
'If you would like to contact ieet xero,
'please email Jaguar at Jaguar32X@Juno.com
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AoIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
AOLIcon (AoIcon%)
End Sub

Public Sub XAOL4_AddRoom(Listboxes As ListBox)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = XAOL4_FindRoom()
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
Listboxes.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub


Sub XAOL4_BuddyVIEW()
Call XAOL4_Keyword("Buddy View")
End Sub
Sub XAOL4_BudList(Lst As ListBox)
'This adds the AOL Buddy List to a VB listbox
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = FindChildByTitle(AOLMDI(), "Buddy List Window")
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
Lst.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub XAOL4_ChangeCaption(newcaption As String)
Call AOLSetText(AOLWindow(), newcaption)
End Sub
Sub XAOL4_ChatManipulator(Who$, What$)
'This makes the chat room text near the VERY TOP
'what u want
view% = FindChildByClass(XAOL4_FindRoom(), "RICHCNTL")
Buffy$ = Chr$(13) & Chr$(10) & "" & (Who$) & ":" & Chr$(9) & "" & (What$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub XAOL4_ChatSend(Txt)
    Room% = XAOL4_FindRoom()
    If Room% Then
        hChatEdit% = Find2ndChildByClass(Room%, "RICHCNTL")
        ret = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, Txt)
        ret = SendMessageByNum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub
Function Find2ndChildByClass(parentw, childhand)
'DO NOT TAMPER WITH THIS CODE!
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    While firs%
        firs% = GetWindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Wend
    Find2ndChildByClass = 0
Found:
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = firs%
End Function

Sub XAOL4_ClearChat()
childs% = XAOL4_FindRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Function XAOL4_CountMail()
'This is an advanced mail counter. Don't change the code.
Dim PASS As Integer
PASS = 0
begin:
themail% = FindChildByTitle(AOLMDI(), XAOL4_GetUser & "'s Online Mailbox")
If themail% = 0 Then
Call XAOL4_MailReadNew
GoTo begin
End If
PASS = PASS + 1
If PASS <> 10 Then
GoTo begin
Else
tabcont% = FindChildByClass(themail%, "_AOL_TabControl")
TabPage% = FindChildByClass(tabcont%, "_AOL_TabPage")
thetree% = FindChildByClass(TabPage%, "_AOL_Tree")
XAOL4_CountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
AOLClose (themail%)
End If
End Function

Function XAOL4_FindRoom()
'Finds the chat room and sets focus on it
    aol% = FindWindow("AOL Frame25", vbNullString) '   MDI% = FindChildByClass(AOL%, "MDIClient")
    mdi% = FindChildByClass(aol%, "MDIClient")
    firs% = GetWindow(mdi%, 5)
    listers% = FindChildByClass(firs%, "RICHCNTL")
    listere% = FindChildByClass(firs%, "RICHCNTL")
    listerb% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (listers% = 0 Or listere% = 0 Or listerb% = 0) And (l <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            listers% = FindChildByClass(firs%, "RICHCNTL")
            listere% = FindChildByClass(firs%, "RICHCNTL")
            listerb% = FindChildByClass(firs%, "_AOL_Listbox")
            If listers% And listere% And listerb% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
       XAOL4_FindRoom = firs%
       Exit Function
     End If
End Function

Function XAOL4_FindToolbar()
ToolBar% = FindChildByClass(AOLWindow, "AOL Toolbar")
toolbar2% = FindChildByClass(ToolBar%, "_AOL_Toolbar")
XAOL4_FindToolbar = toolbar2%
End Function
Function XAOL4_GetChat()
'This gets all the txt from chat room
childs% = XAOL4_FindRoom
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
XAOL4_GetChat = theview$
End Function
Function XAOL4_GetCurrentRoomName()
XAOL4_GetCurrentRoomName = GetCaption(XAOL4_FindRoom)
End Function
Function XAOL4_GetUser()
On Error Resume Next
welcome% = FindChildByTitle(AOLMDI, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
XAOL4_GetUser = User
End Function
Sub XAOL4_Hide()
a = ShowWindow(AOLWindow(), SW_HIDE)
End Sub
Sub XAOL4_IMOff()
Call XAOL4_InstantMessage("$IM_OFF", "Jaguar32")
WaitForOk
End Sub
Sub XAOL4_IMOn()
Call XAOL4_InstantMessage("$IM_ON", "Jaguar32")
WaitForOk
End Sub
Sub XAOL4_InstantMessage(person, Message)
Call XAOL4_Keyword("aol://9293:" & person)
Pause (2)
Do
DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
Loop Until (im% <> 0 And aolrich% <> 0 And IMSEND% <> 0)
Call SendMessageByString(aolrich%, WM_SETTEXT, 0, Message)
For sends = 1 To 9
IMSEND% = GetWindow(IMSEND%, GW_HWNDNEXT)
Next sends
AOLIcon IMSEND%
If im% Then Call AOLKillWindow(im%)
End Sub

Sub XAOL4_Keyword(Txt)
    aol% = FindWindow("AOL Frame25", vbNullString)
    temp% = FindChildByClass(aol%, "AOL Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Toolbar")
    temp% = FindChildByClass(temp%, "_AOL_Combobox")
    KWBox% = FindChildByClass(temp%, "Edit")
    Call SendMessageByString(KWBox%, WM_SETTEXT, 0, Txt)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(KWBox%, WM_CHAR, VK_RETURN, 0)
End Sub
Sub XAOL4_LocateMember(name As String)
Call XAOL4_Keyword("aol://3548:" + name)
End Sub
Sub XAOL4_MailReadNew()
mailicon% = FindChildByClass(XAOL4_FindToolbar, "_AOL_Icon")
AOLIcon (mailicon%)
End Sub
Sub XAOL4_Mail(person, subject, Message)
Const LBUTTONDBLCLK = &H203
aol% = FindWindow("AOL Frame25", vbNullString)
Tool% = FindChildByClass(aol%, "AOL Toolbar")
tool2% = FindChildByClass(Tool%, "_AOL_Toolbar")
ico3n% = FindChildByClass(tool2%, "_AOL_Icon")
icon2% = GetWindow(ico3n%, 2)
X = SendMessageByNum(icon2%, WM_LBUTTONDOWN, 0&, 0&)
X = SendMessageByNum(icon2%, WM_LBUTTONUP, 0&, 0&)
Pause (4)
    aol% = FindWindow("AOL Frame25", vbNullString)
    mdi% = FindChildByClass(aol%, "MDIClient")
    Mail% = FindChildByTitle(mdi%, "Write Mail")
    aoledit% = FindChildByClass(Mail%, "_AOL_Edit")
    aolrich% = FindChildByClass(Mail%, "RICHCNTL")
    subjt% = FindChildByTitle(Mail%, "Subject:")
    subjec% = GetWindow(subjt%, 2)
        Call AOLSetText(aoledit%, person)
        Call AOLSetText(subjec%, subject)
        Call AOLSetText(aolrich%, Message)
e = FindChildByClass(Mail%, "_AOL_Icon")
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
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
AOLIcon (e)
End Sub

Public Sub XAOL4_MassIM(Lst As ListBox, Txt As textbox)
Lst.Enabled = False
i = Lst.ListCount - 1
Lst.ListIndex = 0
For X = 0 To i
Lst.ListIndex = X
Call XAOL4_InstantMessage(Lst.Text, Txt.Text)
Pause (1)
Next X
Lst.Enabled = True
End Sub
Sub XAOL4_OpenChat()
XAOL4_Keyword ("PC")
End Sub
Sub XAOL4_OpenPR(PRrm As String)
Call XAOL4_Keyword("aol://2719:2-2-" & PRrm)
End Sub

Sub XAOL4_Read1Mail()
'This will read the very first mail in the User's box
themail% = FindChildByTitle(AOLMDI(), XAOL4_GetUser & "'s Online Mailbox")
If themail% = 0 Then
Exit Sub
End If
e = FindChildByClass(themail%, "_AOL_Icon")
AOLIcon (e)
End Sub
Function XAOL4_RoomCount()
thechild% = XAOL4_FindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")
getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
XAOL4_RoomCount = getcount
End Function
Sub XAOL4_SetFocus()
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Sub XAOL4_SignOff()
Call RunMenuByString(AOLWindow(), "Sign Off")
End Sub
Function XAOL4_SpiralScroll(Txt As String)
Dim AODCOUNTER, a, thetxtlen
AODCOUNTER = 1
thetxtlen = Len(Txt)
Start:
a = a + 1
If a = thetxtlen Then GoTo last
X = Text_Looping(Txt)
Txt = X
XAOL4_ChatSend X
Pause (0.5)
AODCOUNTER = AODCOUNTER + 1
If AODCOUNTER = 4 Then
   AODCOUNTER = 2
   End If
GoTo Start
last:

End Function

Sub XAOL4_UnHide()
a = ShowWindow(AOLWindow(), SW_SHOW)
End Sub
Sub XAOL4_UnUpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(aolmod%, SW_RESTORE)
Call XAOL4_SetFocus
End Sub
Function XAOL4_UpChat()
die% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(die%, SW_HIDE)
X = ShowWindow(die%, SW_MINIMIZE)
Call XAOL4_SetFocus
End Function
Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
Function AOLGetTopWindow()
AOLGetTopWindow = GetTopWindow(AOLMDI())
End Function

Sub AOLSetFocus()
'SetFocusAPI doesn't work AOL because AOL has added
'a safeguard against other programs calling certain
'API functions (like owner-drawn things and like.)
'This is the only way known for setting the focus
'to AOL.  This is a normal VB command!
X = GetCaption(AOLWindow)
AppActivate X
End Sub
Public Sub AOLMassIM(Lst As ListBox, Txt As textbox)
Lst.Enabled = False
i = Lst.ListCount - 1
Lst.ListIndex = 0
For X = 0 To i
Lst.ListIndex = X
Call AOLInstantMessage(Lst.Text, Txt.Text)
Pause 0.5
Next X
Lst.Enabled = True
End Sub
Public Sub AOLOnlineChecker(person)
Call AOLInstantMessage4(person, "Sup?")
Pause 2
AOLIMScan
End Sub
Public Sub AddRoom_ByBox(Listboxes As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

Room = FindChildByTitle(AOLMDI, "Who's Chatting")
If Room = 0 Then MsgBox "Not Open"
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
OLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
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
LOP = Len(person$)
person$ = right$(person$, LOP - 2)
person$ = person$ & "@AOL.COM"
Listboxes.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub AddRoom(Lst As ListBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
namez$ = String$(256, " ")
ret = AOLGetList(Index, namez$)
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)))
ADD_AOL_LB namez$, Lst
Next Index
end_addr:
Lst.RemoveItem Lst.ListCount - 1
i = GetListIndex(Lst, AOLGetUser())
If i <> -2 Then Lst.RemoveItem i
End Sub

Public Sub AddRoom_WithExt(Listboxes As ListBox)
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
person$ = person$ & "@AOL.COM"
Listboxes.AddItem person$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub


Function ListToList(Source, destination)
counts = SendMessage(Source, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(destination, LB_ADDSTRING, 0, Buffer$)
Next Adding

End Function

Function MouseOverHwnd()
    ' Declares
      Dim pt32 As POINTAPI
      Dim ptx As Long
      Dim pty As Long
   
      Call GetCursorPos(pt32)               ' Get cursor position
      ptx = pt32.X
      pty = pt32.Y
      MouseOverHwnd = WindowFromPointXY(ptx, pty)    ' Get window cursor is over
End Function

Function UntilWindowClass(parent, news$)
Do: DoEvents
e = FindChildByClass(parent, news$)
Loop Until e
UntilWindowClass = e
End Function


Function UntilWindowTitle(parent, news$)
Do: DoEvents
e = FindChildByTitle(parent, news$)
Loop Until e
UntilWindowTitle = e
End Function
Public Function AOLGetList(Index, Buffer As String)
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





Function AddListToString(TheList As ListBox)
For DoList = 0 To TheList.ListCount - 1
AddListToString = AddListToString & TheList.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)

End Function
Public Sub XFader_TextColorChange(textbox As textbox, r, g, b)
'Contributed by TexRngr82@aol.com
    textbox.ForeColor = RGB(r, g, b)
    Pause 0.2
    textbox.ForeColor = RGB(r * 3 / 4, g * 3 / 4, b * 3 / 4)
    Pause 0.1
    textbox.ForeColor = RGB(r / 2, g / 2, b / 2)
    Pause 0.1
    textbox.ForeColor = RGB(r / 4, g / 4, b / 4)
    Pause 0.1
    textbox.ForeColor = RGB(0, 0, 0)
End Sub
Function AddListToMailString(TheList As ListBox)
If TheList.List(0) = "" Then GoTo last
For DoList = 0 To TheList.ListCount - 1
AddListToMailString = AddListToMailString & "(" & TheList.List(DoList) & "), "
Next DoList
AddListToMailString = Mid(AddListToMailString, 1, Len(AddListToMailString) - 2)
last:
End Function
Function ScrambleText(thetext, lbl)
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
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = scrambled$
lbl.Caption = ScrambleText
Exit Function
End Function

Public Sub AOLGetCurrentRoomName()
X = GetCaption(AOLFindRoom())
MsgBox X
End Sub
Function SearchForSelected(Lst As ListBox)
If Lst.List(0) = "" Then
counterf = 0
GoTo last
End If
counterf = -1

Start:
counterf = counterf + 1
If Lst.ListCount = counterf + 1 Then GoTo last
If Lst.Selected(counterf) = True Then GoTo last
If couterf = Lst.ListCount Then GoTo last
GoTo Start

last:
SearchForSelected = counterf
End Function



Sub AddStringToList(theitems, TheList As ListBox)
If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
TheList.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub

Sub Scroll8Line(Txt As textbox)
lonh = String(116, Chr(32))
d = 116 - Len(Text1)
c$ = Left(lonh, d)
AOLChatSend ("" & Txt & c$ & Text1)
AOLChatSend ("" & Txt & c$ & Text1)
lonh = String(116, Chr(32))
d = 116 - Len(Text1)
c$ = Left(lonh, d)
AOLChatSend ("" & Txt & c$ & Text1)
AOLChatSend ("" & Txt & c$ & Text1)
End Sub


Function AOLClickList(hWnd)
clicklist% = SendMessageByNum(hWnd, &H203, 0, 0&)
End Function

Function AOLCountMail()
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function AOLGetListString(parent, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

aolhandle = parent

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

Sub AOLHide()
a = ShowWindow(AOLWindow(), SW_HIDE)
End Sub

Sub AOLOpenChat()
If AOLFindRoom() Then Exit Sub
AOLKeyword ("pc")
Do: DoEvents
Loop Until AOLFindRoom()

End Sub
Public Sub AOLOpenNewMail()
Call AOLRunMenuByString("Read &New Mail")
End Sub


Public Sub AOLOpenOLDMail()
Call AOLRunMenuByString("Check Mail You've &Read")
End Sub
Public Sub AOLOpenSentMail()
Call AOLRunMenuByString("Check Mail You've &Sent")
End Sub
Public Sub AOLSignOnCaption(newcaption As String)
setup% = FindChildByTitle(AOLMDI(), "Welcome")
Call AOLSetText(setup%, newcaption)
End Sub
Sub AOLRespondIM(Message)
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo z
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo z
Exit Sub
z:
e = FindChildByClass(im%, "RICHCNTL")

e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e = GetWindow(e, 2)
e2 = GetWindow(e, 2) 'Send Text
e = GetWindow(e2, 2) 'Send Button
Call AOLSetText(e2, Message)
AOLIcon (e)
Pause 4
killwin (im%)
End Sub

Sub AOLRunMenuByString(stringer As String)
Call RunMenuByString(AOLWindow(), stringer)
End Sub


Sub AOLUnHide()
a = ShowWindow(AOLWindow(), SW_SHOW)
End Sub

Sub AOLWaitMail()
mailwin% = GetTopWindow(AOLMDI())
aoltree% = FindChildByClass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
Pause (10)
secondcount = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop


End Sub


Function EncryptType(Text, types)
'to encrypt, example:
'encrypted$ = EncryptType("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = EncryptType("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(Text)
If types = 0 Then
Current$ = Asc(Mid(Text, God, 1)) - 1
Else
Current$ = Asc(Mid(Text, God, 1)) + 1
End If
Process$ = Process$ & Chr(Current$)
Next God

EncryptType = Process$
End Function

Function FindChildByTitle(parent, child As String) As Integer
childfocus% = GetWindow(parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
Buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, Buffer$, (hwndLength% + 1))

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function

Function FindChildByClass(parent, child As String) As Integer
childfocus% = GetWindow(parent, 5)

While childfocus%
Buffer$ = String$(250, 0)
classbuffer% = GetClassName(childfocus%, Buffer$, 250)

If InStr(UCase(Buffer$), UCase(child)) Then FindChildByClass = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function


Function DescrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Descrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo city
lastchar$ = Mid(chars$, 2, 1)
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 3, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
city:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
scrambled$ = scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed buffers
sniff:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = scrambled$

End Function



Function GetLineCount(Text)

theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Sub HideWindow(hWnd)
hi = ShowWindow(hWnd, SW_HIDE)
End Sub


Function IntegerToString(tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Function LineFromText(Text, theline)
theview$ = Text

For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
c = c + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If theline = c Then GoTo ex
thechars$ = ""
End If

Next FindChar
Exit Function
ex:
thechatext$ = ReplaceText(thechatext$, Chr(13), "")
thechatext$ = ReplaceText(thechatext$, Chr(10), "")

LineFromText = thechatext$
End Function



Sub MaxWindow(hWnd)
ma = ShowWindow(hWnd, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hWnd)
MI2 = ShowWindow(hWnd, SW_MINIMIZE)
End Sub

Function NumericNumber(thenumber)
NumericNumber = Val(thenumber)
'turns the "number" so vb recognizes it for
'addition, subtraction, ect.
End Function

Sub ParentChange(parent%, location%)
doparent% = SetParent(parent%, location%)
End Sub


Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function ReverseText(Text)
For words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, words, 1)
Next words


End Function

Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
MenuItemCount% = GetMenuItemCount(ToSearchSub%)

For getstring = 0 To MenuItemCount% - 1
SubCount% = GetMenuItemID(ToSearchSub%, getstring)
MenuString$ = String$(100, " ")
GetStringMenu% = GetMenuString(ToSearchSub%, SubCount%, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem% = SubCount%
GoTo MatchString
End If

Next getstring

Next FindString
MatchString:
RunTheMenu% = SendMessage(Application, WM_COMMAND, MenuItem%, 0)
End Sub

Sub AOLRunTool(Tool)
ToolBar% = FindChildByClass(AOLWindow(), "AOL Toolbar")
iconz% = FindChildByClass(ToolBar%, "_AOL_Icon")
For X = 1 To Tool - 1
iconz% = GetWindow(iconz%, 2)
Next X
isen% = IsWindowEnabled(iconz%)
If isen% = 0 Then Exit Sub
AOLIcon (iconz%)
End Sub

Function ScrambleGame(thestring As Integer)
Dim bytestring As String
thestringcount = Len(thestring)
If Not Mid(thestring, thestringcount, 1) = " " Then thestring = thestring & " "
For Stringe = 1 To Len(thestring)
characters$ = Mid(thestring, Stringe, 1)
thestrings$ = thestrings$ & characters$

If characters$ = " " Then
smoked:

DoEvents
For Ensemble = 1 To Len(thestrings$) - 1
Randomize
randomstring = Int((Len(thestrings$) * Rnd) + 1)
If randomstring = Len(thestrings$) Then GoTo already
If bytesread Like "*" & randomstring & "*" Then GoTo already
stringrandom$ = Mid(thestrings$, randomstring, 1)
stringfound$ = stringfound$ & stringrandom$
bytesread = bytesread & randomstring
GoTo really
already:
Ensemble = Ensemble - 1
really:
Next Ensemble
If stringfound$ = thestrings$ Then stringfound$ = "": GoTo smoked
thestrings2$ = thestrings2$ & stringfound$ & " "
stringfound$ = ""
thestrings$ = ""
bytesread = ""
strngfound$ = ""
End If

Next Stringe
ScrambleGame = Mid(thestrings2$, 1, Len(thestring) - 1)
End Function


Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For Replace = 1 To Len(Text)
thechar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function


Sub SetBackPre()
Call RunMenuByString(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
AOLIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 0, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

AOLButton (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function AOLStayOnline()
hwndz% = FindWindow(AOLWindow(), "America Online")
childhwnd% = FindChildByTitle(hwndz%, "OK")
AOLButton (childhwnd%)
End Function

Public Sub CenterCorner(frmForm As Form)
'This will center you form in the upper right
'of the users screen
   With frmForm
      .Left = (Screen.Width - .Width) / 1
      .Top = (Screen.Height - .Height) / 2000
   End With
End Sub
Function StringToInteger(tochange As String) As Integer
StringToInteger = tochange
End Function
Function TrimCharacter(thetext, chars)
TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For TrimSpace = 1 To Len(Text)
thechar$ = Mid(Text, TrimSpace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next TrimSpace

TrimSpaces = thechars$
End Function


Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")
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


Function KTEncrypt(ByVal password, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(password) = 0 Then Error 31100
  
  'Is password too long
  If Len(password) > 255 Then Error 31100

  'Is there a strng$ to work with?
  If Len(strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + right$(strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
          strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
        End If
        look = look + 1
      Loop
      
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
    End If
  Else
    'force% flag set, ecrypt string regardless of tag
    EncryptFlag% = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(password)
  
  
  'Tack on leading characters to prevent repetative recognition
  password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
  password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
  password = password + Chr$(Asc(right$(password, 1)) Xor PassMax)
  password = password + Chr$(Asc(right$(password, 2)) Xor Asc(right$(password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(password, 3) + Format$(Asc(right$(password, 1)), "000") + Format$(Len(password), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(tochange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, strng, Chr$(1))
      If look > 0 Then
        strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(strng, Chr$(0))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(strng, Chr$(10))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(strng, Chr$(13))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(strng, Chr$(26))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(right$(password, 1)), "000") + Format$(Len(password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      strng = Mid$(strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function

Public Sub CenterForm(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub

Public Sub CenterFormTop(frm As Form)
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub

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

Public Sub AOLButton(but%)
ClickIcon% = SendMessage(but%, WM_KEYDOWN, VK_SPACE, 0)
ClickIcon% = SendMessage(but%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function AOLIMSTATIC(newcaption As String)
ANTI1% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
STS% = FindChildByClass(ANTI1%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call ChangeCaption(st%, newcaption)
End Function

Function AOLGetUser()
On Error Resume Next
aol& = FindWindow("AOL Frame25", "America  Online")
mdi& = FindChildByClass(aol&, "MDIClient")
welcome% = FindChildByTitle(mdi&, "Welcome, ")
WelcomeLength& = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength& + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User$
End Function

Sub AOLIMOff()
Call AOLInstantMessage("$IM_OFF", "º°˜¨˜°ºº°˜¨˜°º Jaguar32")
End Sub

Sub AOLIMsOn()
Call AOLInstantMessage("$IM_ON", "º°˜¨˜°ºº°˜¨˜°º Jaguar32")

End Sub


Sub AOLChatSend(Txt)
Room% = AOLFindRoom()
Call AOLSetText(FindChildByClass(Room%, "_AOL_Edit"), Txt)
DoEvents
Call SendCharNum(FindChildByClass(Room%, "_AOL_Edit"), 13)
'A1000% = FindChildByClass(Room%, "_AOL_Edit")
'A2000% = GetWindow(A1000%, 2)
'AOLIcon (A2000%)

End Sub


Sub AOLClose(winew)
closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub


Sub AOLCursor()
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Function AOLFindRoom()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "_AOL_Edit")
listere% = FindChildByClass(childfocus%, "_AOL_View")
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function
Function FindOpenMail()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByClass(childfocus%, "RICHCNTL")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindOpenMail = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend


End Function

Function FindForwardWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByTitle(childfocus%, "Send Now")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function



Function AOLGetChat()
childs% = AOLFindRoom()
child = FindChildByClass(childs%, "_AOL_View")


GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$
AOLGetChat = theview$
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
Function PlayMIDI(DriveDirFile As String, Optional loopIT As Boolean)
'Made by aura
Dim returnStr As String * 255
Dim Shortpath$, X&
Shortpath = Space(Len(DriveDirFile))
X = GetShortPathName(DriveDirFile, Shortpath, Len(Shortpath))
If X = 0 Then GoTo ErrorHandler
If X > Len(DriveDirFile) Then 'not a long filename
Shortpath = DriveDirFile
Else                          'it is a long filename
Shortpath = Left(Shortpath, X) 'x is the length of the return buffer
End If
X = mciSendString("close yada", returnStr, 255, 0) 'just in case
X = mciSendString("open " & Chr(34) & Shortpath & Chr(34) & " type sequencer alias yada", returnStr, 255, 0)
    If X <> 0 Then GoTo theEnd  'invalid filename or path
X = mciSendString("play yada", returnStr, 255, 0)
    If X <> 0 Then GoTo theEnd  'device busy or not ready
    If Not loopIT Then Exit Function
Do While DoEvents
    X = mciSendString("status yada mode", returnStr, 255, 0)
        If X <> 0 Then Exit Function 'StopMIDI() was pressed or error
    If Left(returnStr, 7) = "stopped" Then X = mciSendString("play yada from 1", returnStr, 255, 0)
Loop
Exit Function
theEnd:  'MIDI errorhandler
returnStr = Space(255)
X = mciGetErrorString(X, returnStr, 255)
MsgBox Trim(returnStr), vbExclamation 'error message
X = mciSendString("close yada", returnStr, 255, 0)
Exit Function

ErrorHandler:
MsgBox "Invalid Filename or Error.", vbInformation
End Function
Function Play_StopMIDI()
'Made by aura
Dim X&
Dim returnStr As String * 255
X = mciSendString("status yada mode", returnStr, 255, 0)
    If Left(returnStr, 7) = "playing" Then X = mciSendString("stop yada", returnStr, 255, 0)
returnStr = Space(255)
X = mciSendString("status yada mode", returnStr, 255, 0)
    If Left(returnStr, 7) = "stopped" Then X = mciSendString("close yada", returnStr, 255, 0)
End Function
Public Function PWSD_BVScan2(Filename$, Searchstring$) As Long
Free = FreeFile
Dim Where As Long
Open Filename$ For Binary Access Read As #Free
For X = 1 To LOF(Free) Step 32000
Text$ = Space(32000)
Get #Free, X, Text$
 Debug.Print X
 If InStr(1, Text$, Searchstring$, 1) Then
 Where = InStr(1, Text$, Searchstring$, 1)
FileSearch = (Where + X) - 1
 Close #Free
  Exit For
 End If
  Next X
Close #Free
End Function
Public Function PWSD_BVScan(Filename$, ByVal Searchstring$)
Dim Variant1 As Variant
Dim Variant2 As Variant
Dim Variant3 As Variant
Dim Variant4 As Variant
Dim Single1 As Single
Dim String1 As String
Dim EnterKey As String
On Error Resume Next
Open Filename$ For Binary As #1
EnterKey$ = Chr$(13) + Chr$(10)
msg$ = ""
Variant1 = LOF(1)
Variant2 = Variant1
Variant3 = 1
If Variant2 > 32000 Then
Variant4 = 32000
ElseIf Variant2 = 0 Then
Variant4 = 1
Else
Variant4 = Variant2
End If
StringA$ = String$(Variant4, " ")
Get #1, Variant3, String1$
Single1! = InStr(1, String1$, Searchstring$, 1)
If Single1! Then
PWSD_BVScan = 0
Else
PWSD_BVScan = 1
End If
Close #1
End Function

Function XFader_CLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This gets a color from 3 scroll bars
XFader_CLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)
'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = XFader_CLRBars(RedScroll1, GreenScroll1, BlueScroll1)
End Function
Function XFader_FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, thetext$, Wavy As Boolean)
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
dacolor4$ = XFader_RGBtoHEX(Colr4)
dacolor5$ = XFader_RGBtoHEX(Colr5)
dacolor6$ = XFader_RGBtoHEX(Colr6)
dacolor7$ = XFader_RGBtoHEX(Colr7)
dacolor8$ = XFader_RGBtoHEX(Colr8)
dacolor9$ = XFader_RGBtoHEX(Colr9)
dacolor10$ = XFader_RGBtoHEX(Colr10)
rednum1% = Val("&H" + right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))
XFader_FadeByColor10 = XFader_FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, thetext, Wavy)
End Function

Function XFader_FadeByColor2(Colr1, Colr2, thetext$, Wavy As Boolean)
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
rednum1% = Val("&H" + right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
XFader_FadeByColor2 = XFader_FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, thetext, Wavy)
End Function

Function XFader_FadeByColor3(Colr1, Colr2, Colr3, thetext$, Wavy As Boolean)
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
rednum1% = Val("&H" + right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
XFader_FadeByColor3 = XFader_FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, thetext, Wavy)
End Function
Function XFader_FadeByColor4(Colr1, Colr2, Colr3, Colr4, thetext$, Wavy As Boolean)
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
dacolor4$ = XFader_RGBtoHEX(Colr4)
rednum1% = Val("&H" + right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
XFader_FadeByColor4 = XFader_FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, thetext, Wavy)
End Function
Function XFader_FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, thetext$, Wavy As Boolean)
dacolor1$ = XFader_RGBtoHEX(Colr1)
dacolor2$ = XFader_RGBtoHEX(Colr2)
dacolor3$ = XFader_RGBtoHEX(Colr3)
dacolor4$ = XFader_RGBtoHEX(Colr4)
dacolor5$ = XFader_RGBtoHEX(Colr5)
rednum1% = Val("&H" + right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
XFader_FadeByColor5 = XFader_FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, thetext, Wavy)
End Function
Function XFader_FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, thetext$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = right(thetext, frthlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    
    XFader_FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function


Sub XFader_FadeForm(FormX As Form, Colr1, Colr2)

    B1 = XFader_GetRGB(Colr1).Blue
    G1 = XFader_GetRGB(Colr1).Green
    R1 = XFader_GetRGB(Colr1).Red
    B2 = XFader_GetRGB(Colr2).Blue
    G2 = XFader_GetRGB(Colr2).Green
    R2 = XFader_GetRGB(Colr2).Red
    
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
Function XFader_FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, thetext$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = right(thetext, thrdlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    
    XFader_FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Sub XFader_FadePreview(PicB As PictureBox, ByVal FadedText As String)
FadedText$ = XFader_Replacer(FadedText$, Chr(13), "+chr13+")
OSM = PicB.ScaleMode
PicB.ScaleMode = 3
TextOffX = 0: TextOffY = 0
StartX = 2: StartY = 0
PicB.Font = "Arial": PicB.FontSize = 10
PicB.FontBold = False: PicB.FontItalic = False: PicB.FontUnderline = False: PicB.FontStrikethru = False
PicB.AutoRedraw = True: PicB.ForeColor = 0&: PicB.Cls
For X = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, X, 1)
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    Select Case T$
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
        TextOffY = -1
      Case "/sup" 'end superscript
        TextOffY = 0
      Case "sub"  'start subscript
        TextOffY = 1
      Case "/sub" 'end subscript
        TextOffY = 0
      Case Else
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = right$(ColorString$, 2)
          RV = XFader_Hex2Dec!(RedString$)
          GV = XFader_Hex2Dec!(GreenString$)
          BV = XFader_Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then
            fontstart% = InStr(T$, Chr(34))
            dafont$ = right(T$, Len(T$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If c$ = "+" And Mid(FadedText$, X, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        X = X + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print c$
        TextOffX = TextOffX + PicB.TextWidth(c$)
    End If
  End If
Next X
PicB.ScaleMode = OSM
End Sub
Sub XFader_FadePreview2(RichTB As Control, ByVal FadedText As String)
'NOTE: RichTB must be a RichTextBox.
'NOTE: You cannot preview wavy fades with this sub.
Dim StartPlace%
StartPlace% = 0
RichTB.SelStart = StartPlace%
RichTB.Font = "Arial": RichTB.SelFontSize = 10
RichTB.SelBold = False: RichTB.SelItalic = False: RichTB.SelUnderline = False: RichTB.SelStrikeThru = False
RichTB.SelColor = 0&: RichTB.Text = ""
For X = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, X, 1)
  RichTB.SelStart = StartPlace%
  RichTB.SelLength = 1
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    RichTB.SelStart = StartPlace%
    RichTB.SelLength = 1
    Select Case T$
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
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = right$(ColorString$, 2)
          RV = XFader_Hex2Dec!(RedString$)
          GV = XFader_Hex2Dec!(GreenString$)
          BV = XFader_Hex2Dec!(BlueString$)
          RichTB.SelStart = StartPlace%
          RichTB.SelColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then
            fontstart% = InStr(T$, Chr(34))
            dafont$ = right(T$, Len(T$) - fontstart%)
            RichTB.SelStart = StartPlace%
            RichTB.SelFontName = dafont$
        End If
    End Select
  Else  'normal text
    RichTB.SelText = RichTB.SelText + c$
    StartPlace% = StartPlace% + 1
    RichTB.SelStart = StartPlace%
  End If
Next X
End Sub
Function XFader_FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, thetext$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = right(thetext, ninelen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9$)
    For i = 1 To textlen%
        TextDone$ = Left(part9$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    XFader_FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function

Function XFader_FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, thetext$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(thetext, fstlen%)
    part2$ = right(thetext, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
    
    
    XFader_FadeThreeColor = Faded1$ + Faded2$
End Function



Function XFader_FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, thetext$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen$ = Len(thetext)
    For i = 1 To textlen$
        TextDone$ = Left(thetext, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = XFader_RGBtoHEX(ColorX)
        
        If Wavy = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    XFader_FadeTwoColor = Faded$
End Function
Function XFader_GetRGB(ByVal CVal As Long) As COLORRGB
  XFader_GetRGB.Blue = Int(CVal / 65536)
  XFader_GetRGB.Green = Int((CVal - (65536 * XFader_GetRGB.Blue)) / 256)
  XFader_GetRGB.Red = CVal - (65536 * XFader_GetRGB.Blue + 256 * XFader_GetRGB.Green)
End Function
Function XFader_GETVAL%(ByVal strLetter$)

  Select Case strLetter$
    Case "0"
      XFader_GETVAL = 0
    Case "1"
      XFader_GETVAL = 1
    Case "2"
      XFader_GETVAL = 2
    Case "3"
      XFader_GETVAL = 3
    Case "4"
      XFader_GETVAL = 4
    Case "5"
      XFader_GETVAL = 5
    Case "6"
      XFader_GETVAL = 6
    Case "7"
      XFader_GETVAL = 7
    Case "8"
      XFader_GETVAL = 8
    Case "9"
      XFader_GETVAL = 9
    Case "A"
      XFader_GETVAL = 10
    Case "B"
      XFader_GETVAL = 11
    Case "C"
      XFader_GETVAL = 12
    Case "D"
      XFader_GETVAL = 13
    Case "E"
      XFader_GETVAL = 14
    Case "F"
      XFader_GETVAL = 15
  End Select
End Function
Function XFader_Hex2Dec!(ByVal strHex$)

  If Len(strHex$) > 8 Then strHex$ = right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = XFader_GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function
Function XFader_HTMLtoRGB(TheHTML$)
'Converts HTML such as 0000FF to an
'RGB value like &HFF0000 so you can
'use it in the FadeByColor functions
If Left(TheHTML$, 1) = "#" Then TheHTML$ = right(TheHTML$, 6)

RedX$ = Left(TheHTML$, 2)
GreenX$ = Mid(TheHTML$, 3, 2)
BlueX$ = right(TheHTML$, 2)
rgbhex$ = "&H00" + BlueX$ + GreenX$ + RedX$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function XFader_InverseColor(OldColor)
dacolor$ = XFader_RGBtoHEX(OldColor)
RedX% = Val("&H" + right(dacolor$, 2))
GreenX% = Val("&H" + Mid(dacolor$, 3, 2))
BlueX% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - RedX%
newgreen% = 255 - GreenX%
newblue% = 255 - BlueX%
InverseColor = RGB(newred%, newgreen%, newblue%)
End Function

Function XFader_Replacer(TheStr As String, This As String, WithThis As String)
Dim STRwo13s As String
STRwo13s = TheStr
Do While InStr(1, STRwo13s, This)
DoEvents
thepos% = InStr(1, STRwo13s, This)
STRwo13s = Left(STRwo13s, (thepos% - 1)) + WithThis + right(STRwo13s, Len(STRwo13s) - (thepos% + Len(This) - 1))
Loop

Replacer = STRwo13s
End Function
Function XFader_RGBtoHEX(RGB)

    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function
Function XFader_Rich2HTML(RichTXT As Control, StartPos%, EndPos%)
Dim Bolded As Boolean
Dim Undered As Boolean
Dim Striked As Boolean
Dim Italiced As Boolean
Dim LastCRL As Long
Dim LastFont As String
Dim HTMLString As String

For posi% = StartPos To EndPos
RichTXT.SelStart = posi%
RichTXT.SelLength = 1

If Bolded <> RichTXT.SelBold Or posi% = StartPos Then
If RichTXT.SelBold = True Then
HTMLString = HTMLString + "<b>"
Bolded = True
Else
HTMLString = HTMLString + "</b>"
Bolded = False
End If
End If

If Undered <> RichTXT.SelUnderline Or posi% = StartPos Then
If RichTXT.SelUnderline = True Then
HTMLString = HTMLString + "<u>"
Undered = True
Else
HTMLString = HTMLString + "</u>"
Undered = False
End If
End If

If Striked <> RichTXT.SelStrikeThru Or posi% = StartPos Then
If RichTXT.SelStrikeThru = True Then
HTMLString = HTMLString + "<s>"
Striked = True
Else
HTMLString = HTMLString + "</s>"
Striked = False
End If
End If

If Italiced <> RichTXT.SelItalic Or posi% = StartPos Then
If RichTXT.SelItalic = True Then
HTMLString = HTMLString + "<i>"
Italiced = True
Else
HTMLString = HTMLString + "</i>"
Italiced = False
End If
End If

If LastCRL <> RichTXT.SelColor Or posi% = StartPos Then
ColorX = RGB(XFader_GetRGB(RichTXT.SelColor).Blue, XFader_GetRGB(RichTXT.SelColor).Green, XFader_GetRGB(RichTXT.SelColor).Red)
colorhex = XFader_RGBtoHEX(ColorX)
HTMLString = HTMLString + "<Font Color=#" & colorhex & ">"
LastCRL = RichTXT.SelColor
End If

If LastFont <> RichTXT.SelFontName Then
HTMLString = HTMLString + "<font face=" + Chr(34) + RichTXT.SelFontName + Chr(34) + ">"
LastFont = RichTXT.SelFontName
End If

HTMLString = HTMLString + RichTXT.SelText
Next posi%

Rich2HTML = HTMLString

End Function


Sub XFader_FormFade(FormX As Form, Colr1, Colr2)

    B1 = XFader_GetRGB(Colr1).Blue
    G1 = XFader_GetRGB(Colr1).Green
    R1 = XFader_GetRGB(Colr1).Red
    B2 = XFader_GetRGB(Colr2).Blue
    G2 = XFader_GetRGB(Colr2).Green
    R2 = XFader_GetRGB(Colr2).Red
    
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


Sub XAOL4_PRBust(Room$)
'Sub contributed and written by ieet xero.
'sub fixed by uzer
'If you would like to contact ieet xero,
'please email Jaguar at Jaguar32X@Juno.com

chat = XAOL4_GetCurrentRoomName
Do
Call XAOL4_OpenPR(Room$)
If XAOL4_GetCurrentRoomName = Room$ Then GoTo xero
aol% = FindWindow("#32770", "America Online")
If aol% Then
closeaol% = SendMessage(aol%, WM_CLOSE, 0, 0)
Pause 1
End If
Pause 1
Loop
xero:
End Sub


Sub XAOL4_KillModal()
'Sub contributed and written by ieet xero.
'If you would like to contact ieet xero,
'please email Jaguar at Jaguar32X@Juno.com
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call AOLKillWindow(Modal%)
End Sub
Sub AOLAntiIdle()
aol% = FindWindow("_AOL_Modal", vbNullString)
xstuff% = FindChildByTitle(aol%, "Favorite Places")
If xstuff% Then Exit Sub
xstuff2% = FindChildByTitle(aol%, "File Transfer *")
If xstuff2% Then Exit Sub
yes% = FindChildByClass(aol%, "_AOL_Button")
AOLButton yes%
End Sub

Sub AOLGetMemberProfile(name As String)
AOLRunMenuByString ("Get a Member's Profile")
Pause 0.3
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(mdi%, "Get a Member's Profile")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
End Sub


Function FindIMTextwindow()
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMTextwindow = FindChildByClass(im%, "RICHCNTL")
End Function
Function FindIMCaption()
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
FindIMCaption = FindChildByClass(im%, "_AOL_Static")
End Function
Function AOLChangeIMCaption(Txt As String)
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
IMtext% = FindChildByClass(im%, "_AOL_Static")
Call ChangeCaption(IMtext%, Txt)
End Function
Function Mail_GetErrorMessage()
Errors% = FindChildByTitle(AOLMDI(), "Error")
IMtext% = FindChildByClass(Errors%, "_AOL_VIEW")
Mail_GetErrorMessage = AOLGetText(IMtext%)
End Function

Function MakeSpaceInGoto(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchr$ = " " Then Let nextchr$ = "%20"
Let newsent$ = newsent$ + nextchr$
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
MakeSpaceInGoto = newsent$
End Function

Sub AOLAntiPunter()
Do
ANT% = FindChildByTitle(AOLMDI(), "Untitled")
IMRich% = FindChildByClass(ANT%, "RICHCNTL")
STS% = FindChildByClass(ANT%, "_AOL_Static")
st% = GetWindow(STS%, GW_HWNDNEXT)
st% = GetWindow(st%, GW_HWNDNEXT)
Call AOLSetText(st%, "SouthPark FINAL - This IM Window Should Remain OPEN.")
mi = ShowWindow(ANT%, SW_MINIMIZE)
DoEvents:
If IMRich% <> 0 Then
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
Lab = SendMessageByNum(IMRich%, WM_CLOSE, 0, 0)
End If
Loop
End Sub

Sub AOLChangeWavDirect(wav As String)
'change the directory of the wav's
'AOLChangeWavDirect("C:\aol30\download")
Open "C:\AOL25\tool\chat.aol" For Binary As #1
Seek #1, 6935
Put #1, , wav
Close #1
End Sub
Public Sub AOLChangeWelcome(newwelcome As String)
Welc% = FindChildByTitle(AOLMDI(), "Welcome, " & AOLGetUser & "!")
Call AOLSetText(Welc%, newwelcome)
End Sub
Public Sub AOLChatManipulator(Who$, What$)
view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "" & (Who$) & ":" & Chr$(9) & "" & (What$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLClearChatRoom()
'clears the chat room
X$ = Format$(String$(100, Chr$(13)))
Call AOLChatManipulator(" ", X$)
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
AOLChatSend "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Pause 0.3
AOLChatSend "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Pause 0.3
AOLChatSend "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Pause 0.3
AOLChatSend "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
Pause 0.7
End Sub


Sub WaitForLoadedMail()
Do
Box = FindChildByTitle(FindWindow("AOL Frame25", "America  Online"), "New Mail")
Pause (0.1)
Loop Until Box <> 0
List = FindChildByClass(Box, "_AOL_Tree")
Do
DoEvents
MailNum = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call Pause(0.5)
MailNum2 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Call Pause(0.5)
MailNum3 = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until MailNum = MailNum2 And MailNum2 = MailNum3
    MailNum = SendMessage(List, LB_GETCOUNT, 0, 0&)

End Sub
Sub AOLHostManipulator(What$)
'AOLHostManipulator "You are gay"
'This will make the online host say you are gay!
view% = FindChildByClass(AOLFindRoom(), "_AOL_View")
Buffy$ = Chr$(13) & Chr$(10) & "OnlineHost:" & Chr$(9) & "" & (What$) & ""
X% = SendMessageByString(view%, WM_SETTEXT, 0, Buffy$)
End Sub
Sub AOLHostNameChange(SN As String)
X = AOLVersion()
If X = "C:\aol30" Then
Open "C:\aol30\tool\aolchat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
ElseIf X = "C:\aol25" Then
Open "C:\aol25\tool\chat.aol" For Binary As #1
Seek #1, 6887
Put #1, , SN

Close #1
End If
End Sub
Function AOLFindChatWindow() As Integer
  Dim genhWnd%
  Dim AOLChildhWnd%
  Dim ChildWnd As Integer
  Dim ControlWnd As Integer
  Dim ChatWnd As Integer
  Dim TargetsFound As Integer
  Dim RetClsName As String * 255
  Dim X%
genhWnd% = GetWindow(FindWindow("AOL Frame25", 0&), GW_CHILD)
Do
  X% = GetClassName(genhWnd%, RetClsName$, 254)
    If InStr(RetClsName$, "MDIClient") Then
      AOLChildhWnd% = genhWnd% 'Child window found!
    End If
  genhWnd% = GetWindow(genhWnd%, GW_HWNDNEXT)
  DoEvents
Loop While genhWnd% <> 0
ChildWnd = GetWindow(AOLChildhWnd%, GW_CHILD)
Do
  ControlWnd = GetWindow(ChildWnd, GW_CHILD)
  Do
    X% = GetClassName(ControlWnd, RetClsName$, 254)

    
    If InStr(RetClsName$, "_AOL_Edit") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_View") Then
      TargetsFound = TargetsFound + 1
    ElseIf InStr(RetClsName$, "_AOL_Listbox") Then
      TargetsFound = TargetsFound + 1:
    End If
    ControlWnd = GetWindow(ControlWnd, GW_HWNDNEXT)
    DoEvents
  Loop While ControlWnd <> 0

  If TargetsFound = 3 Then ChatWnd = ChildWnd: Exit Do

  
  ChildWnd = GetWindow(ChildWnd, GW_HWNDNEXT)
Loop While ChildWnd <> 0
Chat_FindTheWin = ChatWnd

End Function
Sub Click(Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub
Function AOLKillDupes()
'Gets rid of excess mail
num = AOLCountMail()
DELNUM% = 0
aol% = FindWindow("AOL Frame25", 0&)
mld% = FindChildByClass(aol%, "MDIClient")
FNMB% = FindChildByTitle(mld%, "New Mail")
If FNMB% = 0 Then
AOLRunTool (1)
Pause (0.4)
Do: DoEvents
    NMB% = FindChildByTitle(mofo%, "New Mail")
    Pause (0.1)
Loop Until NMB% <> 0
WaitForLoadedMail
End If

Do: DoEvents
LSTTXT$ = ","
DELTXT$ = ","
btnDEL% = FindChildByTitle(FindChildByTitle(mofo%, "New Mail"), "Delete")
If AOLCountMail() = 0 Then MsgBox "You have no New Mail.", 12, "Dupe Killer": Exit Function
List% = FindChildByClass(FindChildByTitle(mofo%, "New Mail"), "_AOL_Tree")
For i = 0 To AOLCountMail() - 1
Ln = SendMessage(List%, LB_GETTEXTLEN, i, 0)
If Ln = -1 And i >= AOLCountMail() Then
    Exit For
ElseIf Ln = -1 And i <= AOLCountMail() Then
    MAILTXT$ = String$(60, 0)
Else
    MAILTXT$ = String$(Ln, 0)
End If
GTTXT = SendMessageByString(List%, LB_GETTEXT, i, MAILTXT$)
MAILTXT$ = right$(MAILTXT$, Len(MAILTXT$) - InStr(InStr(MAILTXT$, Chr$(9)) + 1, MAILTXT$, Chr$(9)))
If InStr(LSTTXT$, "," & MAILTXT$ & ",") And InStr(DELTXT$, "," & MAILTXT$ & ",") = 0 Then
            X = SendMessage(List%, LB_SETCURSEL, i, 0)
            Call Click(btnDEL%)
            DELNUM% = DELNUM% + 1
            num = num - 1
            i = i - 1
            DELTXT$ = DELTXT$ + MAILTXT$ + ","
Else
LSTTXT$ = LSTTXT$ + MAILTXT$ + ","
End If
Next i
Loop Until Len(DELTXT$) = 1
MsgBox "There were " & DELNUM% & " duplicate mails deleted.", 12, "Dupe Count"
Mail_KillDupes = DELNUM%
End Function
Sub AOLMakeMeParent(frm As Form)
'AOLMakeParent Me
'this makes the form an aol parent
aol% = FindChildByClass(FindWindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = SetParent(frm.hWnd, aol%)
End Sub
Sub AOLMail(person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(Mess%, WM_SETTEXT, 0, Message)
AOLIcon (icone%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub AOLLocateMember(name As String)
'locates, if possible, member "name"
AOLRunMenuByString ("Locate a Member Online")
Pause 0.3
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
prof% = FindChildByTitle(mdi%, "Locate Member Online")
putname% = FindChildByClass(prof%, "_AOL_Edit")
Call AOLSetText(putname%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Button")
AOLButton okbutton%
closes = SendMessage(prof%, WM_CLOSE, 0, 0)
End Sub

Function WriteErrorNameToList(Lst As ListBox)
'Not working yet 4/8/98
messa = Mail_GetErrorMessage
LC = GetLineCount(messa)

a = 1

Start:
thetext = LineFromText(messa, a)
Stcount = Len(thetext)
SC = Stcount - 34
AC2 = Left$(thetext, SC)
Lst.AddItem AC2
a = a + 1
GoTo Start
last:
End Function

Function MessageFromIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
IMtext% = FindChildByClass(im%, "RICHCNTL")
IMmessage = AOLGetText(IMtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(IMmessage, Len(IMmessage) - 1)
End Function

Sub SizeFormToWindow(frm As Form, win%)
Dim wndRect As RECT, lRet As Long
lRet = GetWindowRect(win%, wndRect)
With frm
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub
Sub StuffOff(frm As Form, btn As Object)
With btn
   .FontItalic = False
   .FontStrikethru = False
   .FontUnderline = False
End With
End Sub
Function SNfromIM()
im% = FindChildByTitle(AOLMDI(), ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
heh$ = GetCaption(im%)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
SNfromIM = naw$
End Function
Function AOLGetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

AOLGetText = TrimSpace$
End Function

Sub AOLIcon(icon%)
Click2% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click2% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLInstantMessage(person, Message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSEND% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, person)
Call AOLSetText(aolrich%, Message)
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
IMSEND% = GetWindow(IMSEND%, GW_HWNDNEXT)
Next sends
AOLIcon (IMSEND%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
End Sub
Sub AOLInstantMessage2(person)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, ">Instant Message From: ")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, person)
End Sub
Sub AOLInstantMessage3(person, Message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSEND% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, person)
Call AOLSetText(aolrich%, Message)
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
IMSEND% = GetWindow(IMSEND%, 2)
Next sends
AOLIcon (IMSEND%)
Loop Until im% = 0
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop

End Sub
Sub AOLInstantMessage4(person, Message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSEND% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, person)
Call AOLSetText(aolrich%, Message)
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
For sends = 1 To 9
IMSEND% = GetWindow(IMSEND%, 2)
Next sends
AOLIcon (IMSEND%)
End Sub

Function AOLChildIM()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
AOLChildIM = IMSEND%
End Function
Function AOLCreateIM(person, Message)
Call RunMenuByString(AOLWindow(), "Send an Instant Message")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
IMSEND% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And IMSEND% <> 0 Then Exit Do
Loop
Call AOLSetText(aoledit%, person)
Call AOLSetText(aolrich%, Message)
SendKeys "{TAB}"
End Function
Function AOLIsOnline()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
welcome% = FindChildByTitle(mdi%, "Welcome, ")
If welcome% = 0 Then
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function
Public Sub AOLLoadAol()
On Error Resume Next
Dim X%
X% = Shell("C:\aol30\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol30a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol30b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol25\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol25a\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
X% = Shell("C:\aol25b\waol.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Function AOLIMScan()
aolcl% = FindWindow("#32770", "America Online")
If aolcl% > 0 Then
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs OFF and can't recieve IMs"
End If
If aolcl% = 0 Then
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
im% = FindChildByTitle(mdi%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): Exit Do
If im% = 0 Then Exit Do
Loop
MsgBox "This person has their IMs ON and can recieve IMs."
End If
End Function

Sub AOLKeyword(Text)
Call RunMenuByString(AOLWindow(), "Keyword...")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
keyw% = FindChildByTitle(mdi%, "Keyword")
Kedit% = FindChildByClass(keyw%, "_AOL_Edit")
If Kedit% Then Exit Do
Loop

editsend% = SendMessageByString(Kedit%, WM_SETTEXT, 0, Text)
pausing = DoEvents()
Sending% = SendMessage(Kedit%, WM_CHAR, 13, 0)
pausing = DoEvents()
End Sub

Function AOLLastChatLine()
getpar% = AOLFindRoom()
child = FindChildByClass(getpar%, "_AOL_View")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

theview$ = TrimSpace$


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
AOLLastChatLine = lastline
End Function

Sub Mail_SendNew(person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(Mess%, WM_SETTEXT, 0, Message)
AOLIcon (icone%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub
Sub Mail_SendNew3(person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(Mess%, WM_SETTEXT, 0, Message)
AOLIcon (icone%)
HideWindow (mailwin%)
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub
Sub Mail_SendNew2(person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(Mess%, WM_SETTEXT, 0, Message)
X = MsgBox("Please Attch File And Send", vbCritical, "BuM Auto Tagger l3y GenghisX")
last:
End Sub



Sub AOLResetNewUser(SN As String, tru_sn As String, pth As String)
'creates a new sn
'example : Call AOLResetNewUser("NewSN", "CurrentSN", "C:\aol30\Organize")
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already on new user!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The screen name has to be at least 7 characters long :)"): Exit Sub
tru_sn = tru_sn + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(16384, Chr$(0))
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
m0226 = String$(16384, " ")
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 16384
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next
End Sub

Sub File_BirthDay(file$)
If Not File_IfileExists(file$) Then Exit Sub
MsgBox FileDateTime(file$)
NoFreeze% = DoEvents()
End Sub
Sub File_Copy(file$, DestFile$)
If Not File_IfileExists(file$) Then Exit Sub
FileCopy file$, DestFile$
End Sub
Sub File_Delete(file$)
If Not File_IfileExists(file$) Then Exit Sub
Kill file$
NoFreeze% = DoEvents()
End Sub



Sub File_DeleteDir(DirName$)
If Not File_IfDirectoryExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub
Sub File_DeleteDirectory(DirName$)
If Not File_IfDirectoryExists(DirName$) Then Exit Sub
RmDir DirName$
End Sub



Function File_IfDirectoryExists(TheDirectory)
Dim Check As Integer
On Error Resume Next
If right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(Dir$(TheDirectory))
If Err Or Check = 0 Then
    File_IfDirectoryExists = False
Else
    File_IfDirectoryExists = True
End If
End Function


Sub File_MakeDirectory(DirName$)
MkDir DirName$
End Sub
Function File_IfileExists(ByVal sFileName As String) As Integer
'Example: If Not File_ifileexists("win.com") then...
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        File_IfileExists = False
        Else
            File_IfileExists = True
    End If

End Function
Function File_LoadINI(look$, FileNamer$) As String
On Error GoTo Sla
Open FileNamer$ For Input As #1
Do While Not EOF(1)
    Input #1, CheckOut$
    If InStr(UCase$(CheckOut$), UCase$(look$)) Then
        Where = InStr(UCase$(CheckOut$), UCase$(look$))
        out$ = Mid$(CheckOut$, Where + Len(look$))
        File_LoadINI = out$
    End If
Loop
Sla:
Close #1
Resume nigger
nigger:
End Function
Sub File_OpenEXE(file$)
OpenEXE = Shell(file$, 1): NoFreeze% = DoEvents()
End Sub






Sub File_ReName(file$, NewName$)
Name file$ As NewName$
NoFreeze% = DoEvents()
End Sub
Sub FormFlash(frm As Form)
frm.Show
frm.BackColor = &H0&
Pause (".1")
frm.BackColor = &HFF&
Pause (".1")
frm.BackColor = &HFF0000
Pause (".1")
frm.BackColor = &HFF00&
Pause (".1")
frm.BackColor = &H8080FF
Pause (".1")
frm.BackColor = &HFFFF00
Pause (".1")
frm.BackColor = &H80FF&
Pause (".1")
frm.BackColor = &HC0C0C0
End Sub

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

GetListIndex = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function
Sub RemoveItemFromListbox(Lst As ListBox)
'this code works well in the double click part of your listbox
Jaguar% = Lst.ListIndex
Lst.RemoveItem (Jaguar%)
End Sub
Public Sub TransferListToTextBox(Lst As ListBox, Txt As textbox)
'This moves the individual highlighted part of a
'listbox to a textbox
Ind = Lst.ListIndex
daname$ = Lst.List(Ind)
Txt.Text = ""
Txt.Text = daname$
End Sub
Function AOLUpChat()
Do
    X% = DoEvents()
aolmod = FindWindow("_AOL_Modal", 0&)
killwin (aolmod)
Loop Until aolmod = 0
End Function
Sub Win95_StartButton()
wind% = FindWindow("Shell_TrayWnd", 0&)
btn% = FindChildByClass(wind%, "Button")
SendNow% = SendMessageByNum(btn%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(btn%, WM_LBUTTONUP, &HD, 0)
End Sub
Sub AOLHIDENMail(person, subject, Message)
Call RunMenuByString(AOLWindow(), "Compose Mail")
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
HideWindow mailwin%
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
Mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And Mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, person)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, subject)
a = SendMessageByString(Mess%, WM_SETTEXT, 0, Message)
AOLIcon (icone%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
mailwin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
AOLButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub


Sub AOLMainMenu()
Call RunMenu(2, 3)
End Sub

Function AOLRoomCount()
thechild% = AOLFindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")

getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOLRoomCount = getcount
End Function

Sub AOLSetText(win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub

Sub AOLSignOff()
aol% = FindWindow("AOL Frame25", vbNullString)
If aol% = 0 Then MsgBox "AOL client error: Please open Windows America Online before continuing.", 64, "Error: Windows America Online": Exit Sub
Call RunMenu(2, 0)

Exit Sub
'ignore since of new aol....
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
pfc% = FindChildByTitle(aol%, "Sign Off?")
If pfc% <> 0 Then
icon1% = FindChildByClass(pfc%, "_AOL_Icon")
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
icon1% = GetWindow(icon1%, 2)
ClickIcon% = SendMessage(icon1%, WM_LBUTTONDOWN, 0, 0&)
ClickIcon% = SendMessage(icon1%, WM_LBUTTONUP, 0, 0&)
Exit Do
End If
Loop

End Sub

Function AOLVersion()
aol% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(aol%)

submenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(submenu%, subitem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 2.5
End If
End Function

Function AOLWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = aol%
End Function



Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hWndTitle$, (hwndLength% + 1))

GetCaption = hWndTitle$
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Function GetWindowDir()
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
If right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function

Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub

Function Pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Function


Sub SendCharNum(win, chars)
e = SendMessageByNum(win, WM_CHAR, chars, 0)

End Sub

Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function

Sub SetPreference()
Call RunMenuByString(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
AOLIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

AOLButton (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub











Public Sub DisableCRTL_ALT_DEL()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub EnableCRTL_ALT_DEL()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub StayOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub

Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub


Sub UnHideWindow(hWnd)
un = ShowWindow(hWnd, SW_SHOW)
End Sub



Sub WaitForOk()
Do: DoEvents
aol% = FindWindow("#32770", "America Online")

If aol% Then
closeaol% = SendMessage(aol%, WM_CLOSE, 0, 0)
Exit Do
End If

aolw% = FindWindow("_AOL_Modal", vbNullString)

If aolw% Then
AOLButton (FindChildByTitle(aolw%, "OK"))
Exit Do
End If
Loop

End Sub

Sub WaitWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
topmdi% = GetWindow(mdi%, 5)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
topmdi2% = GetWindow(mdi%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop

End Sub


Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop

End Function


Function AOLKillChatRoom()
'This finds the chat room
Room = AOLFindRoom()
'This kills the chat room
Do
    killwin (Room)
    Loop Until Room = 0
End Function
Sub killwin(windo)
X = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub


Function CurrentDate()
X = "" & Month(Now) & "/" & Day(Now) & "/" & Year(Now) & ""
CurrentDate = X
End Function

Function CurrentTime()
X = "" & Hour(Now) & ":" & Minute(Now) & ":" & Second(Now) & ""
CurrentTime = X
End Function




Function AOLLamerSensor()
spcpos1 = InStr(1, AOLLastChatLine, ":")
SN1 = Left(AOLLastChatLine, spcpos1 - 1)
help = Len(AOLLastChatLine) - (spcpos1 + 1)
chat = right(AOLLastChatLine, help)
AOL_LamerSensor = SN1
End Function

Public Sub AOLBotCode()
'Candy Bot
'For use with AOL 3.0 for Windows 95.

'Items needed to create this bot:
'1 Timer, 2 command buttons
'* The timer needs to be set so enabled = false and it has an interval of 1
'In the first command button put:
'Timer1.enabled = true
'In the second command button put
'Timer1.enabled = false
'Those are the start and stop buttons.
'Now put this code in the timer1
'On Error Resume Next
'Dim last As String
'Dim name As String
'Dim a As String
'Dim n As Integer
'Dim x As Integer
'DoEvents
'a = AOLLastChatLine
'last = Len(a)
'For x = 1 To last
'name = Mid(a, x, 1)
'final = final & name
'If name = ":" Then Exit For
'Next x
'final = Left(final, Len(final) - 1)
'If InStr(a, "/candy") Then
'Randomize
'rand = Int((Rnd * 5) + 1)
'If rand = 1 Then Call AOLChatSend("º°˜¨˜°ºº°˜¨˜°º " & final & " you get some gum")
'If rand = 2 Then Call AOLChatSend("º°˜¨˜°ºº°˜¨˜°º " & final & " you get some nerds")
'If rand = 3 Then Call AOLChatSend("º°˜¨˜°ºº°˜¨˜°º " & final & " you get a Snickers")
'If rand = 4 Then Call AOLChatSend("º°˜¨˜°ºº°˜¨˜°º " & final & " you get a Butterfinger ")
'If rand = 5 Then Call AOLChatSend("º°˜¨˜°ºº°˜¨˜°º " & final & " you get more nerds than the last guy did.")
'Call Pause(0.6)
'End If

End Sub

Sub AOLDecodeSN(Text)
Dim trans As String
trans = Text
X = Text_Decode(trans)
MsgBox X 'change this to whatever you need
' you might want to send this to chat or whatever, but im just msgboxing
' it for now for the sake of an example.
End Sub


