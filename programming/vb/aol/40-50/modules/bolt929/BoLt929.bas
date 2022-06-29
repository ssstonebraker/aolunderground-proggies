Attribute VB_Name = "BoLt929"
'ßØL†929 version 2.5
'You can also use the public const on the bottom for colors instead of typing the Anyonning color code
'for the fadeforms like fadeform or iceform u always have to have "CALL" before it
'This is for Mostly AOL 4.0
'Created by:Uyhs
'Modified by:Uyhs
'E-mail:Oddish0923@aol.com
'Website:Http://clik.to/uyhs
'I must thank Mokefade for the great Sub's, and Function's
'Aim Screen Name:Bolt50432
'If you want to fade lets say red,blue,navy then type:
'''''''Bolt=fadebycolor3(red,blue,navy,THETEXT,true)
'''''''sendchat "<b>"+bolt
'I made that wavy by haveing ht elast parameter say true
'if you dont want it make the last parameter say false
'The Color Subs are suppose to make it easier instead of entering the color code
'Now that AOL Has Totally CRAPPED UP THE CHATROOMS the fadetext's wont work for aol
'I ll still kepp'em up there anyway.I havent done much so far
'i put up a dragform thing,I got a 2 aim subs cuz i wanted to put them there
'cs is 1 and clearaimchat is the other
'cs is a chatsend for aim and clearaimchat is pretty self explanatory
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function movewindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, LPRect As Rect) As Long
Declare Function SetRect Lib "user32" (LPRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Declare Function setparent Lib "user32" Alias "SetParent" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function findwindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function showwindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function Getmenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function Gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function getwindowtext Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function getwindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function enablewindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_SYSCOMMAND = &H112
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
Public Const WM_Close = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GetTextLength = &HE
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
Public Const LB_Setcursel = &H186
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
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

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
Public Const Red = &HFF&
Public Const Green = &HFF00&
Public Const Blue = &HFF0000
Public Const yellow = &HFFFF&
Public Const white = &HFFFFFF
Public Const black = &H0&
Public Const purple = &HFF00FF
Public Const grey = &HC0C0C0
Public Const pink = &HFF80FF
Public Const TURQUOISE = &HC0C000
Public Const SEAGREEN = &H80FF80
Public Const LBLUE = &HFFFFC0
Public Const LGREEN = &HFF00&
Public Const brown = &H4080&
Public Const DGREEN = &H8000&
Public Const NAVY = &H800000
Public Const GOLD = &H8080&
Public Const BLUEPRPL = &HDA2C68
Public Const YELGRN = &H5EF7B6
Public Const MAGENTA = &H640DE8
Public Const MAROON = &H291F76
Public Const orange = &H80FF&
Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type
 Const STANDARD_RIGHTS_REQUIRED = &HF0000

Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type pointapi
   X As Long
   Y As Long
End Type
Sub WAVPlay(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndplaysound(SoundName$, wFlags%)

End Sub
Sub FormOnTop(theform As Form)
SetWinOnTop = SetWindowPos(theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Function RGBtoHEX(RGB)
'heh, I didnt make this one...
    a$ = Hex(RGB)
    B% = Len(a$)
    If B% = 5 Then a$ = "0" & a$
    If B% = 4 Then a$ = "00" & a$
    If B% = 3 Then a$ = "000" & a$
    If B% = 2 Then a$ = "0000" & a$
    If B% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, Wavy As Boolean)

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

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, Wavy)

End Function
Sub Scrambler()
'This is how you make a scrambler
'Make 2 textboxes,1 command button, and a timer
'in the command button(Name Scramble) put this code in it
'''chatsend "`·.….·>Scrambled Word" + scrambletext(text1.text)
'''chatsend "`·.….·>Hint" + text2.text
'''timer1.enabled=true
'Then in the timer put this code in(With interval set to 1)
'''If lcase(chatlastline)like lcase(text1.text) then
'''chatsend "`·.….·>" + snlastchatline + " got it!!"
'''timer1.enabled=false
'''end if
'Text1 should be the word you scramble
'text2 should be the hint
End Sub
Sub AnyonningError()
'This can be used in your program when someone does something they dont wanna do
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
Beep
Freeze 0.4
End Sub



Function SNLastChatLine()
chattext$ = ChatLastLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        sn = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = sn
End Function
Sub Chat_Anyonner()
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "WHo's doing that???"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "STOP!!!"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "STOP NOW!!"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "SSSSTTTOOOOOPPPPP!!!"
ChatSend "/////////"
ChatSend "/////////"
ChatSend "/////////"

End Sub
Sub Sound_Anyonner()
ChatSend "{S im"
Freeze 5#
ChatSend "{S gotmail"
Freeze 5#
ChatSend "{S filedone"
Freeze 5#
ChatSend "{S mailbeep"
Freeze 5#
ChatSend "{S welcome"
Freeze 5#
ChatSend "{S goodbye"
ChatSend "{S im"
Freeze 5#
ChatSend "{S gotmail"
Freeze 5#
ChatSend "{S filedone"
Freeze 5#
ChatSend "{S mailbeep"
Freeze 5#
ChatSend "{S welcome"
Freeze 5#
ChatSend "{S goodbye"
ChatSend "{S im"
Freeze 5#
ChatSend "{S gotmail"
Freeze 5#
ChatSend "{S filedone"
Freeze 5#
ChatSend "{S mailbeep"
Freeze 5#
ChatSend "{S welcome"
Freeze 5#
ChatSend "{S goodbye"
End Sub

Sub EliteText(word$)
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
SendChat (Made$)
End Sub
Sub FormCenter(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Function ChatLastLineWithSN()
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
Function ChatLastLine()
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function
Public Sub DragForm(theform As Form)
    ReleaseCapture
SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Sub BlackBrownForm(frm As Form)
Call FadeForm(frm, black, brown)
End Sub
Sub CS(SayWhat As String)
    Dim ChatWindow As Long, THing As Long, Thing2 As Long
    Dim SetChatText As Long, Buttin As Long, Buttin2 As Long, Buttin3 As Long
    Dim SendButtin As Long, Click As Long
    
    ChatWindow& = findwindow("AIM_ChatWnd", vbNullString)
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


Sub Subtract(l As Label, Huh As String)
l.Caption = Val(l) - Huh
End Sub
Sub Add(l As Label, Huh As String)
l.Caption = Val(l) + Huh
End Sub




Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen$ = Len(TheText)
    For I = 1 To textlen$
        TextDone$ = Left(TheText, I)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen$ * I) + B1, ((G2 - G1) / textlen$ * I) + G1, ((R2 - R1) / textlen$ * I) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
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
    Next I
    FadeTwoColor = Faded$
End Function


Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For I = 1 To textlen%
        TextDone$ = Left(part1$, I)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * I) + B1, ((G2 - G1) / textlen% * I) + G1, ((R2 - R1) / textlen% * I) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
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
    Next I
    'part2
    textlen% = Len(part2$)
    For I = 1 To textlen%
        TextDone$ = Left(part2$, I)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * I) + B2, ((G3 - G2) / textlen% * I) + G2, ((R3 - R2) / textlen% * I) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
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
    Next I
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function
Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function

Function FindChatRoom()
AOL% = findwindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function
Function FindChildByClass(parentw, childhand)
firs% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = getwindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = getwindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = getwindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function
Sub WaitOk()
Do
DoEvents
okw = findwindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)


End Sub
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = getwindowtext(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function
Function FindChildByTitle(parentw, childhand)
firs% = getwindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = getwindow(parentw, GW_CHILD)

While firs%
firss% = getwindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = getwindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs%
FindChildByTitle = Room%
End Function


Sub FadeForm(FormX As Form, Colr1, Colr2)
    B1 = GetRGB(Colr1).Blue
    G1 = GetRGB(Colr1).Green
    R1 = GetRGB(Colr1).Red
    B2 = GetRGB(Colr2).Blue
    G2 = GetRGB(Colr2).Green
    R2 = GetRGB(Colr2).Red
   
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
Function FadeByColor2(Colr1, Colr2, TheText$, Wavy As Boolean)
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, Wavy)

End Function
Sub ShowTime(TheTime As Label)
Do
TheTime = time
Freeze 1#
Loop
End Sub

Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndplaysound(SoundName$, wFlags%)

End Sub


Public Sub FadePicture(Pic As PictureBox, TopColor&, BottomColor&)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim I&, j&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim TopColorRed%, TopColorGreen%, TopColorBlue%
Dim BottomColorRed%, BottomColorGreen%, BottomColorBlue%

SaveScale = Pic.ScaleMode
SaveStyle = Pic.DrawStyle
SaveRedraw = Pic.AutoRedraw

Pic.ScaleMode = 3
TopColorRed = TopColor And 255
TopColorGreen = (TopColor And 65280) / 256
TopColorBlue = (TopColor And 16711680) / 65536
BottomColorRed = BottomColor And 255
BottomColorGreen = (BottomColor And 65280) / 256
BottomColorBlue = (BottomColor And 16711680) / 65536
    aRed = TopColorRed
    aGreen = TopColorGreen
    aBlue = TopColorBlue
    pixels = Pic.ScaleWidth
    
    If pixels <= 0 Then Exit Sub
    
    ColorDifRed = (BottomColorRed - TopColorRed)
    ColorDifGreen = (BottomColorGreen - TopColorGreen)
    ColorDifBlue = (BottomColorBlue - TopColorBlue)
    
    RedDelta = ColorDifRed / pixels
    GreenDelta = ColorDifGreen / pixels
    BlueDelta = ColorDifBlue / pixels
    
    Pic.DrawStyle = 5
    Pic.AutoRedraw = True
    
    
 For Y = 0 To pixels
        aRed = aRed + RedDelta
        If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
        If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
        If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
        If ThisColor > -1 Then
        
        Pic.Line (Y - 2, -2)-(Y - 2, Pic.Height + 2), ThisColor, BF
        End If
    Next Y
  

Pic.ScaleMode = SaveScale
Pic.DrawStyle = SaveStyle
Pic.AutoRedraw = SaveRedraw
End Sub
Public Sub FadePicture3(Pic As PictureBox, TopColor&, MidColor&, BottomColor&)
Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim I&, j&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim TopColorRed%, TopColorGreen%, TopColorBlue%
Dim BottomColorRed%, BottomColorGreen%, BottomColorBlue%
Dim MidColorRed%, Midcolorgreen%, middcolorblue%
SaveScale = Pic.ScaleMode
SaveStyle = Pic.DrawStyle
SaveRedraw = Pic.AutoRedraw

Pic.ScaleMode = 3
TopColorRed = TopColor And 255
TopColorGreen = (TopColor And 65280) / 256
TopColorBlue = (TopColor And 16711680) / 65536
BottomColorRed = BottomColor And 255
BottomColorGreen = (BottomColor And 65280) / 256
BottomColorBlue = (BottomColor And 16711680) / 65536
MidColorRed = MidColor And 255
Midcolorgreen = (MidColor And 65280) / 256
midcolorblue = (BottomColor And 16711680) / 65536

    aRed = TopColorRed
    aGreen = TopColorGreen
    aBlue = TopColorBlue
    pixels = Pic.ScaleWidth
    
    If pixels <= 0 Then Exit Sub
    
    ColorDifRed = (BottomColorRed - MidColorRed - TopColorRed)
    ColorDifGreen = (BottomColorGreen - Midcolorgreen - TopColorGreen)
    ColorDifBlue = (BottomColorBlue - midcolorblue - TopColorBlue)
    
    RedDelta = ColorDifRed / pixels
    GreenDelta = ColorDifGreen / pixels
    BlueDelta = ColorDifBlue / pixels
    
    Pic.DrawStyle = 5
    Pic.AutoRedraw = True
    
    
 For Y = 0 To pixels
        aRed = aRed + RedDelta
        If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
        If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
        If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
        If ThisColor > -1 Then
        
        Pic.Line (Y - 2, -2)-(Y - 2, Pic.Height + 2), ThisColor, BF
        End If
    Next Y
  

Pic.ScaleMode = SaveScale
Pic.DrawStyle = SaveStyle
Pic.AutoRedraw = SaveRedraw
End Sub

Sub Freeze(interval) 'AKA Pause or Timeout
Dim time
time = Timer
Do While Timer - time < Val(interval)
DoEvents
Loop
End Sub
Sub CoolFormBegining(form1 As Form, LeftPostion As String, HeightPostion As String, WidthPostion As String, THETIMER As Timer)
'put this in a timer and in the form put thetimer.enabled=true
' make the timers interval 1

form1.Enabled = False
form1.Left = 510
form1.Left = 550
form1.Height = 720
form1.Left = 600
Call SunriseForm(form1)
Freeze 0.1
form1.Left = 650
form1.Left = 700
form1.Height = 750
form1.Left = 750
Call IceForm(form1)
Freeze 0.1
form1.Left = 800
form1.Left = 850
form1.Height = 800
form1.Left = 900
Call NeonForm(form1)
Freeze 0.1
form1.Left = 950
form1.Left = 1000
form1.Height = 850
form1.Left = 1050
Call BlackWhiteForm(form1)
Freeze 0.1
form1.Left = 1100
form1.Left = 1150
form1.Height = 900
form1.Left = 1200
Call FadeForm(form1, GOLD, black)
Freeze 0.1
form1.Left = 1250
form1.Left = 1300
form1.Height = 950
form1.Left = 1350
Call FadeForm(form1, SEAGREEN, black)
Freeze 0.1
form1.Left = 510
form1.Left = 550
form1.Height = 720
form1.Left = 600
Call SunriseForm(form1)
Freeze 0.1
form1.Left = 650
form1.Left = 700
form1.Height = 750
form1.Left = 750
Call IceForm(form1)
Freeze 0.1
form1.Left = 800
form1.Left = 850
form1.Height = 800
form1.Left = 900
Call NeonForm(form1)
Freeze 0.1
form1.Left = 950
form1.Left = 1000
form1.Height = 850
form1.Left = 1050
Call BlackWhiteForm(form1)
Freeze 0.1
form1.Left = 1100
form1.Left = 1150
form1.Height = 900
form1.Left = 1200
Call FadeForm(form1, GOLD, black)
Freeze 0.1
form1.Left = 1250
form1.Left = 1300
form1.Height = 950
form1.Left = 1350
Call FadeForm(form1, SEAGREEN, black)
Freeze 0.1
form1.Left = 510
form1.Left = 550
form1.Height = 720
form1.Left = 600
Call SunriseForm(form1)
Freeze 0.1
form1.Left = 650
form1.Left = 700
form1.Height = 750
form1.Left = 750
Call IceForm(form1)
Freeze 0.1
form1.Left = 800
form1.Left = 850
form1.Height = 800
form1.Left = 900
Call NeonForm(form1)
Freeze 0.1
form1.Left = 950
form1.Left = 1000
form1.Height = 850
form1.Left = 1050
Call BlackWhiteForm(form1)
Freeze 0.1
form1.Left = 1100
form1.Left = 1150
form1.Height = 900
form1.Left = 1200
Call FadeForm(form1, GOLD, black)
Freeze 0.1
form1.Left = 1250
form1.Left = 1300
form1.Height = 950
form1.Left = 1350
Call FadeForm(form1, SEAGREEN, black)
Freeze 0.1
form1.Left = 510
form1.Left = 550
form1.Height = 720
form1.Left = 600
Call SunriseForm(form1)
Freeze 0.1
form1.Left = 650
form1.Left = 700
form1.Height = 750
form1.Left = 750
Call IceForm(form1)
Freeze 0.1
form1.Left = 800
form1.Left = 850
form1.Height = 800
form1.Left = 900
Call NeonForm(form1)
Freeze 0.1
form1.Left = 950
form1.Left = 1000
form1.Height = 850
form1.Left = 1050
Call BlackWhiteForm(form1)
Freeze 0.1
form1.Left = 1100
form1.Left = 1150
form1.Height = 900
form1.Left = 1200
Call FadeForm(form1, GOLD, black)
Freeze 0.1
form1.Left = 1250
form1.Left = 1300
form1.Height = 950
form1.Left = 1350
Call FadeForm(form1, SEAGREEN, black)
Freeze 0.1
form1.Left = 510
form1.Left = 550
form1.Height = 720
form1.Left = 600
Call SunriseForm(form1)
Freeze 0.1
form1.Left = 650
form1.Left = 700
form1.Height = 750
form1.Left = 750
Call IceForm(form1)
Freeze 0.1
form1.Left = 800
form1.Left = 850
form1.Height = 800
form1.Left = 900
Call NeonForm(form1)
Freeze 0.1
form1.Left = 950
form1.Left = 1000
form1.Height = 850
form1.Left = 1050
Call BlackWhiteForm(form1)
Freeze 0.1
form1.Left = 1100
form1.Left = 1150
form1.Height = 900
form1.Left = 1200
Call FadeForm(form1, GOLD, black)
Freeze 0.1
form1.Left = 1250
form1.Left = 1300
form1.Height = 950
form1.Left = 1350
Call FadeForm(form1, SEAGREEN, black)
Freeze 0.1
form1.Left = LastPostion
form1.Height = HeightPostion
form1.Width = WidthPostion
form1.Enabled = True
THETIMER.Enabled = False
End Sub
Sub RedBlack(txt As String)
bolt = FadeByColor2(Red, black, txt, False)
ChatSend "<b>" + bolt
End Sub
Sub GreenBlack(txt As String)
bolt = FadeByColor2(Green, black, txt, False)
ChatSend "<b>" + bolt
End Sub
Sub GoldBlack(txt As String)
bolt = FadeByColor2(GOLD, black, txt, False)
ChatSend "<b>" + bolt
End Sub


Sub IceForm(theform As Form)
Call FadeForm(theform, LBLUE, NAVY)
End Sub
Sub SunriseForm(theform As Form)
Call FadeForm(theform, yellow, Red)
End Sub
Sub NeonForm(theform As Form)
Call FadeForm(theform, LGREEN, yellow)
End Sub
Sub RedGreenText(TheText As String)
bolt = FadeByColor2(FADE_RED, FADE_GREEN, TheText, False)
ChatSend "<b>" + bolt
End Sub
Sub RedBlackText(TheText As String)
bolt = FadeByColor2(FADE_RED, FADE_BLACK, TheText, False)
ChatSend "<b>" + bolt
End Sub
Sub IceFadeText(TheText As String)
bolt = FadeByColor3(LBLUE, TURQUOISE, Blue, TheText, False)
ChatSend "<B>" + bolt
End Sub
Sub Attention(message As String)
ChatSend "<B><I><U>—·•·—»»Attetion««—·•·—</B></I></U>"
ChatSend (message)
ChatSend "<B><I><U>—·•·—»»Attetion««—·•·—</B></I></U>"

End Sub
Function ScrambleText(TheText)
'scrambles words man!!!!
findlastspace = Mid(TheText, Len(TheText), 1)
If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If
For scrambling = 1 To Len(TheText)
thechar$ = Mid(TheText, scrambling, 1)
Char$ = Char$ & thechar$
If thechar$ = " " Then
chars$ = Mid(Char$, 1, Len(Char$) - 1)
firstchar$ = Mid(chars$, 1, 1)
On Error GoTo Error
lastchar$ = Mid(chars$, Len(chars$), 1)
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo stuff
Error:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo stuffs
stuff:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "
stuffs:
Char$ = ""
backchar$ = ""
End If
Next scrambling
ScrambleText = scrambled$
Exit Function

End Function
Sub LGreenBlack(txt As String)
bolt = FadeByColor2(LGREEN, black, txt, False)
ChatSend "<b>" + bolt
End Sub
Sub ClearAIMChat()
    Dim ChatWindow As Long, BorderThing As Long

    ChatWindow& = findwindow("AIM_ChatWnd", vbNullString)
    BorderThing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Call SendMessageByString(BorderThing&, WM_SETTEXT, 0, "")
End Sub
Sub ChatSend(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)
AORich% = getwindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call sendmessagebynum(AORich%, WM_CHAR, 13, 0)
End Sub
Sub BlackWhiteForm(theform As Form)
Call FadeForm(theform, black, white)
End Sub

