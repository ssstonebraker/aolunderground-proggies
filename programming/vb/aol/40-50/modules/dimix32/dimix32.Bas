Attribute VB_Name = "DÓmÓx≥≤"
'DÓmÓx≥≤ By MÂcËX
'Email MaceX@macex.zzn.com
'Web: macex.cjb.net
'Credit: To l3usta me and him started it and i took over and finished it
'also to ChiChis i used his api spy and any simularities it due to the
'same api spy
'ths bas has only aol4 & 5 stuff
'next will have others
'                 Greetz:
'          Oreo69,Progking,mmR,ecko
'          Kalo,NcR,SiC,l3usta,rca
'          YaZ,d0pey,Edge,Seed,Atom
'          ccz,jester,HaK,File,MBF
'          Warrior,UqG,Linux,natron
'          alt,Dos,Billy,TiTo,masta
'          esc, Monk-e-God,firex
'          saint, spore, bry, laf,sg2
'          mistical,bofen,chichis, jet
          'naz,sjm,seed & if i forgot u i'm sorry


'Functions
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd&, ByVal lpClassName$, ByVal nMaxCount&)
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextB Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' Contstants
Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const Key_Home = &H24
Public Const Key_Right = &H27
Public Const Key_Control = &H11
Public Const Key_DELETE = &H2E
Public Const Key_Down = &H28
Public Const Key_Left = &H25
Public Const Key_Enter = &HD
Public Const Key_Space = &H20
Public Const Key_Tab = &H9



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

Const HWND_TOP = 0
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

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
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

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

'X = Mp3Play1.Authorize("ieet kryo", "629436287")
        
        
       

Function aol_ascii()
aol_ascii = "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix ∑ï macex ∑ï load"
End Function

Sub form_ontop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Sub aol_back()
'clicks back button on aol
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLIcon As Long, i As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 15
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub
Sub chat_setfont(font As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLCombobox As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
SendMessageByString AOLCombobox, &HC, 0&, font
End Sub

Sub aol_home()
'clicks home button
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLIcon As Long, i As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 19
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub aol_pfc()
popupicon "4", "P"
End Sub

Public Sub aol_search(Werds As String)
  'searches for a string in the aol search engine
keyword "http://aolsearch.aol.com/dirsearch.adp?query=" & LCase$(Werds)
End Sub
Sub aol_shell()
Shell "C:\America Online 4.0\waol.exe"
Shell "C:\America Online 5.0\waol.exe"
End Sub

Public Sub aol_uninstall()
  'runs the uninstall aol program
  file_shell ("c:\windows\aolunins.exe")
End Sub
Sub aol_refresh()
'clicks refesh button
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLIcon As Long, i As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)

AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 18
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub aol_stop()
'clicks stop button
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLIcon As Long, i As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)

AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 17
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub aol_forward()
'clicks forward button
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLIcon As Long, i As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 16
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub channel_computin()
popupicon "8", "U"
End Sub

Sub channel_entertainment()
popupicon "8", "E"
End Sub

Sub channel_familes()
popupicon "8", "F"
End Sub
Sub channel_games()
popupicon "8", "G"
End Sub


Sub channel_health()
popupicon "8", "H"
End Sub
Sub channel_influence()
popupicon "8", "I"

End Sub
Sub channel_interests()
popupicon "8", "T"
End Sub
Sub channel_international()
popupicon "8", "l"
End Sub

Sub channel_kidsonly()
popupicon "8", "K"
End Sub
Sub channel_lifestyles()
popupicon "8", "y"

End Sub

Sub channel_local()
popupicon "8", "c"
End Sub

Sub channel_news()
popupicon "8", "N"
End Sub
Sub channel_personalfinance()
popupicon "8", "P"

End Sub

Sub channel_researchnlearn()
popupicon "8", "R"
End Sub
Sub channel_shopping()
popupicon "8", "p"
End Sub
Sub channel_sports()
popupicon "8", "s"
End Sub

Sub channel_travel()
popupicon "8", "v"
End Sub

Sub channel_workplace()
popupicon "8", "W"
End Sub
Function chat_lastsn()
'gets last sn

ChatText$ = chat_lastline
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
chat_lastsn = SN
End Function
Sub chat_now()
popupicon "9", "c"

End Sub

Sub desktop_close()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long

Progman = FindWindow("Progman", vbNullString)
SHELLDLLDefView = FindWindowEx(Progman, 0&, "SHELLDLL_DefView", vbNullString)

InternetExplorerServer = FindWindowEx(SHELLDLLDefView, 0&, "Internet Explorer_Server", vbNullString)
SendMessage InternetExplorerServer, &H10, 0&, 0&
End Sub

Sub desktop_disable()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long

Progman = FindWindow("Progman", vbNullString)
SHELLDLLDefView = FindWindowEx(Progman, 0&, "SHELLDLL_DefView", vbNullString)

InternetExplorerServer = FindWindowEx(SHELLDLLDefView, 0&, "Internet Explorer_Server", vbNullString)
EnableWindow InternetExplorerServer, 0
End Sub
Sub desktop_enable()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long

Progman = FindWindow("Progman", vbNullString)
SHELLDLLDefView = FindWindowEx(Progman, 0&, "SHELLDLL_DefView", vbNullString)

InternetExplorerServer = FindWindowEx(SHELLDLLDefView, 0&, "Internet Explorer_Server", vbNullString)
EnableWindow InternetExplorerServer, 1
End Sub

Sub desktop_hide()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long

Progman = FindWindow("Progman", vbNullString)
SHELLDLLDefView = FindWindowEx(Progman, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer = FindWindowEx(SHELLDLLDefView, 0&, "Internet Explorer_Server", vbNullString)
ShowWindow InternetExplorerServer, 0
End Sub


Sub desktop_show()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long

Progman = FindWindow("Progman", vbNullString)
SHELLDLLDefView = FindWindowEx(Progman, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer = FindWindowEx(SHELLDLLDefView, 0&, "Internet Explorer_Server", vbNullString)
ShowWindow InternetExplorerServer, 1
End Sub
Sub file_shell(File As String)
Shell File
End Sub

Function mail_flashcount()
Dim Flash As Long, AoLTree As Long
Flash = find_flashmailbox
If Flash = 0 Then
MsgBox "You Must Open Your Flash Mail", vbError, "Dimix Error"
Exit Function
End If
AoLTree = FindWindowEx(Flash, 0&, "_AOL_Tree", vbNullString)
mail_flashcount = SendMessage(AoLTree, LB_GETCOUNT, 0, 0)
End Function

Public Sub popupicon2(IconNumber As Long, Character As String, Character2 As String)
    ' same as other but goew thought 2 menus
    Dim Message1 As Long, Message2 As Long, AOLFrame As Long
    Dim AOLToolbar As Long, Toolbar As Long, AOLIcon As Long
    Dim NextOfClass As Long, AscCharacter As Long, AscCharacter2 As Long
    Message1& = FindWindow("#32768", vbNullString)
    AOLFrame& = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AOLToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    For NextOfClass& = 1 To IconNumber&
        AOLIcon& = GetWindow(AOLIcon&, 2)
    Next NextOfClass&
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Message2& = FindWindow("#32768", vbNullString)
    Loop Until Message2& <> Message1&
    AscCharacter& = Asc(Character$)
    AscCharacter2& = Asc(Character2$)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter2&, 0&)
End Sub

Public Function getchattext() As String
    Dim Room As Long, Chat As Long, ChatText As String
    Room = find_chatroom
    Chat = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    ChatText = gettext(Chat)
    getchattext = ChatText
End Function
Function chat_lastline()
'gets last sn + msg
ChatText = getchattext
For FindChar = 1 To Len(ChatText)
thechar = Mid(ChatText, FindChar, 1)
thechars = thechars & thechar
If thechar = Chr(13) Then
TheChatText = Mid(thechars, 1, Len(thechars) - 1)
thechars = ""
End If
Next FindChar
lastlen = Val(FindChar) - Len(thechars)
lastline = Mid(ChatText, lastlen, Len(thechars))
chat_lastline = lastline
End Function
Function chat_lastlinemsg()
'this gets just the last msg
chatline$ = chat_lastline
If chatline$ = "" Then Exit Function
ChatTrim$ = Left$(chatline$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
ScreenName$ = SN
ChatTrimNum = Len(ScreenName$)
ChatTrim$ = Mid$(chatline$, ChatTrimNum + 4, Len(chatline$) - Len(ScreenName$))
chat_lastlinemsg = ChatTrim$
End Function
Sub chat_bold()
'clicks bold button
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub


Sub chat_italic()
'clicks italic button
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub chat_underline()
'clicks underline button
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub


Function getclass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)
getclass = buffer$
End Function

Function gettext(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
getstring = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
gettext = TrimSpace$
End Function
Function getchatname() As String
Dim Room As Long
    Room = FindChatRoom
    getchatname = gettext(Room)
End Function
Function getcaption(hWnd As Long) As String
    Dim Length As Long
    Dim Title As String
    Length = GetWindowTextLength(hWnd)
    Title = String(Length, 0)
    GetWindowText hWnd, Title, (Length + 1)
    getcaption = Title
End Function
Sub aol_show()
Dim AOLFrame As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
ShowWindow AOLFrame, 1
End Sub

Sub chat_greetz(Greetz)
'This is to make it were you can send all of your
'greets on one chat line
For i = 1 To 1
a = a + Greetz
Next
chatsend (".<p=" & a)
End Sub

Function chat_hacker(Strin As String)
Let inptxt$ = Strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = "†"
Let NewSent$ = NewSent$ + NextChr$
Loop
chatsend (NewSent$)
End Function
Function chat_lag(text As String)
 
g$ = thetext
a = Len(g$)
For W = 1 To a Step 4
    r$ = Mid$(g$, W, 1)
    u$ = Mid$(g$, W + 1, 1)
    S$ = Mid$(g$, W + 2, 1)
    T$ = Mid$(g$, W + 3, 1)
    p$ = p$ & "</html></html><html><html>" & r$ & "</html></html><html><font color=#0000ff><html>" & u$ & "</html></html><html><html>" & S$ & "</html></html><html><font color=#ff0000><html></html><html></html><html>" & T$
Next W
chatsend (p$)
  
End Function


Function chat_lag2(text As String)
  
g$ = thetext
a = Len(g$)
For W = 1 To a Step 4
    r$ = Mid$(g$, W, 1)
    u$ = Mid$(g$, W + 1, 1)
    S$ = Mid$(g$, W + 2, 1)
    T$ = Mid$(g$, W + 3, 1)
    p$ = p$ & "</html></html><html><html><html>" & r$ & "</html></html><html><html><html><html>" & u$ & "</html><html></html><html><html>" & S$ & "</html><html></html><html><html><html></html><html><html></html><html></html><html>" & T$
Next W
chatsend (p$)
  
End Function

Function chat_link(Link, txt)
chatsend ("< a href=" + Link + ">" + txt + "</a>")
End Function
Function chat_spaced(text As String) As String
Dim NextChr As String, inptxt As String, lenth As Integer
Dim NumSpc As Integer, NewSent As String, Dotz As String
    Let inptxt$ = text
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + " "
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    chatsend NewSent$
   End Function
Sub chat_fakeoh()
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & " "
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "MaceX Ownz AoL" & Chr(4) & ""
timeout (0.5)
chatsend "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "ââââââââââââ" & Chr(4) & "Dimix Ownz AoL" & Chr(4) & ""
End Sub


Sub chat_spiralscroll(txt As TextBox)
'Spiral Scroller
X = txt.text
macexx:
Dim MYLEN As Integer
MyString = txt.text
MYLEN = Len(MyString)
MYSTR = Mid(MyString, 2, MYLEN) + Mid(MyString, 1, 1)
txt.text = MYSTR
timeout 1
chatsend (txt)
If txt.text = X Then
Exit Sub
End If
GoTo macexx
End Sub



Function chat_wavy(Wat As String)
    For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<sup>" & i$ & "</sup>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$
 
    Next X
       chatsend Strin$
        
End Function

Function chat_wavy2(Wat As String)
 'waves a little different
    For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<sub>" & i$ & "</sub>" & i2$ & "<sup>" & i3$ & "</sup>" & i4$
 
    Next X
        chatsend Strin$
     
End Function
Sub ClickIcon(Button%)
Call SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
Call SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub

Function clicklist(hWnd)
'clicks a list
Call SendMessageByNum(hWnd, &H203, 0, 0&)
End Function

Sub kill_glyph()
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLGlyph As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)

AOLGlyph = FindWindowEx(AOLToolbar, 0&, "_AOL_Glyph", vbNullString)
SendMessage AOLGlyph, &H10, 0&, 0&
End Sub


Public Function list_count(listbox As Long) As Long
    ListCount& = SendMessageLong(listbox&, LB_GETCOUNT, 0&, 0&)
End Function

Public Function list_gettext(listbox As Long, Index As Long) As String
    Dim ListText As String * 256
    Call SendMessageByString(listbox&, LB_GETTEXT, Index&, ListText$)
    ListGetText$ = ListText$
End Function

Public Function getsnfromim() As String
    Dim imwin As Long
    Dim thecaption As String
    imwin = FindIMSent
    thecaption = getcaption(imwin)
    getsnfromim = Mid(thecaption, InStr(thecaption, ":") + 2)
End Function

Public Sub list_killdups(listbox As listbox)
    Dim FirstCount As Long, SecondCount As Long
    On Error Resume Next
    For FirstCount& = 0& To listbox.ListCount - 1
        For SecondCount& = 0& To listbox.ListCount - 1
            If LCase(listbox.List(FirstCount&)) Like LCase(listbox.List(SecondCount&)) And FirstCount& <> SecondCount& Then
                listbox.RemoveItem SecondCount&
            End If
        Next SecondCount&
    Next FirstCount&
End Sub

Public Sub list_load(Directory As String, thelist As listbox)
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        thelist.AddItem MyString$
    Wend
    Close #1
End Sub
Public Sub list_removeselected(listbox As listbox)
    Dim ListCount As Long
    ListCount& = listbox.ListCount
    Do While ListCount& > 0&
       ListCount& = ListCount& - 1
       If listbox.Selected(ListCount&) = True Then
          listbox.RemoveItem (ListCount&)
       End If
    Loop
End Sub
Public Sub list_save(Directory As String, thelist As listbox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.List(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Function list_search(listbox As listbox, SearchString As String) As Boolean
    Dim search As Long
    On Error Resume Next
    For search& = 0 To list_count(listbox.hWnd) - 1
        If listbox.List(search&) = SearchString$ Then
            ListSearch = True
            Exit Function
        End If
    Next search&
End Function
Function list_search2(lst As listbox, txt As String) As Integer
'search for a list

Dim X%
For X% = 0 To lst.ListCount - 1
If InStr(UCase(lst.List(X%)), UCase((txt))) Then
ListSearch2 = X%
Exit Function
End If
Next X%
ListSearch2 = -1
End Function

Sub startbutton_kill()
Dim ShellTrayWnd As Long, Button As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
SendMessage Button, &H10, 0&, 0&
End Sub

Sub startbutton_enable()
Dim ShellTrayWnd As Long, Button As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
EnableWindow Button, 1
End Sub

Sub startbutton_disable()
Dim ShellTrayWnd As Long, Button As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
EnableWindow Button, 0
End Sub
Sub startbutton_show()
Dim ShellTrayWnd As Long, Button As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
ShowWindow Button, 1
End Sub
Sub startbutton_hide()
Dim ShellTrayWnd As Long, Button As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
Button = FindWindowEx(ShellTrayWnd, 0&, "Button", vbNullString)
ShowWindow Button, 0
End Sub
Sub systray_enable()
Dim ShellTrayWnd As Long, TrayNotifyWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
TrayNotifyWnd = FindWindowEx(ShellTrayWnd, 0&, "TrayNotifyWnd", vbNullString)
EnableWindow TrayNotifyWnd, 1
End Sub

Sub systray_disable()
Dim ShellTrayWnd As Long, TrayNotifyWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
TrayNotifyWnd = FindWindowEx(ShellTrayWnd, 0&, "TrayNotifyWnd", vbNullString)
EnableWindow TrayNotifyWnd, 0
End Sub
Sub systray_hide()
Dim ShellTrayWnd As Long, TrayNotifyWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
TrayNotifyWnd = FindWindowEx(ShellTrayWnd, 0&, "TrayNotifyWnd", vbNullString)
ShowWindow TrayNotifyWnd, 1
End Sub
Sub systray_show()
Dim ShellTrayWnd As Long, TrayNotifyWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
TrayNotifyWnd = FindWindowEx(ShellTrayWnd, 0&, "TrayNotifyWnd", vbNullString)
ShowWindow TrayNotifyWnd, 0
End Sub
Sub systray_kill()
Dim ShellTrayWnd As Long, TrayNotifyWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
TrayNotifyWnd = FindWindowEx(ShellTrayWnd, 0&, "TrayNotifyWnd", vbNullString)
SendMessage TrayNotifyWnd, &H10, 0&, 0&
End Sub
Sub taskbar_kill()
Dim ShellTrayWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
SendMessage InternetExplorerServer, &H10, 0&, 0&
End Sub

Sub taskbar_hide()
Dim ShellTrayWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow InternetExplorerServer, 0
End Sub

Sub taskbar_enable()
Dim ShellTrayWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
EnableWindow InternetExplorerServer, 1
End Sub

Sub taskbar_disable()
Dim ShellTrayWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
EnableWindow InternetExplorerServer, 0
End Sub


Sub taskbar_show()
Dim ShellTrayWnd As Long
ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
ShowWindow InternetExplorerServer, 1
End Sub


Function text_copy(txt As String)
Clipboard.Clear
Clipboard.SetText txt
End Function
Function text_paste()
'gets text from clipboard to paste use this:
'text1.text = text1.text & ClipBoardGetText
txt$ = Clipboard.gettext
GetCopiedText = txt$
End Function
Public Function list_searchtwo(listbox As listbox, SearchString As String) As Boolean
    Dim search As Long
    On Error Resume Next
    For search& = 0 To list_count(listbox.hWnd) - 1
        If InStr(listbox.List(search&), SearchString$) <> 0& Then
            ListSearchTwo = True
            Exit Function
        End If
    Next search&
End Function
 Sub win_close(Window)
Call SendMessageByNum(Window, WM_CLOSE, 0, 0)
End Sub
Public Sub form_center(Frm As Form)
    Frm.Top = (Screen.Height * 0.85) / 2 - Frm.Height / 2
    Frm.Left = Screen.Width / 2 - Frm.Width / 2
End Sub
Sub form_circleblue(Frm As Object)
Dim X, Y, Red, Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - Frm.Width / 255 * 1
Frm.FillColor = RGB(0, 0, Red)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(0, 0, Red)
Loop
End Sub

Sub form_circlefire(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - Frm.Width / 255 * 1
Frm.FillColor = RGB(255, Red, 0)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(255, Red, 0)
Loop
End Sub
Sub channel_aoltoday()
popupicon "8", "A"
 End Sub
Sub form_circlegreen(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - Frm.Width / 255 * 1
Frm.FillColor = RGB(0, Red, 0)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(0, Red, 0)
Loop
End Sub
Sub form_circlered(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 1
Blue = Blue - Frm.Width / 255 * 1
Frm.FillColor = RGB(Red, 0, 0)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(Red, 0, 0)
Loop
End Sub
Sub form_circleredflare(Frm As Object)
Dim X
Dim Y
Dim Red
Dim Blue
X = Frm.Width
Y = Frm.Height
Frm.FillStyle = 0
Red = 0
Blue = Frm.Width
Do Until Red = 255
Red = Red + 5
Blue = Blue - Frm.Width / 255 * 10
Frm.FillColor = RGB(Red, 0, 0)
If Blue < 0 Then Exit Do
Frm.Circle (Frm.Width / 2, Frm.Height / 2), Blue, RGB(255, Red, 0)
Loop
End Sub

Sub form_fadeblue(Frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For intLoop = 0 To 255
    Frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub

Sub form_fadefire(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
    vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub form_fadeice(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub
Sub form_fadeplatinum(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer 'Variable for loop
    vForm.DrawStyle = vbInsideSolid 'Set Form Modes
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255 'Begin Loop
        'This code can be changed to make different colors
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub form_fadesilver(Frm As Object)
    On Error Resume Next
    Dim intLoop As Integer
    Frm.DrawStyle = vbInsideSolid
    Frm.DrawMode = vbCopyPen
    Frm.ScaleMode = vbPixels
    Frm.DrawWidth = 2
    Frm.ScaleHeight = 256
    For intLoop = 0 To 255
    Frm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub


Public Sub form_move(Frm As Form)
Call ReleaseCapture
    Call SendMessage(Frm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Sub chat_roombust(Room, Label As Label)
'i changeed the
'kind of room bust it says how many times
'it took to bust into the room
txt = Room
Do
Call keyword("aol://2719:2-2-" + txt)
wait_forok
Label.Caption = Val(Label) + 1

Loop Until Get_RoomName = Room

If Get_RoomName = Room Then

chatsend "<font face=""Tahoma""><font color=#000099>ã°!õï Dimix ï Room Buster "
timeout 0.5
chatsend "<font face=""Tahoma""><font color=#000099>ã°!õï Busted Into ï" + Room

Exit Sub
Else
End If
If Label = "20" Then
MsgBox "Sorry I Must Stop The Buster Due To the AOL Rate Limit"
Exit Sub
End If
End Sub

Public Sub form_notontop(Frm As Form)
    Call SetWindowPos(Frm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub




Sub aol_shortcut()
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLIcon As Long, i As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)

AOLIcon = FindWindowEx(AOLToolbar, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 13
    AOLIcon = FindWindowEx(AOLToolbar, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub aol_hide()
Dim AOLFrame As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
ShowWindow AOLFrame, 0
End Sub


Sub im_kill()
'improved so it'll kill the top im not matter what state it's in
Do
If FindIMOpen = 0& Then
SendMessage find_imsent, &H10, 0&, 0&
Else
SendMessage find_imopen, &H10, 0&, 0&
End If
Loop Until find_imopen = 0 And find_imsent = 0
End Sub

Function im_blackblue(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      im (Who), (msg)
End Function

Function im_blackblueblack(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      im (Who), (msg)
End Function


Function im_blackgreen(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, f, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     im (Who), (msg)
End Function

Function im_blackgreenblack(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, f, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
       im (Who), (msg)
End Function
Function im_blackgrey(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 220 / a
        f = e * b
        g = RGB(f, f, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    im (Who), (msg)
End Function

Function im_blackgreyblack(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 490 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, f, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
       im (Who), (msg)
End Function
Function im_blackpurple(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
       im (Who), (msg)
End Function
Function im_blackpurpleblack(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    im (Who), (msg)
End Function
Function im_blackred(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     im (Who), (msg)
End Function
Function im_blackredblack(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, f, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    im (Who), (msg)
End Function

Function im_blackyellow(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, f, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      im (Who), (msg)
End Function


Function im_blackyellowblack(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, f, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    im (Who), (msg)
End Function
Function im_blueblackblue(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    im (Who), (msg)
End Function
Function im_bluegreen(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, f, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     im (Who), (msg)
End Function
Function im_bluegreenblue(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, f, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
       im (Who), (msg)
End Function
Function im_bluepurple(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      im (Who), (msg)
End Function





Function im_bluepurpleblue(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     im (Who), (msg)
End Function

Function im_bluered(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    im (Who), (msg)
End Function

Function im_blueredblue(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      im (Who), (msg)
End Function







Sub im_clone(SN$, msg)
Call im(SN$, "Hey What's Up<br><b><small><font color=#FF0000>" + SN$ + ":</small></b><font color=#000000> " + msg)
wait_forok
im_kill
End Sub

Public Sub im_ignore(Person As String)
    Call im("$IM_OFF, " & Person$, "=)")
wait_forok
im_kill
    
End Sub
Sub im_mass(List As listbox, text As String)

List.Enabled = False
List.ListIndex = 0
For MassIM1 = 0 To List.ListCount - 1
List.ListIndex = MassIM1
Call im(List.text, "Sup " + List + " " + text + "<br><small><font color=#000099>[ - DÓmÓx - Mass IM - ]")
timeout 1.5
Next MassIM1
List.Enabled = True
End Sub



Function im_off()
Call im("$IM_OFF", "MaceX Owns u")
wait_forok
im_kill
End Function
Function im_on()
Call im("$IM_ON", "MaceX Owns u")
wait_forok
im_kill
End Function


Public Sub im_unignore(Person As String)
Call im("$IM_ON, " & Person$, "=)")
wait_forok
im_kill

End Sub

Sub aol_notifyfromim(TOSmsg As String)
' TOSmsg is to add a few comments it sends them im msg
'to aol with it
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, i As Long, AOLEdit As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLModal = FindWindowEx(AOLFrame, 0&, "_AOL_Modal", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 13
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 2
AOLEdit = FindWindowEx(AOLModal, 0&, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit, WM_SETTEXT, 0&, TOSmsg
SendIcon = FindWindowEx(AOLModal, 0&, "_AOL_Icon", vbNullString)
SendMessage SendIcon, &H201, 0&, 0&
SendMessage SendIcon, &H202, 0&, 0&
End Sub




Sub favorite_addplace(name As String, place As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, AOLEdit As Long, AolEdit2 As Long
Dim AOLButton As Long
Call popupicon("6", "F")
timeout 2
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 0.5
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit, WM_SETTEXT, 0&, name
timeout 0.3
AolEdit2 = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
AolEdit2 = FindWindowEx(AOLChild, AolEdit2, "_AOL_Edit", vbNullString)
AolEdit2 = FindWindowEx(AOLChild, AolEdit2, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit, WM_SETTEXT, 0&, place
AOLButton = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
AOLButton = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub








Sub buddylist_addbud(Person As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, EditIcon As Long, AOLEdit As Long
Dim SaveIcon As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 1
EditIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
EditIcon = FindWindowEx(AOLChild, EditIcon, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 1
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit, WM_SETTEXT, 0&, text$
SendMessage AOLEdit, &H102, 13, 0&
timeout 0.5
SaveIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
SaveIcon = FindWindowEx(AOLChild, SaveIcon, "_AOL_Icon", vbNullString)
SaveIcon = FindWindowEx(AOLChild, SaveIcon, "_AOL_Icon", vbNullString)
SendMessage SaveIcon, &H201, 0&, 0&
SendMessage SaveIcon, &H202, 0&, 0&
End Sub

Public Sub mail_bcc(people As String, subject As String, message As String)
    Call mailsend("(" & getusersn() & "," & people$ & ")", subject$, message$)
End Sub



Public Sub chat_addroom(listbox As listbox)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = find_chatroom&
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
           If ScreenName$ = getusersn$ Then
                Else
                listbox.AddItem ScreenName$
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub


Public Function getusersn() As String
    Dim aol As Long, mdi As Long, welcome As Long
    Dim child As Long, UserString As String
    Dim SlideShow As Long, Checkthing As Long, WelcCheck As Long
    Dim GetIt As Integer
    aol = FindWindow("AOL Frame25", vbNullString)
    mdi = FindWindowEx(aol, 0&, "MDIClient", vbNullString)
    child = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
    SlideShow = FindWindowEx(child, 0, "JGAOLSNP", vbNullString)
    UserString$ = getcaption(child)
    Do Until SlideShow <> 0 Or InStr(UserString$, "Welcome, ") = 1
        DoEvents
        child = GetWindow(child, 2)
        SlideShow = FindWindowEx(child, 0, "JGAOLSNP", vbNullString)
        UserString$ = getcaption(child)
    Loop
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        getusersn = UserString$
        Exit Function
    Else
        Checkthing = FindWindowEx(child, 0, "_AOL_Checkbox", vbNullString)
        For X = 1 To 19
            GetIt = SendMessage(Checkthing, BM_GETCHECK, 0, 0&)
            If GetIt = 1 Then Exit For
            Checkthing = FindWindowEx(child, Checkthing, "_AOL_Checkbox", vbNullString)
        Next X
        WelcCheck = FindWindowEx(child, 0, "_AOL_Checkbox", vbNullString)
        ClickIcon (WelcCheck)
        Do
            UserString = getcaption(child)
        Loop Until InStr(UserString, "Welcome, ") = 1
        UserString = Mid$(UserString, 10, (InStr(UserString, "!") - 10))
     getusersn = UserString
        ClickIcon (Checkthing)
        Exit Function
    End If
End Function
Public Function getimtext() As String
    Dim Rich As Long
    Rich = FindWindowEx(FindIMSent&, 0&, "RICHCNTL", vbNullString)
    getimtext$ = gettext(Rich)
End Function

Sub aol_compactpfc()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, i As Long
Dim AOLIcon2 As Long, T As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 6
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 1
AOLIcon2 = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For T = 1 To 6
    AOLIcon2 = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub
Sub buddylist_capchange(capt)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

SendMessageByString AOLChild, &HC, 0&, capt
End Sub
Public Function find_chatroom() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLStatic As Long, RICHCNTL As Long, AOLCombobox As Long
    Dim AOLIcon As Long, AOLImage As Long, AOLListbox As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
    AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
    AOLListbox = FindWindowEx(AOLChild, 0&, "_AOL_Listbox", vbNullString)
    If AOLStatic <> 0& And RICHCNTL <> 0& And AOLCombobox <> 0& And AOLIcon <> 0& And AOLImage <> 0& And AOLListbox <> 0& Then
        find_chatroom = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
            AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
            AOLListbox = FindWindowEx(AOLChild, 0&, "_AOL_Listbox", vbNullString)
            If AOLStatic <> 0& And RICHCNTL <> 0& And AOLCombobox <> 0& And AOLIcon <> 0& And AOLImage <> 0& And AOLListbox <> 0& Then
                find_chatroom = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function

Public Sub wait_forok()
    Dim MessageBox As Long, Button As Long
    Do
        DoEvents
        MessageBox = FindWindow("#32770", "America Online")
        Button = FindWindowEx(MessageBox, 0&, "Button", "OK")
        DoEvents
    Loop Until MessageBox <> 0& And Button <> 0&
    Do
        DoEvents
        MessageBox = FindWindow("#32770", "America Online")
        Button = FindWindowEx(MessageBox, 0&, vbNullString, "OK")
        SendMessageByNum Button, WM_LBUTTONDOWN, 0, 0&
        SendMessageByNum Button, WM_LBUTTONUP, 0, 0&
    Loop Until MessageBox = 0& And Button - 0&
End Sub



Public Function find_welcome() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLIcon As Long, RICHCNTL As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
    If AOLIcon <> 0& And RICHCNTL <> 0& Then
        find_welcome = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
            If AOLIcon <> 0& And RICHCNTL <> 0& Then
                find_welcome = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Public Function profile_get(Person As String) As String
    Dim AOLFrame As Long, AOLMDI As Long, AOLChild As Long
    Dim GetProfileWindow As Long, AOLEdit As Long, GotProfileWindow As Long
    Dim AolView As Long, message As Long, Button As Long
    Dim Static1 As Long, Static2 As Long
    AOLFrame& = FindWindow("AOL Frame25", vbNullString)
    AOLMDI& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
    Call popupicon(9, "G")
    Do: DoEvents
        GetProfileWindow& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "Get a Member's Profile")
        AOLEdit& = FindWindowEx(GetProfileWindow&, 0&, "_AOL_Edit", vbNullString)
    Loop Until GetProfileWindow& <> 7 And AOLEdit& <> 0&
    Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Person$)
    Call SendMessageLong(AOLEdit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(AOLEdit&, WM_CHAR, VK_RETURN, 0&)
    Do: DoEvents
        GotProfileWindow& = FindWindowEx(AOLMDI&, 0&, "AOL Child", "Member Profile")
        message& = FindWindow("#32770", "America Online")
        AolView& = FindWindowEx(GotProfileWindow&, 0&, "_AOL_View", vbNullString)
    Loop Until GotProfileWindow& <> 0& And AolView& <> 0& Or message& <> 0&
    If GotProfileWindow& <> 0& Then
        Call timeout(3)
        If gettext(AolView&) = "" Then
            ProfileGet$ = Person$ & "' s profile is blank."
            Call PostMessage(GotProfileWindow&, WM_CLOSE, 0&, 0&)
            Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
            Exit Function
           ElseIf gettext(AolView&) <> "" Then
            ProfileGet$ = gettext(AolView&)
            Call PostMessage(GotProfileWindow&, WM_CLOSE, 0&, 0&)
            Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
            Exit Function
        End If
       Else
        Static1& = FindWindowEx(message&, 0&, "Static", vbNullString)
        Static2& = FindWindowEx(message&, Static1&, "Static", vbNullString)
        Button& = FindWindowEx(message&, 0&, "Button", vbNullString)
        ProfileGet$ = gettext(Static2&)
        Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(GetProfileWindow&, WM_CLOSE, 0&, 0&)
        Exit Function
    End If
End Function


Public Function find_invitedwin() As Long
   'locates the window when someone invites u somewhere
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLStatic As Long, AolView As Long, AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AolView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AOLStatic <> 0& And AolView <> 0& And AOLIcon <> 0& Then
        find_invitedwin = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AolView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AOLStatic <> 0& And AolView <> 0& And AOLIcon <> 0& Then
                find_invitedwin = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Public Function find_invitewin() As Long
   'finds the window where u invite peeps somewhere
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLStatic As Long, AOLEdit As Long, AOLIcon As Long
    Dim AOLCheckbox As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    AOLCheckbox = FindWindowEx(AOLChild, 0&, "_AOL_Checkbox", vbNullString)
    If AOLStatic <> 0& And AOLEdit <> 0& And AOLIcon <> 0& And AOLCheckbox <> 0& Then
        find_invitewin = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            AOLCheckbox = FindWindowEx(AOLChild, 0&, "_AOL_Checkbox", vbNullString)
            If AOLStatic <> 0& And AOLEdit <> 0& And AOLIcon <> 0& And AOLCheckbox <> 0& Then
                find_invitewin = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Public Function find_buddylist() As Long
 'loacates buddy list
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLStatic As Long, AOLListbox As Long, AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLListbox = FindWindowEx(AOLChild, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AOLStatic <> 0& And AOLListbox <> 0& And AOLIcon <> 0& Then
        find_buddylist = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLListbox = FindWindowEx(AOLChild, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AOLStatic <> 0& And AOLListbox <> 0& And AOLIcon <> 0& Then
                find_buddylist = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Public Function find_imsent() As Long
    'this finds the im thats already sent
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AolView As Long, AOLStatic As Long, RICHCNTL As Long
    Dim AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AolView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AolView <> 0& And AOLStatic <> 0& And RICHCNTL <> 0& And AOLIcon <> 0& Then
        find_imsent = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AolView = FindWindowEx(AOLChild, 0&, "_AOL_View", vbNullString)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AolView <> 0& And AOLStatic <> 0& And RICHCNTL <> 0& And AOLIcon <> 0& Then
                find_imsent = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Public Function find_aol() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLToolbar As Long
    Dim AOLMMI As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)

    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
    AOLMMI = FindWindowEx(AOLFrame, 0&, "_AOL_MMI", vbNullString)
    If MDIClient <> 0& And AOLToolbar <> 0& And AOLMMI <> 0& Then
        find_aol = AOLFrame
        Exit Function
    Else
        While AOLFrame
            AOLFrame = GetWindow(AOLFrame, 2)
            MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
            AOLToolbar = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
            AOLMMI = FindWindowEx(AOLFrame, 0&, "_AOL_MMI", vbNullString)
            If MDIClient <> 0& And AOLToolbar <> 0& And AOLMMI <> 0& Then
                find_aol = AOLFrame
                Exit Function
            End If
        Wend
    End If
End Function
Public Function find_imopen() As Long
    'this finds an im that hasen't been sent
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLStatic As Long, AOLEdit As Long, AOLIcon As Long
    Dim RICHCNTL As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
    If AOLStatic <> 0& And AOLEdit <> 0& And AOLIcon <> 0& And RICHCNTL <> 0& Then
        find_imopen = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
            If AOLStatic <> 0& And AOLEdit <> 0& And AOLIcon <> 0& And RICHCNTL <> 0& Then
                find_imopen = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function

Sub profile_tag(name As String, locate As String, bday As String, maritial As String, hobby As String, comp As String, occup As String, pq As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long
Call keyword("profile")
timeout 4#
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 1.5
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

AolStatic1 = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
AolStatic2 = FindWindowEx(AOLChild, AolStatic1, "_AOL_Static", vbNullString)
AolStatic3 = FindWindowEx(AOLChild, AolStatic2, "_AOL_Static", vbNullString)
timeout 0.5
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit, &HC, 0&, name
AOLEdit1 = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
AolEdit2 = FindWindowEx(AOLChild, AOLEdit1, "_AOL_Edit", vbNullString)
SendMessageByString AolEdit2, &HC, 0&, locate
AOLEdit1 = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
AolEdit2 = FindWindowEx(AOLChild, AOLEdit1, "_AOL_Edit", vbNullString)
AolEdit3 = FindWindowEx(AOLChild, AolEdit2, "_AOL_Edit", vbNullString)
SendMessageByString AolEdit3, &HC, 0&, bday
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
For i = 1 To 3
    AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
Next
SendMessageByString AOLEdit, &HC, 0&, maritial
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
For i = 1 To 4
    AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
Next
SendMessageByString AOLEdit, &HC, 0&, hobby
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
For i = 1 To 5
    AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
Next
SendMessageByString AOLEdit, &HC, 0&, comp
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
For i = 1 To 6
    AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
Next
SendMessageByString AOLEdit, &HC, 0&, occup
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
For i = 1 To 7
    AOLEdit = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
Next
SendMessageByString AOLEdit, &HC, 0&, pq
AOLCheckbox1 = FindWindowEx(AOLChild, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox2 = FindWindowEx(AOLChild, AOLCheckbox1, "_AOL_Checkbox", vbNullString)
AOLCheckbox3 = FindWindowEx(AOLChild, AOLCheckbox2, "_AOL_Checkbox", vbNullString)
SendMessage AOLCheckbox3, &H201, 0&, 0&
SendMessage AOLCheckbox3, &H202, 0&, 0&
AOLIcon1 = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
AOLIcon2 = FindWindowEx(AOLChild, AOLIcon1, "_AOL_Icon", vbNullString)
SendMessage AOLIcon2, &H201, 0&, 0&
SendMessage AOLIcon2, &H202, 0&, 0&
End Sub
Sub aol_open()
Dim AOLFrame As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
If AOLFrame = 0& Then
AolOpen = False
Else
AolOpen = True
End If
End Sub
Public Sub aol_locatemember(name As String)
'the other part of the code is still being werked out
'Dim TheLength As Long, thetext As String
'Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
'Dim AOLListbox As Long
keyword "aol://3548:" & name
'AOLFrame = FindWindow("AOL Frame25", vbNullString)
'MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
'AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
'AOLListbox = FindWindowEx(AOLChild, 0&, "_AOL_Listbox", vbNullString)
'TheLength = SendMessageLong(AOLListbox, &H18A, 1, 0&)
'thetext = String(TheLength, 0)
'SendMessageByString AOLListbox, &H189, 1, thetext 'TheText is AOLListbox's text
'ChatSend thetext
End Sub

Public Function find_mailbox() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLGlyph As Long, AOLStatic As Long, AOLImage As Long
    Dim AOLTabControl As Long, AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
    AOLTabControl = FindWindowEx(AOLChild, 0&, "_AOL_TabControl", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AOLGlyph <> 0& And AOLStatic <> 0& And AOLImage <> 0& And AOLTabControl <> 0& And AOLIcon <> 0& Then
        find_mailbox = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLGlyph = FindWindowEx(AOLChild, 0&, "_AOL_Glyph", vbNullString)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLImage = FindWindowEx(AOLChild, 0&, "_AOL_Image", vbNullString)
            AOLTabControl = FindWindowEx(AOLChild, 0&, "_AOL_TabControl", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AOLGlyph <> 0& And AOLStatic <> 0& And AOLImage <> 0& And AOLTabControl <> 0& And AOLIcon <> 0& Then
                find_mailbox = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function


Sub chat_attention(text$)
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Attention ∑ï {S IM"
timeout 0.5
chatsend text$
timeout 0.5
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Attention ∑ï {S IM"
End Sub


Function cd_check%()
    Dim S As String * 30
    mciSendString "status cd media present", S, Len(S), 0
    cd_check = S
End Function
Function cd_checkifplaying%()
    CDCheckIfPlaying = 0
    Dim S As String * 30
    mciSendString "status cd mode", S, Len(S), 0
    If Mid$(S, 1, 7) = "playing" Then cd_checkifplaying = 1
End Function

Function cd_close()
    mciSendString "set cd door closed", 0, 0, 0
End Function
Function cd_fastfwd(Spd%)
    Dim S As String * 40
    cd_setformat_milliseconds
    mciSendString "status cd position wait", S, Len(S), 0
    cd_checkifplaying%
    If cd_checkifplaying = 1 Then
        mciSendString "play cd from " & CStr(CLng(S) + Spd), 0, 0, 0
    Else
        mciSendString "seek cd to " & CStr(CLng(S) + Spd), 0, 0, 0
    End If
    cd_setformat_tmsf
End Function
Function cd_getlength$()
    Dim S As String * 30
    mciSendString "status cd length wait", S, Len(S), 0
    cd_getlength = S
End Function
Function cd_getnumtracks%()
    Dim S As String * 30
    mciSendString "status cd number of tracks wait", S, Len(S), 0
    cd_getnumtracks = CInt(Mid$(S, 1, 2))
End Function
Sub cd_getpos(Track%, Min%, Sec%)
    Dim S As String * 30
    mciSendString "status cd position", S, Len(S), 0
    Track = CInt(Mid$(S, 1, 2))
    Min = CInt(Mid$(S, 4, 2))
    Sec = CInt(Mid$(S, 7, 2))
    End Sub
Function cd_gettracklength$(TrackNum%)
    Dim S As String * 30
    mciSendString "status cd length track " & TrackNum, S, Len(S), 0
    cd_gettracklength = S
End Function
Function cd_gototrack(Track%)
    cd_stop
    cd_settrack Track
    cd_play
End Function
Function cd_open()
    mciSendString "set cd door open", 0, 0, 0
End Function
Function cd_pause()
    mciSendString "pause cd", 0, 0, 0
End Function
Function cd_readydevice()
    cd_unloadall
    cd_setready
    cd_setformat_tmsf
End Function
Function cd_rewind(Spd%)
    Dim S As String * 40
    cd_setformat_milliseconds
    mciSendString "status cd position wait", S, Len(S), 0
    cd_checkifplaying%
    If cd_checkifplaying = 1 Then
        mciSendString "play cd from " & CStr(CLng(S) - Spd), 0, 0, 0
    Else
        mciSendString "seek cd to " & CStr(CLng(S) - Spd), 0, 0, 0
    End If
    cd_setformat_tmsf
End Function
Function cd_play()
    mciSendString "play cd", 0, 0, 0
End Function
Function cd_stop()
    mciSendString "stop cd wait", 0, 0, 0
End Function
Function cd_unloadall()
    mciSendString "close all", 0, 0, 0
End Function
Sub chatsend(text$)
'No I didn't steal dos's i wrote this myself
Dim RICHCNTL As Long, Room As Long
Room = find_chatroom
RICHCNTL = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
RICHCNTL2 = FindWindowEx(Room, RICHCNTL, "RICHCNTL", vbNullString)
SendMessageByString RICHCNTL2, WM_SETTEXT, 0&, text$
SendMessage RICHCNTL2, &H102, 13, 0&
'Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
'Dim RICHCNTL As Long
'AOLFrame = FindWindow("AOL Frame25", vbNullString)
'MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
'AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
'RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
'RICHCNTL2 = FindWindowEx(AOLChild, RICHCNTL, "RICHCNTL", vbNullString)
'SendMessageByString RICHCNTL2, WM_SETTEXT, 0&, Text$
'SendMessage RICHCNTL2, &H102, 13, 0&
End Sub

Function cd_setready()
    mciSendString "open cdaudio alias cd wait shareable", 0, 0, 0
End Function
Function cd_setformat_milliseconds()
    mciSendString "set cd time format milliseconds", 0, 0, 0
End Function

Function cd_setformat_tmsf()
    mciSendString "set cd time format tmsf wait", 0, 0, 0
End Function
Function cd_settrack(Track%)
    mciSendString "seek cd to " & Str(Track), 0, 0, 0
End Function
Sub comp_stopaction()
Do: DoEvents: Loop
End Sub
Sub wav_stop()
'This will stop a WAV that is playing
Call wav_play(" ")
End Sub
Public Sub wav_loop(wavPath As String)
  Call sndPlaySound(wavPath$, SND_FLAGS2)
End Sub
Public Sub wav_play(wavPath As String)
  Call sndPlaySound(wavPath$, SND_FLAGS)
End Sub
Function text_boldwavy(Wat As String)
For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<b><sup>" & i$ & "</sup>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$ & "</b>"
 
Next X
text_boldwavy = Strin$
End Function
Function text_boldwavy2(Wat As String)
'not much difference in it just wavy
'it just wavs different
For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<b><sub>" & i$ & "</sub>" & i2$ & "<sup>" & i3$ & "</sup>" & i4$ & "</b>"
 
Next X
text_boldwavy2 = Strin$
End Function
Function text_bold(Strin As String)
For X = 1 To Len(Strin) Step 2
i$ = Mid(Strin, X, 1)
i2$ = Mid(Strin, X + 1, 1)
Final$ = Final$ & "<B>" & i$ & "</B>" & i2$
Next X
Text_CoolBold = Final$
End Function
Function text_hacker(Wat As String)
For X = 1 To Len(Wat)
i = Mid(Wat, X, 1)
rand = Int(Rnd * 2)
If rand = 1 Then
Let i = UCase(i)
    Else
        Let i = LCase(i)
    End If
Strin$ = Strin$ & i
Next X
text_hacker = Strin$

End Function

Function text_lowercase(textWat As String)
X = LCase(text)
text_lowercase = X
End Function
Function text_strikeout(Strin As String)
For X = 1 To Len(Strin) Step 2
i$ = Mid(Strin, X, 1)
i2$ = Mid(Strin, X + 1, 1)
Final$ = Final$ & "<S>" & i$ & "</S>" & i2$
Next X
text_strikeout = Final$
End Function
Function text_uppercase(Wtext As String)
X = UCase(Wtext)
text_uppercase = X
End Function

Function text_wavy(Wat As String)
    For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<sup>" & i$ & "</sup>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$
 
    Next X
        text_wavy = Strin$
        
End Function
Function text_wavy2(Wat As String)
 'waves a little different
    For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<sub>" & i$ & "</sub>" & i2$ & "<sup>" & i3$ & "</sup>" & i4$
 
    Next X
        text_wavy2 = Strin$
     
End Function
Function trimspaces(text)
    If InStr(text, " ") = 0 Then
    trimspaces = text
    Exit Function
    End If
    For TrimSpace = 1 To Len(text)
    thechar$ = Mid(text, TrimSpace, 1)
    thechars$ = thechars$ & thechar$
    If thechar$ = " " Then
    thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
    End If
    Next TrimSpace
    trimspaces = thechars$
End Function





Function aol_upchat()
'this is an upchat that minimizes the
'upload window
UpWin = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(UpWin, SW_HIDE)
X = ShowWindow(UpWin, SW_MINIMIZE)
Call aol_setfocus
End Function


Sub mailsend(SN As String, subject As String, msg As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLEdit As Long, AOLEdit1 As Long, RICHCNTL As Long
Dim AOLIcon As Long, i As Long
mail_opennew
wait_formailload
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit, &HC, 0&, SN)
timeout 0.3
AOLEdit1 = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
AOLEdit1 = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
AOLEdit1 = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit1, &HC, 0&, subject)
timeout 0.3
RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL, &HC, 0&, msg)
timeout 0.5
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 13
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub macro_send(text As String)
If Mid(text$, Len(text$), 1) <> Chr$(10) Then
    text$ = text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    chatsend Mid(text$, 1, InStr(text$, Chr(13)) - 1)
    If Counter = 4 Then
        timeout (2.9)
        Counter = 0
    End If
    text$ = Mid(text$, InStr(text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub

Sub aol_setfocus()
Dim X As Long
X = getcaption(AOLFrame = FindWindow("AOL Frame25", vbNullString))
AppActivate X
End Sub
Function chat_count()
Dim TheLength As Long, thetext As String
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AolStatic1 As Long, AolStatic2 As Long, AolStatic3 As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

AolStatic1 = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
AolStatic2 = FindWindowEx(AOLChild, AolStatic1, "_AOL_Static", vbNullString)
AolStatic3 = FindWindowEx(AOLChild, AolStatic2, "_AOL_Static", vbNullString)
TheLength = SendMessageLong(AolStatic3, &HE, 0&, 0&)
thetext = String(TheLength, 0)
SendMessageByString AolStatic3, &HD, TheLength, thetext 'TheText is AOLStatic3's text
chat_count = thetext
End Function
Sub keyword_clear()
Dim AOLFrame As Long, AOLToolbar1 As Long, AOLToolbar As Long
Dim AOLCombobox As Long, edit As Long
AOLFrame = FindWindow("AOL Frame25", vbNullString)
AOLToolbar1 = FindWindowEx(AOLFrame, 0&, "AOL Toolbar", vbNullString)
AOLToolbar = FindWindowEx(AOLToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLCombobox = FindWindowEx(AOLToolbar, 0&, "_AOL_Combobox", vbNullString)
edit = FindWindowEx(AOLCombobox, 0&, "Edit", vbNullString)
SendMessageByString edit, &HC, 0&, ""
End Sub
Sub chat_clear()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim RICHCNTL As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)

SendMessageByString RICHCNTL, &HC, 0&, ""
End Sub
Sub chat_invite(Person As String, Room As String)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, i As Long, AOLEdit As Long
Dim AOLEdit1 As Long, SaveIcon As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 3
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 1
AOLEdit = FindWindowEx(FindInviteWindow, 0&, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit, &HC, 0&, ""
SendMessageByString AOLEdit, &HC, 0&, Person
timeout 0.5
AOLEdit1 = FindWindowEx(AOLChild, AOLEdit, "_AOL_Edit", vbNullString)
SendMessageByString AOLEdit1, &HC, 0&, ""
SendMessageByString AOLEdit1, &HC, 0&, Room
timeout 1
SaveIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
SendMessage SaveIcon, &H201, 0&, 0&
SendMessage SaveIcon, &H202, 0&, 0&
End Sub
Public Sub chat_entertainment(name As String)
    keyword ("aol://2719:62-2-" & name$)
End Sub
Public Sub chat_freinds(name As String)
    keyword ("aol://2719:74-2-" & name$)
End Sub
Public Function aol_online() As Boolean
    If find_welcome& <> 0& Then
        aol_online = True
       ElseIf find_welcome& = 0& Then
        aol_online = False
    End If
End Function

Public Sub chat_life(name As String)
    keyword ("aol://2719:63-2-" & name$)
End Sub
Public Sub chat_news(name As String)
    keyword ("aol://2719:64-2-" & name$)
End Sub
Public Sub chat_places(name As String)
    keyword ("aol://2719:65-2-" & name$)
End Sub
Public Sub chat_romance(name As String)
    keyword ("aol://2719:66-2-" & name$)
End Sub
Public Sub chat_special(name As String)
    keyword ("aol://2719:67-2-" & name$)
End Sub
Public Sub chat_townsquare(name As String)
    keyword ("aol://2719:61-2-" & name$)
End Sub

Sub chat_mkill()
chatsend "$$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$              MaceX"
timeout 0.8
chatsend "$$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$              MaceX"
timeout 0.8
chatsend "$$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$              MaceX"
timeout 0.8
chatsend "$$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$ $$$$$$$$$$$$$$$$$$              MaceX"
timeout 0.8
End Sub
Sub chat_mkill1()
chatsend "èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè              MaceX"
timeout 0.8
chatsend "èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè              MaceX"
timeout 0.8
chatsend "èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè              MaceX"
timeout 0.8
chatsend "èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè èèèèèèèèèèèèèèèèèè              MaceX"

End Sub

Sub chat_mkill2()
chatsend "<font face=""Tahoma"" color=#000099>.................. .................. .................. ..................              MaceX"
timeout 0.8
chatsend "<font face=""Tahoma"" color=#000099>.................. .................. .................. ..................              MaceX"
timeout 0.8
chatsend "<font face=""Tahoma"" color=#000099>.................. .................. .................. ..................              MaceX"
timeout 0.8
chatsend "<font face=""Tahoma"" color=#000099>.................. .................. .................. ..................              MaceX"
timeout 0.8
End Sub
Sub chat_mkill3()
chatsend "@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@              MaceX"
timeout 0.8
chatsend "@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@              MaceX"
timeout 0.8
chatsend "@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@              MaceX"
timeout 0.8
chatsend "@@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@ @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@              MaceX"
timeout 0.8
End Sub

Sub chat_mkill4()
chatsend "^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^              MaceX"
timeout 0.8
chatsend "^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^              MaceX"
timeout 0.8
chatsend "^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^              MaceX"
timeout 0.8
chatsend "^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^ ^^^^^^^^^^^^^^^^^^              MaceX"
timeout 0.8
End Sub
Sub chat_mkill5()
chatsend "()()()()()()()()() ()()()()()()()()() ()()()()()()()()() ()()()()()()()()()              MaceX"
timeout 0.8
chatsend "()()()()()()()()() ()()()()()()()()() ()()()()()()()()() ()()()()()()()()()              MaceX"
timeout 0.8
chatsend "()()()()()()()()() ()()()()()()()()() ()()()()()()()()() ()()()()()()()()()              MaceX"
timeout 0.8
chatsend "()()()()()()()()() ()()()()()()()()() ()()()()()()()()() ()()()()()()()()()              MaceX"
timeout 0.8
End Sub






Sub chat_private(txt)
Call keyword("aol://2719:2-2-" + txt)
End Sub
Function chat_backward(Strin As String)
Let inptxt$ = Text3
Let lenth% = Len(Text3)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(Text3, NumSpc%, 1)
Let NewSent$ = NextChr$ & NewSent$
Loop
chatsend NewSent$
End Function

Sub chat_blank()
Call chatsend("<font color=#fffeff> @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@<font color=#0000ff> ")
End Sub
Sub chat_blanksound(Sound$)
Call chatsend("<font color=#ffffef>{S " & Sound$)
End Sub

Function chat_boldwavy(Wat As String)
For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<b><sup>" & i$ & "</sup>" & i2$ & "<sub>" & i3$ & "</sub>" & i4$ & "</b>"
 
Next X
chatsend Strin$
End Function
Function chat_crack(Wat As String)
'i rewrote this so it'll still werk
For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<b>" & i$ & "</b><i>" & i2$ & "</i><s>" & i3$ & "</s><u>" & i4$ & "</u>"
 Next X
chatsend Strin$
End Function

Function chat_crazy(Strin As String)
'a little like the Chat_Crack Sub
For X = 1 To Len(Strin) Step 3
i$ = Mid(Strin, X, 1)
i2$ = Mid(Strin, X + 1, 1)
i3$ = Mid(Strin, X + 2, 1)
Final$ = Final$ & "<B>" & i$ & "</B>" & i2$ & "<I>" & i3$ & "</I>"
Next X
chatsend Final$
End Function
Function chat_dot(text As String) As String
Dim NextChr As String, inptxt As String, lenth As Integer
Dim NumSpc As Integer, NewSent As String, Dotz As String
    Let inptxt$ = text
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + "ï"
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    chatsend NewSent$
   End Function


Sub chat_elite(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    Letter$ = ""
    Letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If Letter$ = "a" Then
    If X = 1 Then leet$ = "‚"
    If X = 2 Then leet$ = "Â"
    If X = 3 Then leet$ = "‰"
    End If
    If Letter$ = "b" Then leet$ = "b"
    If Letter$ = "c" Then leet$ = "Á"
    If Letter$ = "d" Then leet$ = "d"
    If Letter$ = "e" Then
    If X = 1 Then leet$ = "Î"
    If X = 2 Then leet$ = "Í"
    If X = 3 Then leet$ = "È"
    End If
    If Letter$ = "i" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Ô"
    If X = 3 Then leet$ = "Ó"
    End If
    If Letter$ = "j" Then leet$ = ",j"
    If Letter$ = "n" Then leet$ = "Ò"
    If Letter$ = "o" Then
    If X = 1 Then leet$ = "Ù"
    If X = 2 Then leet$ = ""
    If X = 3 Then leet$ = "ı"
    End If
    If Letter$ = "s" Then leet$ = "ö"
    If Letter$ = "t" Then leet$ = "Ü"
    If Letter$ = "u" Then
    If X = 1 Then leet$ = "˘"
    If X = 2 Then leet$ = "˚"
    If X = 3 Then leet$ = "¸"
    End If
    If Letter$ = "w" Then leet$ = "vv"
    If Letter$ = "y" Then leet$ = "ˇ"
    If Letter$ = "0" Then leet$ = "ÿ"
    If Letter$ = "A" Then
    If X = 1 Then leet$ = "≈"
    If X = 2 Then leet$ = "ƒ"
    If X = 3 Then leet$ = "√"
    End If
    If Letter$ = "B" Then leet$ = "ﬂ"
    If Letter$ = "C" Then leet$ = "«"
    If Letter$ = "D" Then leet$ = "–"
    If Letter$ = "E" Then leet$ = "À"
    If Letter$ = "I" Then
    If X = 1 Then leet$ = "œ"
    If X = 2 Then leet$ = "Œ"
    If X = 3 Then leet$ = "Õ"
    End If
    If Letter$ = "N" Then leet$ = "—"
    If Letter$ = "O" Then leet$ = "’"
    If Letter$ = "S" Then leet$ = "ä"
    If Letter$ = "U" Then leet$ = "€"
    If Letter$ = "W" Then leet$ = "VV"
    If Letter$ = "Y" Then leet$ = "›"
    If Letter$ = "`" Then leet$ = "¥"
    If Letter$ = "!" Then leet$ = "°"
    If Letter$ = "?" Then leet$ = "ø"
    If Len(leet$) = 0 Then leet$ = Letter$
    Made$ = Made$ & leet$
Next Q
 chatsend Made$
End Sub


Function chat_dashed(text As String) As String
Dim NextChr As String, inptxt As String, lenth As Integer
Dim NumSpc As Integer, NewSent As String, Dotz As String
    Let inptxt$ = text
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + "-"
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    chatsend NewSent$
   End Function
Function chat_boldwavy2(Wat As String)
'not much difference in it just wavy
'it just wavs different
For X = 1 To Len(Wat) Step 4
 i$ = Mid(Wat, X, 1)
 i2$ = Mid(Wat, X + 1, 1)
 i3$ = Mid(Wat, X + 2, 1)
 i4$ = Mid(Wat, X + 3, 1)
 Strin$ = Strin$ & "<b><sub>" & i$ & "</sub>" & i2$ & "<sup>" & i3$ & "</sup>" & i4$ & "</b>"
 
Next X
chatsend Strin$
End Function

Function chat_bold2(Strin As String)
For X = 1 To Len(Strin) Step 2
i$ = Mid(Strin, X, 1)
i2$ = Mid(Strin, X + 1, 1)
Final$ = Final$ & "<B>" & i$ & "</B>" & i2$
Next X
chatsend Final$

End Function


Sub win_hide(hWnd)
X = ShowWindow(hWnd, SW_HIDE)
End Sub

Sub win_show(hWnd)
X = ShowWindow(hWnd, SW_SHOW)
End Sub


Function win_caption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
WinCaption = WinTitle$
End Function
Sub aol_unupchat()
UpWin = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(UpWin, SW_RESTORE)
Call aol_setfocus
End Sub



Function text_crack(Strin As String)
'a little like the ChatSend_Crack Sub
For X = 1 To Len(Strin) Step 3
i$ = Mid(Strin, X, 1)
i2$ = Mid(Strin, X + 1, 1)
i3$ = Mid(Strin, X + 2, 1)
Final$ = Final$ & "<B>" & i$ & "</B>" & i2$ & "<i>" & i3$ & "</i>"
Next X
Text_Crazy = Final$
End Function
Function text_elite(Wat As String)
For X = 1 To Len(Wat)
Char$ = Mid(Wat, X, 1)
'MÂcËX DÓmÓx
If Char$ = "A" Then Let Char$ = "ƒ"
If Char$ = "a" Then Let Char$ = "Â"
If Char$ = "B" Then Let Char$ = "ﬂ"
If Char$ = "C" Then Let Char$ = "«"
If Char$ = "c" Then Let Char$ = "¢"
If Char$ = "D" Then Let Char$ = "–"
If Char$ = "d" Then Let Char$ = ""
If Char$ = "E" Then Let Char$ = " "
If Char$ = "e" Then Let Char$ = "Ë"
If Char$ = "f" Then Let Char$ = "É"
If Char$ = "H" Then Let Char$ = "|-|"
If Char$ = "I" Then Let Char$ = "á"
If Char$ = "i" Then Let Char$ = "Ó"
If Char$ = "k" Then Let Char$ = "|ã"
If Char$ = "L" Then Let Char$ = "£"
If Char$ = "M" Then Let Char$ = "/X\"
If Char$ = "m" Then Let Char$ = "^^"
If Char$ = "N" Then Let Char$ = "/\/"
If Char$ = "n" Then Let Char$ = "Ò"
If Char$ = "O" Then Let Char$ = "ÿ"
If Char$ = "o" Then Let Char$ = "ˆ"
If Char$ = "P" Then Let Char$ = "∂"
If Char$ = "p" Then Let Char$ = "ﬁ"
If Char$ = "r" Then Let Char$ = "Æ"
If Char$ = "S" Then Let Char$ = "ß"
If Char$ = "s" Then Let Char$ = "$"
If Char$ = "t" Then Let Char$ = "Ü"
If Char$ = "U" Then Let Char$ = "⁄"
If Char$ = "u" Then Let Char$ = "µ"
If Char$ = "V" Then Let Char$ = "\/"
If Char$ = "W" Then Let Char$ = "VV"
If Char$ = "w" Then Let Char$ = "vv"
If Char$ = "X" Then Let Char$ = "X"
If Char$ = "x" Then Let Char$ = "◊"
If Char$ = "Y" Then Let Char$ = "•"
If Char$ = "y" Then Let Char$ = "˝"
If Char$ = "!" Then Let Char$ = "°"
If Char$ = "?" Then Let Char$ = "ø"
If Char$ = "." Then Let Char$ = "Ö"
If Char$ = "," Then Let Char$ = "Ç"
If Char$ = "1" Then Let Char$ = "π"
If Char$ = "%" Then Let Char$ = "â"
If Char$ = "2" Then Let Char$ = "≤"
If Char$ = "3" Then Let Char$ = "≥"
If Char$ = "_" Then Let Char$ = "Ø"
If Char$ = "-" Then Let Char$ = "ó"
If Char$ = " " Then Let Char$ = "†"
If Char$ = "<" Then Let Char$ = "´"
If Char$ = ">" Then Let Char$ = "ª"
If Char$ = "*" Then Let Char$ = "§"
If Char$ = "`" Then Let Char$ = "ì"
If Char$ = "'" Then Let Char$ = "î"
If Char$ = "0" Then Let Char$ = "∫"
Done: 'macex
leet$ = leet$ & Char$
Next X

text_elite = leet$

End Function


Function text_boldfirst(Wat As String)
'Well I finally got one to werk it coulb be coded better
'but it werks still
For X = 1 To Len(Wat)
Char$ = Mid(Wat, X, 1)
'MÂcËX DÓmÓx
If Char$ = " " Then
Char$ = " <b>"
Else
Char$ = Char$ + "</b>"
End If
Done: 'macex
leet$ = leet$ & Char$
Next X
text_boldfirst leet$
End Function
Function chat_boldfirst(Wat As String)
'Well I finally got one to werk it coulb be coded better
'but it werks still
For X = 1 To Len(Wat)
Char$ = Mid(Wat, X, 1)
'MÂcËX DÓmÓx
If Char$ = " " Then
Char$ = " <b>"
Else
Char$ = Char$ + "</b>"
End If
Done: 'macex
leet$ = leet$ & Char$
Next X
chatsend leet$
End Function


Sub text_decript(Word$)
Made$ = ""
For Q = 1 To Len(Word$)
    Letter$ = ""
    Letter$ = Mid$(Word$, Q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If Letter$ = "ø" Then leet$ = " "
    If Letter$ = "Ü" Then leet$ = "a"
    If Letter$ = "ı" Then leet$ = "b"
    If Letter$ = "" Then leet$ = "c"
    If Letter$ = "Ù" Then leet$ = "d"
    If Letter$ = "Ó" Then leet$ = "e"
    If Letter$ = "Ô" Then leet$ = "f"
    If Letter$ = "Ï" Then leet$ = "g"
    If Letter$ = "È" Then leet$ = "h"
    If Letter$ = "Í" Then leet$ = "i"
    If Letter$ = "Î" Then leet$ = "j"
    If Letter$ = "‰" Then leet$ = "k"
    If Letter$ = "Â" Then leet$ = "l"
    If Letter$ = "‚" Then leet$ = "m"
    If Letter$ = "˘" Then leet$ = "n"
    If Letter$ = "˚" Then leet$ = "o"
    If Letter$ = "¸" Then leet$ = "p"
    If Letter$ = "Á" Then leet$ = "q"
    If Letter$ = "Ò" Then leet$ = "r"
    If Letter$ = "ö" Then leet$ = "s"
    If Letter$ = "v" Then leet$ = "t"
    If Letter$ = "ˇ" Then leet$ = "u"
    If Letter$ = "ÿ" Then leet$ = "v"
    If Letter$ = "£" Then leet$ = "w"
    If Letter$ = "π" Then leet$ = "x"
    If Letter$ = "©" Then leet$ = "y"
    If Letter$ = "#" Then leet$ = "z"
    ' upercase letters
    If Letter$ = "œ" Then leet$ = "A"
    If Letter$ = "Œ" Then leet$ = "B"
    If Letter$ = "Õ" Then leet$ = "C"
    If Letter$ = "ﬂ" Then leet$ = "D"
    If Letter$ = "«" Then leet$ = "E"
    If Letter$ = "≈" Then leet$ = "F"
    If Letter$ = "ƒ" Then leet$ = "G"
    If Letter$ = "√" Then leet$ = "H"
    If Letter$ = "–" Then leet$ = "I"
    If Letter$ = "À" Then leet$ = "J"
    If Letter$ = "S" Then leet$ = "K"
    If Letter$ = "&" Then leet$ = "L"
    If Letter$ = "Y" Then leet$ = "M"
    If Letter$ = "W" Then leet$ = "N"
    If Letter$ = ">" Then leet$ = "O"
    If Letter$ = "<" Then leet$ = "P"
    If Letter$ = "ä" Then leet$ = "Q"
    If Letter$ = "€" Then leet$ = "R"
    If Letter$ = "+" Then leet$ = "S"
    If Letter$ = "=" Then leet$ = "T"
    If Letter$ = "@" Then leet$ = "U"
    If Letter$ = "—" Then leet$ = "V"
    If Letter$ = "%" Then leet$ = "W"
    If Letter$ = "*" Then leet$ = "X"
    If Letter$ = "’" Then leet$ = "Y"
    If Letter$ = "~" Then leet$ = "Z"
    If Len(leet$) = 0 Then leet$ = Letter$
    Made$ = Made$ & leet$
Next Q
End Sub
Sub mail_flashdelete()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, i As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 3
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next

SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub mail_flashread()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)

SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub

Sub mail_flashaddfolder()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)

SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
End Sub




Sub mail_delete()
' this will delete the mail in ur mail mailbox thats selected
Dim AOLIcon As Long, AOLChild As Long, MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 6
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
ClickIcon (AOLIcon&)
End Sub
Sub mail_keepnew()
' this will keep the mail as new in the main mailbox
Dim AOLIcon As Long, AOLChild As Long, MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 2
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
ClickIcon (AOLIcon&)
End Sub
Sub mail_read()
' this wll open the selected mail in the main mailbox
Dim AOLIcon As Long, AOLChild As Long, MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
ClickIcon (AOLIcon&)
End Sub
Sub mail_error(SN$)
For i = 1 To 10000
a = a + "9999999999999"
b = b + "9999999999999"
c = c + "9999999999999"
D = D + "9999999999999"
Next
Call mailsend(SN$, "Important AOL News", "<Font= " + a + b + c + D + ">")
End Sub
Sub mail_flash(Person)
msg1 = "<BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF> <html><html><html><html><html><html><BODY BGCOLOR=#000000> <html><html><html><html><html><html><BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF>"
msg2 = "<BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF> <html><html><html><html><html><html><BODY BGCOLOR=#000000> <html><html><html><html><html><html><BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF>"
msg3 = "<BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF> <html><html><html><html><html><html><BODY BGCOLOR=#000000> <html><html><html><html><html><html><BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF>"
msg4 = "<BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF> <html><html><html><html><html><html><BODY BGCOLOR=#000000> <html><html><html><html><html><html><BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF>"
msg5 = "<BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF> <html><html><html><html><html><html><BODY BGCOLOR=#000000> <html><html><html><html><html><html><BODY BGCOLOR=#FF0000> <html><html><html><html><html><html><BODY BGCOLOR=#FFFF00> <html><html><html><html><html><html><BODY BGCOLOR=#00FF00> <html><html><html><html><html><html><BODY BGCOLOR=#0000FF>"
mailsend (Person), ("Read This"), (msg1 + msg2 + msg3 + msg4 + msg5)
End Sub
Sub mail_flashminimize()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

X = ShowWindow(AOLChild, SW_MINIMIZE)
End Sub

Sub mail_flashrestore()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

X = ShowWindow(AOLChild, SW_RESTORE)
End Sub

Sub mail_forward(Person, message)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, i As Long, AOLEdit As Long, RICHCNTL As Long
'
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 3
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 0.5

AOLEdit = FindWindowEx(AOLChild, 0&, "_AOL_Edit", vbNullString)
RICHCNTL = FindWindowEx(AOLChild, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(AOLEdit, WM_SETTEXT, 0, Person)
DoEvents
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0, message)
End Sub

Function mail_lag(Person As String)
'prett lame to use but oh well just don't use it
'when i'm around
For i = 1 To 10000
a = a + "<html></html>"
b = b + "<html></html>"
c = c + "<html></html>"
D = D + "<html></html>"
Next
mailsend Person, "Chat Room Update", "HeHe Your Gonne Be Around Fer A while HeHe -MaceX Owns U" & a + b + c + D
End Function
Sub mail_center()
'this will always werk hehe
'Keyword "MailCenter"
' here's other code
popupicon "2", "M"
End Sub
Sub mail_massmail(List As listbox, subject, text As String)
List.Enabled = False
List.ListIndex = 0
For Mass = 0 To List.ListCount - 1
List.ListIndex = Mass
 mailsend (List.text), (subject), (text)
timeout 0.5
Next Mass
List.Enabled = True
End Sub
Sub mail_minimize()
Dim X, AOLChild As Long, MDIClient As Long, AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
X = ShowWindow(AOLChild&, SW_MINIMIZE)
End Sub
Sub mail_openflash()
popupicon2 "3", "D", "I"

End Sub

Sub mail_list(Person$, List As listbox)
' This will open a new email and then enter in all the items from
' the given listbox in the boddy of the message. for a server
For i = 0 To List.ListCount - 1
a = a & Chr$(13) & Chr$(10) & List.List(i)
Next i
Item = a
DoEvents
mailsend (Person$), ("Server List ∑ 1 - " & List.ListCount), (Item)
End Sub
Sub aol_setfont(font, size)
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
Dim AOLIcon As Long, i As Long, AOLFontCombo As Long
Dim AOLCombobox As Long
popupicon "5", "p"
timeout 0.7
AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 11
    AOLIcon = FindWindowEx(AOLChild, AOLIcon, "_AOL_Icon", vbNullString)
Next
SendMessage AOLIcon, &H201, 0&, 0&
SendMessage AOLIcon, &H202, 0&, 0&
timeout 1
AOLFontCombo = FindWindowEx(AOLChild, 0&, "_AOL_FontCombo", vbNullString)
AOLCombobox = FindWindowEx(AOLChild, 0&, "_AOL_Combobox", vbNullString)
SendMessage AOLFontCombo, &H201, 0&, 0&
timeout 1
SendMessageByString AOLFontCombo, WM_SETTEXT, 0&, font
timeout 1
SendMessage AOLCombobox, &H201, 0&, 0&
timeout 1
SendMessageByString AOLCombobox, WM_SETTEXT, 0&, size

End Sub

Function wait_formailload()
Dim AOLChild As Long, MDIClient As Long, AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
Do: DoEvents
Loop Until AOLChild& <> 0
End Function
Function mail_blackblue(Who, Subj, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(f, 0, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      mailsend (Who), (Subj), (msg)
End Function

Public Sub popupicon(IconNumber As Long, Character As String)
    Dim Message1 As Long, Message2 As Long, AOLFrame As Long
    Dim AOLToolbar As Long, Toolbar As Long, AOLIcon As Long
    Dim NextOfClass As Long, AscCharacter As Long
    Message1& = FindWindow("#32768", vbNullString)
    AOLFrame& = FindWindow("AOL Frame25", vbNullString)
    AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(AOLToolbar, 0&, "_AOL_Toolbar", vbNullString)
    AOLIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
    For NextOfClass& = 1 To IconNumber&
        AOLIcon& = GetWindow(AOLIcon&, 2)
    Next NextOfClass&
    Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        Message2& = FindWindow("#32768", vbNullString)
    Loop Until Message2& <> Message1&
    AscCharacter& = Asc(Character$)
    Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
End Sub
Function mail_blackblueblack(Who, Subj, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(f, 0, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      mailsend (Who), (Subj), (msg)
End Function
Function mail_blackgreen(Who, Subj, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, f, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     mailsend (Who), (Subj), (msg)
     
End Function

Function mail_blackgreenblack(Who, Subj, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(0, f, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     mailsend (Who), (Subj), (msg)
End Function
Function RGBtoHEX(RGB)
    a = Hex(RGB)
    b = Len(a)
    If b = 5 Then a = "0" & a
    If b = 4 Then a = "00" & a
    If b = 3 Then a = "000" & a
    If b = 2 Then a = "0000" & a
    If b = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function
Function rgb2hex(r, g, b)
    Dim X&
    Dim xx&
    Dim Color&
    Dim Divide
    Dim Answer&
    Dim Remainder&
    Dim Configuring$
    For X& = 1 To 3
        If X& = 1 Then Color& = b
        If X& = 2 Then Color& = g
        If X& = 3 Then Color& = r
        For xx& = 1 To 2
            Divide = Color& / 16
            Answer& = Int(Divide)
            Remainder& = (10000 * (Divide - Answer&)) / 625
            If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
            If Remainder& = 10 Then Configuring$ = "A" + Configuring$
            If Remainder& = 11 Then Configuring$ = "B" + Configuring$
            If Remainder& = 12 Then Configuring$ = "C" + Configuring$
            If Remainder& = 13 Then Configuring$ = "D" + Configuring$
            If Remainder& = 14 Then Configuring$ = "E" + Configuring$
            If Remainder& = 15 Then Configuring$ = "F" + Configuring$
            Color& = Answer&
        Next xx&
    Next X&
    Configuring$ = trimspaces(Configuring$)
    rgb2hex = Configuring$
End Function


Function mail_blackred(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(0, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     mailsend (Who), (Subj), (msg)
End Function
Function mail_blackredblack(Who, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, f, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    mailsend (Who), (Subj), (msg)
End Function
Function mail_blueblack(Who, Subj, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 0, 0)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
     mailsend (Who), (Subj), (msg)
End Function
Function mail_bluered(Who, Subj, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 255 / a
        f = e * b
        g = RGB(255 - f, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    mailsend (Who), (Subj), (msg)
End Function
Function mail_blueredblue(Who, Subj, txt)
a = Len(txt)
    For b = 1 To a
        c = Left(txt, b)
        D = Right(c, 1)
        e = 510 / a
        f = e * b
        If f > 255 Then f = (255 - (f - 255))
        g = RGB(255 - f, 0, f)
        H = RGBtoHEX(g)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
      mailsend (Who), (Subj), (msg)
End Function
Sub mail_opennew()
Dim AOLIcon As Long, AOLToolbar2 As Long, AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
ClickIcon (AOLIcon&)
End Sub
Sub mail_openold()
Dim AOLTabControl As Long, AOLChild As Long, MDIClient As Long
Dim AOLIcon As Long, AOLToolbar2 As Long, AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
ClickIcon (AOLIcon&)
timeout 1
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
AOLTabControl& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
ClickIcon (AOLTabControl&)
End Sub
Sub mail_opensent()
Dim AOLTabControl As Long, AOLChild As Long, MDIClient As Long
Dim AOLIcon As Long, AOLToolbar2 As Long, AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
ClickIcon (AOLIcon&)
timeout 1
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
AOLTabControl& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
ClickIcon (AOLTabControl&)
End Sub
Sub mail_restore()
Dim X, AOLChild As Long, MDIClient As Long, AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'sOnline Mailbox")
X = ShowWindow(AOLChild&, SW_RESTORE)
End Sub
Function mail_count()
Dim AoLTree As Long, AOLTabPage As Long, AOLFrame As Long
Dim AOLTabControl As Long, AOLChild As Long, MDIClient As Long
Dim AOLIcon As Long, AOLToolbar2 As Long, AOLToolbar As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
ClickIcon (AOLIcon&)
timeout 0.7
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", UserSN + "'s Online Mailbox")
AOLTabControl& = FindWindowEx(AOLChild&, 0&, "_AOL_TabControl", vbNullString)
AOLTabPage& = FindWindowEx(AOLTabControl&, 0&, "_AOL_TabPage", vbNullString)
AoLTree& = FindWindowEx(AOLTabPage&, 0&, "_AOL_Tree", vbNullString)
mail_count = SendMessage(AoLTree&, LB_GETCOUNT, 0, 0)
End Function
Public Function find_flashmailbox() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AoLTree As Long, AOLIcon As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AoLTree = FindWindowEx(AOLChild, 0&, "_AOL_Tree", vbNullString)
    AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
    If AoLTree <> 0& And AOLIcon <> 0& Then
        find_flashmailbox = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AoLTree = FindWindowEx(AOLChild, 0&, "_AOL_Tree", vbNullString)
            AOLIcon = FindWindowEx(AOLChild, 0&, "_AOL_Icon", vbNullString)
            If AoLTree <> 0& And AOLIcon <> 0& Then
                find_flashmailbox = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Function chat_sndis(name As String)
Call chatsend("ï^v∑ï Sn Disser ∑ï " & name)
timeout (0.5)
For X = 1 To Len(name)
lcse$ = LCase(name)
letr$ = Mid(lcse$, X, 1)
If letr$ = "a" Then Let dis$ = "ï^v∑ï a -is for the animals your momma fucks": GoTo Dissem
If letr$ = "b" Then Let dis$ = "ï^v∑ï b -is for all the boys you love": GoTo Dissem
If letr$ = "c" Then Let dis$ = "ï^v∑ï c -is for the cunt you are": GoTo Dissem
If letr$ = "d" Then Let dis$ = "ï^v∑ï d -is for all the times your dissed": GoTo Dissem
If letr$ = "e" Then Let dis$ = "ï^v∑ï e -is for that egghead of yours": GoTo Dissem
If letr$ = "f" Then Let dis$ = "ï^v∑ï f -is for the friday nights you stay home": GoTo Dissem
If letr$ = "g" Then Let dis$ = "ï^v∑ï g -is for the girls who hate you": GoTo Dissem
If letr$ = "h" Then Let dis$ = "ï^v∑ï h -is for the ho your momma is": GoTo Dissem
If letr$ = "i" Then Let dis$ = "ï^v∑ï i -is for the idiotic dumbass you are": GoTo Dissem
If letr$ = "j" Then Let dis$ = "ï^v∑ï j -is for all the times you jerkoff to your dog": GoTo Dissem
If letr$ = "k" Then Let dis$ = "ï^v∑ï k -is for you self esteem that the cool kids killed": GoTo Dissem
If letr$ = "l" Then Let dis$ = "ï^v∑ï l -is for the lame ass you are": GoTo Dissem
If letr$ = "m" Then Let dis$ = "ï^v∑ï m -is for the many men you sucked": GoTo Dissem
If letr$ = "n" Then Let dis$ = "ï^v∑ï n -is for the nights you spent alone": GoTo Dissem
If letr$ = "o" Then Let dis$ = "ï^v∑ï o -is for the sex operation you had": GoTo Dissem
If letr$ = "p" Then Let dis$ = "ï^v∑ï p -is for the times people p on you": GoTo Dissem
If letr$ = "q" Then Let dis$ = "ï^v∑ï q -is for the queer you are": GoTo Dissem
If letr$ = "r" Then Let dis$ = "ï^v∑ï r -is for all the times i raped your sister": GoTo Dissem
If letr$ = "s" Then Let dis$ = "ï^v∑ï s -is for your lover Steve Case": GoTo Dissem
If letr$ = "t" Then Let dis$ = "ï^v∑ï t -is for the tits youll never see": GoTo Dissem
If letr$ = "u" Then Let dis$ = "ï^v∑ï u -is for your underwear hangin on the flagpole": GoTo Dissem
If letr$ = "v" Then Let dis$ = "ï^v∑ï v -is for the victories you'll never have": GoTo Dissem
If letr$ = "w" Then Let dis$ = "ï^v∑ï w -is for the 400 pounds you wiegh":  GoTo Dissem
If letr$ = "x" Then Let dis$ = "ï^v∑ï x -is for all the lamers who" & Chr(34) & "[x]'ed" & Chr(34) & " you online": GoTo Dissem
If letr$ = "y" Then Let dis$ = "ï^v∑ï y -is for the question of, y your even alive?": GoTo Dissem
If letr$ = "z" Then Let dis$ = "ï^v∑ï z -is for zits thast are all over your face":  GoTo Dissem

If letr$ = "1" Then Let dis$ = "ï^v∑ï 1 -is for the amount of girls you'll ever get": GoTo Dissem
If letr$ = "2" Then Let dis$ = "ï^v∑ï 2 -is for the 2 dollars you make an hour": GoTo Dissem
If letr$ = "3" Then Let dis$ = "ï^v∑ï 3 -is for the amount of men your girl takes at once": GoTo Dissem
If letr$ = "4" Then Let dis$ = "ï^v∑ï 4 -is for your mom bein a whore":  GoTo Dissem
If letr$ = "5" Then Let dis$ = "ï^v∑ï 5 -is for 5 times an hour you whack off": GoTo Dissem
If letr$ = "6" Then Let dis$ = "ï^v∑ï 6 -is for the years you been single": GoTo Dissem
If letr$ = "7" Then Let dis$ = "ï^v∑ï 7 -is for the times your girl cheated on you..with me": GoTo Dissem
If letr$ = "8" Then Let dis$ = "ï^v∑ï 8 -is for how many people beat the hell outta you today": GoTo Dissem
If letr$ = "9" Then Let dis$ = "ï^v∑ï 9 -is for how many boyfriends your momma has": GoTo Dissem
If letr$ = "0" Then Let dis$ = "ï^v∑ï 0 -is for the amount of girls you get": GoTo Dissem
Dissem:
Call chatsend(dis$)
timeout (1)
Next X

End Function




Sub timeout(HowLong)
Dur = Timer
Do While Timer - Dur < HowLong
DoEvents
Loop
End Sub



Sub keyword(place$)
Dim edit As Long, AOLCombobox As Long, AOLToolbar2 As Long
Dim AOLToolbar As Long, AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLCombobox& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Combobox", vbNullString)
edit& = FindWindowEx(AOLCombobox&, 0&, "Edit", vbNullString)
SendMessageByString edit&, WM_SETTEXT, 0, place$
SendMessageByNum edit&, WM_CHAR, Key_Space, 0
SendMessageByNum edit&, WM_CHAR, Key_Enter, 0
End Sub


Public Sub im(name$, Word$)
'this im will stay open
Dim AOLIcon As Long, RICHCNTL As Long, AOLChild As Long
Dim MDIClient As Long, AOLFrame As Long
Call keyword("aol://9293:" & name$)
timeout 2
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 8
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0, Word$)

SendMessage AOLIcon&, &H201, 0&, 0&
SendMessage AOLIcon&, &H202, 0&, 0&
End Sub
Sub wel_hide()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

ShowWindow AOLChild, 0
End Sub


Sub wel_show()
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

ShowWindow AOLChild, 1
End Sub
Sub bot_echo(ScreenName)
'You can use a textbox or just a screen name
'Ex Bot_Echo (sn here)
If chat_lastsn Like (ScreenName) Then
chatsend (chat_lastlinemsg)
End If
End Sub
Sub aol_capchange(capt As String)
'this changes the aol caption to what you want
Dim AOLChild As Long
AOLChild = FindWindow("AOL Frame25", vbNullString)
SendMessageByString AOLChild, &HC, 0&, capt
End Sub
Public Sub aol_45kill()
  'clicks 45 minute window
  Dim win As Long, Button As Long
  win = FindWindow("_AOL_Palette", vbNullString)
  Button = FindWindowEx(win, 0&, "_AOL_Icon", vbNullString)
  ClickIcon (Button)
End Sub
Sub aol_idle()
'will click the inactivity Window
'Code: Call aol_idle
Dim Button2 As Long, Modal As Long, Button As Long
Dim Palette As Long
Palette = FindWindow("_AOL_Palette", vbNullString)
Button = FindWindowEx(Palette&, 0&, "_AOL_Icon", vbNullString)
ClickIcon (Button)
Modal = FindWindow("_AOL_Modal", vbNullString)
Button2 = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
ClickIcon (Button2)
End Sub
Sub favorite_addtopwin()
popupicon "6", "a"

End Sub
Public Function find_dlwin() As Long
    Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long
    Dim AOLStatic As Long, AOLGauge As Long, AOLCheckbox As Long
    Dim AOLButton As Long

    AOLFrame = FindWindow("AOL Frame25", vbNullString)
    MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)
    AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)

    AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
    AOLGauge = FindWindowEx(AOLChild, 0&, "_AOL_Gauge", vbNullString)
    AOLCheckbox = FindWindowEx(AOLChild, 0&, "_AOL_Checkbox", vbNullString)
    AOLButton = FindWindowEx(AOLChild, 0&, "_AOL_Button", vbNullString)
    If AOLStatic <> 0& And AOLGauge <> 0& And AOLCheckbox <> 0& And AOLButton <> 0& Then
        find_dlwin = AOLChild
        Exit Function
    Else
        While AOLChild
            AOLChild = GetWindow(AOLChild, 2)
            AOLStatic = FindWindowEx(AOLChild, 0&, "_AOL_Static", vbNullString)
            AOLGauge = FindWindowEx(AOLChild, 0&, "_AOL_Gauge", vbNullString)
            AOLCheckbox = FindWindowEx(AOLChild, 0&, "_AOL_Checkbox", vbNullString)
            AOLButton = FindWindowEx(AOLChild, 0&, "_AOL_Button", vbNullString)
            If AOLStatic <> 0& And AOLGauge <> 0& And AOLCheckbox <> 0& And AOLButton <> 0& Then
                find_dlwin = AOLChild
                Exit Function
            End If
        Wend
    End If
End Function
Sub chat_capchange(Caption As String)
'this changes the chat's caption
Dim AOLFrame As Long, MDIClient As Long, AOLChild As Long

AOLFrame = FindWindow("AOL Frame25", vbNullString)
MDIClient = FindWindowEx(AOLFrame, 0&, "MDIClient", vbNullString)

AOLChild = FindWindowEx(MDIClient, 0&, "AOL Child", vbNullString)
SendMessageByString AOLChild, &HC, 0&, Caption
End Sub
Public Sub bot_food()
Dim rand
Dim peep As String
If chat_lastlinemsg = "/food" Then
peep$ = ChatLastSN
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call chatsend("" & peep$ & ", You Get a Hamburger!")
If rand = 2 Then Call chatsend("" & peep$ & ", You Get a Nice Banna!")
If rand = 3 Then Call chatsend("" & peep$ & ", You Get a Hot Dog!")
If rand = 4 Then Call chatsend("" & peep$ & ", You Get a Steak!")
If rand = 5 Then Call chatsend("" & peep$ & ", You Get a rotton Apple")
If rand = 6 Then Call chatsend("" & peep$ & ", You Get an Orange")
If rand = 7 Then Call chatsend("" & peep$ & ", You Get a Chicken")
If rand = 8 Then Call chatsend("" & peep$ & ", You Get a Turkey")
If rand = 9 Then Call chatsend("" & peep$ & ", You Get Nothing Fat Ass!")
If rand = 10 Then Call chatsend("" & peep$ & ", You Get A Pig And Become a Fat Ass!")
Call timeout(0.6)
End If
End Sub
Public Sub bot_fortune()
Dim rand
Dim peep As String
If chat_lastlinemsg = "/fortune" Then
peep$ = ChatLastSN
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call chatsend("" & peep$ & ", You will win the lottery and spend it all on BEER!")
If rand = 2 Then Call chatsend("" & peep$ & ", You will kill Steve Case and take over AoL!")
If rand = 3 Then Call chatsend("" & peep$ & ", You will marry Carmen Electra!")
If rand = 4 Then Call chatsend("" & peep$ & ", You will DL a PWS and get thousands of bucks charged on your account!")
If rand = 5 Then Call chatsend("" & peep$ & ", You will end up werking at McDonalds and die a lonely man")
If rand = 6 Then Call chatsend("" & peep$ & ", You will get a check for ONE MILLION $$ from me! Yeah right!")
If rand = 7 Then Call chatsend("" & peep$ & ", You will be OWNED by MaceX")
If rand = 8 Then Call chatsend("" & peep$ & ", You will be OWNED by UqG")
If rand = 9 Then Call chatsend("" & peep$ & ", You will get an OH and delete Steve Case's SN!")
If rand = 10 Then Call chatsend("" & peep$ & ", You will slip on a banana peel in Japan and land on some egg foo yung!")
Call timeout(0.6)
End If
End Sub
Sub bot_idle(Label, Status As Boolean)
'same things
If Status = True Then
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix32 ∑ï Idle on "
Do: DoEvents
Call aol_idle
Label.Caption = Val(Label) + 1
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix32 ∑ï Idle ∑ï  " + Label
timeout (60) 'Makes it wait 1 minute
Loop
Else
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix32 ∑ï Idle off "
Label = "0"
Exit Sub
End If
End Sub

Sub bot_request(Wat)
'You can use a textbox
'put in timer
'Ex Bot_Request (wat)
'chatsend "<font face=Tahoma><font color=#000099>∑ï∑õ Request Bot - [ " + Wat + " ]"
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix32 ∑ï Reqesting ∑ï " + Wat
timeout 0.3
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Type ∑ï /I got it"

If chat_lastlinemsg = "/I got It" Then
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>" + chat_lastsn + " ∑ï can u send?"
timeout 0.2
End If
End Sub
Sub bot_afk(msg, Status As Boolean)
'In a button put Call BotAfk(msg,label,true)
'the True means to turn it on
'to turn off use this Call BotAfk(msg,label,false)
If Status = True Then
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix32 ∑ï Afk on "
Do: DoEvents
timeout 60
Label.Caption = Val(Label) + 1
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix32∑ï Afk "
timeout 0.1
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Reason ∑ï " + msg
Loop
Else
chatsend "</b><font color=#ff0000>ï^v∑ï <font color=#000000>Dimix32 ∑ï Afk Off"
'chatsend "<font face=""Tahoma""><font color=#000099>óã°!õï Afk Bot  - [ off ] "
Label = "0"
Exit Sub
End If
'∏.∑¥Ø`∑.∏ ∏.∑¥Ø`∑.∏
'`∑.∏_∏.∑¥ `∑.∏_∏.∑¥
End Sub


Public Sub bot_8Ball()
'By MaceX
Dim rand, peep
If LCase(chat_lastlinemsg) = ("/8Ball") Then
DoEvents
Randomize
rand = Int((Rnd * 10) + 1)
If rand = 1 Then Call chatsend("Answer:  Hell Yeah!")
If rand = 2 Then Call chatsend("Answer:  Hell No!")
If rand = 3 Then Call chatsend("Answer:  Fuck No!")
If rand = 4 Then Call chatsend("Answer:  Fuck Yes!")
If rand = 5 Then Call chatsend("Answer:  No Way")
If rand = 6 Then Call chatsend("Answer:  Yes Way")
If rand = 7 Then Call chatsend("Answer:  No")
If rand = 8 Then Call chatsend("Answer:  Yes")
If rand = 9 Then Call chatsend("Answer:  Ask Me L8er")
If rand = 10 Then Call chatsend("Answer: Go Away Lamer!")
Call timeout(0.6)
End If
End Sub

