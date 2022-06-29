Attribute VB_Name = "e2AOL7"
'e2 module for AOL7.0 - blike
'first version released
'december 24th, 2001
'some subs from DOS32, Shaggy
'check subs for indiviual authors

'If you recieve this module without the updater
'please visit http://www.blike.com/e2mod/
'to download the zip with updater included.


Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



' API constants...

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186

Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_SETTEXT = &HC
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101

Public Const VK_SPACE = &H20

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Function find_RoomList()
Dim lngAOLChild As Long
Dim lngAOLListbox As Long
lngAOLChild = find_Room
lngAOLListbox = FindWindowEx(lngAOLChild, 0, "_AOL_Listbox", vbNullString)
find_RoomList = lngAOLListbox
End Function
Public Function find_UserInfo(screen_name As String)
'gets the handle of the window
'with user info (for ignore)
Dim lngAOLFrame As Long
Dim lngMDIClient As Long
Dim lngAOLChild As Long
lngAOLFrame = FindWindow("AOL Frame25", vbNullString)
lngMDIClient = FindWindowEx(lngAOLFrame, 0, "MDIClient", vbNullString)
lngAOLChild = FindWindowEx(lngMDIClient, 0, "AOL Child", screen_name)
find_UserInfo = lngAOLChild
End Function

Public Sub form_Move(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Function chat_Caption()
chat_Caption = get_Caption(find_Room)
End Function

Public Function find_IMWindow()
Dim lngAOLFrame As Long
Dim lngMDIClient As Long
Dim lngAOLChild As Long
Dim lngRICHCNTL As Long
lngAOLFrame = FindWindow("AOL Frame25", vbNullString)
lngMDIClient = FindWindowEx(lngAOLFrame, 0, "MDIClient", vbNullString)
lngAOLChild = FindWindowEx(lngMDIClient, 0, "AOL Child", "Send Instant Message")
find_IMWindow = lngAOLChild
End Function

Public Function find_MailWin() As Long

Dim lngAOLChild As Long

lngAOLChild = FindWindowEx(find_MDI, 0, "AOL Child", "Write Mail")

find_MailWin = lngAOLChild
End Function

Public Function find_Welcome()

    Dim aol As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    aol& = find_AOL
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = get_Caption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
    find_Welcome = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = get_Caption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
            find_Welcome = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
find_Welcome = 0
End Function

Public Sub hide_Welcome()
Call window_Hide(find_Welcome)
End Sub
Public Sub chat_IgnoreByIndex(intIndex As Integer)
Dim lngAOLChild As Long
Dim lngAOLListbox As Long
Dim lngAOLCheckbox As Long

If find_RoomList = 0 Then
    MsgBox "You must be in a chatroom!"
    Exit Sub
End If

lngAOLChild = find_Room
lngAOLListbox = FindWindowEx(lngAOLChild, 0, "_AOL_Listbox", vbNullString)
' number 1 is the index number that will be selecting in the listbox.
Call SendMessage(lngAOLListbox, LB_SETCURSEL, intIndex, 0)
Call PostMessage(lngAOLListbox, WM_LBUTTONDBLCLK, 0, 0)

Do
DoEvents
Loop Until find_UserInfo("XDeepArcticX") <> 0

lngAOLCheckbox = FindWindowEx(find_UserInfo("XDeepArcticX"), 0, "_AOL_Checkbox", vbNullString)
MsgBox lngAOLCheckbox
End Sub

Public Sub show_Welcome()
Call window_Show(find_Welcome)
End Sub
Public Sub INI_Write(Section As String, Key As String, KeyValue As String, Directory As String)
Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub


Public Function INI_Read(Section As String, Key As String, Directory As String) As String
Dim strBuffer As String
strBuffer = String(750, Chr(0))
Key$ = LCase$(Key$)
INI_Read = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Public Sub chat_Clear()
'contributed by: blike (blikeism@hotmail.com)

lngAOLChild = find_Room
lngRICHCNTLREADONLY = FindWindowEx(lngAOLChild, 0, "RICHCNTLREADONLY", vbNullString)
' This will clear all the text out of the window

Call SendMessageByString(lngRICHCNTLREADONLY, WM_SETTEXT, 0&, "")

End Sub

Public Function find_AOL() As Long
Dim lngAOLFrame As Long
lngAOLFrame = FindWindow("AOL Frame25", vbNullString)
find_AOL = lngAOLFrame
End Function


Public Function find_Room() As Long
    Dim aol As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        find_Room& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                find_Room& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    find_Room& = child&
End Function
Public Function find_MDI() As Long
Dim lngAOLFrame As Long
Dim lngMDIClient As Long
lngAOLFrame = find_AOL
lngMDIClient = FindWindowEx(lngAOLFrame, 0, "MDIClient", vbNullString)

find_MDI = lngMDIClient
End Function


Public Function get_Chat() As String
Dim lngRICHCNTLREADONLY As Long
Dim lngLen As Long
Dim strText As String

lngRICHCNTLREADONLY = FindWindowEx(find_Room, 0, "RICHCNTLREADONLY", vbNullString)
' This will return the text from the window.
lngLen = SendMessage(lngRICHCNTLREADONLY, WM_GETTEXTLENGTH, 0, 0) + 1
strText = String(lngLen, vbNullChar)
Call SendMessage(lngRICHCNTLREADONLY, WM_GETTEXT, lngLen, ByVal strText)

get_Chat = strText

End Function
Public Function get_ChatLastLine()
Dim chatText As String
Dim chatLines As Integer

chatText = get_ChatText

chatText = chatText & Chr(13)

chatLines = textCountLines(chatText)

get_ChatLastLine = textLine(chatText, chatLines)
End Function

Public Sub mail_OpenNewBox()
Dim lngAOLIcon As Long
Dim lngAOLChild As Long
Dim lngRecipient As Long, lngSubject As Long
Dim lngMessage As Long

lngAOLFrame = find_AOL
lngAOLToolbar = FindWindowEx(lngAOLFrame, 0, "AOL Toolbar", vbNullString)
lngAOLToolbar = FindWindowEx(lngAOLToolbar, 0, "_AOL_Toolbar", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, lngAOLIcon, "_AOL_Icon", vbNullString)

Call PressIcon(lngAOLIcon)
End Sub

Public Function get_User() As String
'modified from dos32
    Dim aol As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    aol& = find_AOL
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = get_Caption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        get_User$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = get_Caption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                get_User$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    get_User$ = ""
End Function
Public Function Mail_SendButton() As Long
'Contributed by: Runaz (runaz@jialz.net)

Dim lngAOLChild As Long, lngAOLIcon As Long

lngAOLChild = find_MailWin
lngAOLIcon = FindWindowEx(lngAOLChild, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
Mail_SendButton = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
End Function

Sub stayONTOP(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Function Open_Chat_PeopleConnection()
Call clickToolbar("3", "C")
End Function
Public Function OpenGetMemberProfile()
Call clickToolbar("3", "G")
End Function
Public Function OpenFindAChat()
Call clickToolbar("3", "F")
End Function
Public Function OpenCreateHomePage()
Call clickToolbar("3", "H")
End Function
Public Function OpenStartYourOwnChat()
Call clickToolbar("3", "S")
End Function
Public Function OpenJoinOnlineGroup()
Call clickToolbar("3", "J")
End Function
Public Function OpenLiveEvents()
Call clickToolbar("3", "E")
End Function
Public Function OpenSignOnAFriend()
Call clickToolbar("3", "O")
End Function
Public Function OpenBuddylist()
Call clickToolbar("3", "B")
End Function
Public Function OpenInvitations()
Call clickToolbar("3", "V")
End Function
Public Function OpenLocateMemOnline()
Call clickToolbar("3", "L")
End Function
Public Function OpenMemberDirectory()
Call clickToolbar("3", "N")
End Function
Public Function OpenMessage2Pager()
Call clickToolbar("3", "M")
End Function
Public Function OpenPersonals()
Call clickToolbar("3", "P")
End Function
Public Function OpenWhitePages()
Call clickToolbar("3", "W")
End Function

Public Function opentvkistings()
Call clickToolbar("6", "T")
End Function
Public Function openshopataol()
Call clickToolbar("6", "S")
End Function
Public Function Add_To_Calander()
'contributed by: lead (lelead@hotmail.com)
Call clickToolbar("6", "A")
End Function
Public Function OpenCalender()
Call clickToolbar("6", "C")
End Function
Public Function openCarBuying()
Call clickToolbar("6", "B")
End Function
Public Function openDownloadcenter()
Call clickToolbar("6", "D")
End Function
Public Function openHomeWorkHelp()
Call clickToolbar("6", "K")
End Function
Public Function openMapsnDirections()
Call clickToolbar("6", "M")
End Function
Public Function openGovermentGuide()
Call clickToolbar("6", "U")
End Function
Public Function openMedicalReferences()
Call clickToolbar("6", "N")
End Function
Public Function openMovieShowtimes()
Call clickToolbar("6", "W")
End Function
Public Function OpenPersonals2()
Call clickToolbar("6", "P")
End Function
Public Function openRadio()
Call clickToolbar("6", "R")
End Function
Public Function openRecipeFinder()
Call clickToolbar("6", "F")
End Function
Public Function openSportsScores()
Call clickToolbar("6", "O")
End Function
Public Function openStockPortfolios()
Call clickToolbar("6", "L")
End Function
Public Function openStockQuotes()
Call clickToolbar("6", "Q")
End Function
Public Function openTravelReservations()
Call clickToolbar("6", "V")
End Function
Public Function openYellowPages()
Call clickToolbar("6", "E")
End Function
Public Function openYouveGotPictures()
Call clickToolbar("6", "Y")
End Function

Public Function OpenFavorites()
Call clickToolbar("11", "F")
End Function

Public Sub Playwav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub
Public Function ADD_Top_Window_to_Favorites()
'contributed by: Anonymous
Call clickToolbar("11", "A")
End Function
Public Function keyword_Open()
Call clickToolbar("11", "G")
End Function
Public Function My_Hot_Keys()
Call clickToolbar2("11", "M", "E")
End Function

Public Sub Mail_New()
Dim lngAOLIcon As Long

lngAOLFrame = find_AOL
lngAOLToolbar = FindWindowEx(lngAOLFrame, 0, "AOL Toolbar", vbNullString)
lngAOLToolbar = FindWindowEx(lngAOLToolbar, 0, "_AOL_Toolbar", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, lngAOLIcon, "_AOL_Icon", vbNullString)

Call PressIcon(lngAOLIcon)

End Sub
Public Sub Mail_Send(Recipient As String, subject As String, message As String)
'Contributed by: lead (lelead@hotmail.com)

Dim lngAOLIcon As Long
Dim lngAOLChild As Long
Dim lngRecipient As Long, lngSubject As Long
Dim lngMessage As Long

lngAOLFrame = find_AOL
lngAOLToolbar = FindWindowEx(lngAOLFrame, 0, "AOL Toolbar", vbNullString)
lngAOLToolbar = FindWindowEx(lngAOLToolbar, 0, "_AOL_Toolbar", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLToolbar, lngAOLIcon, "_AOL_Icon", vbNullString)

Call PressIcon(lngAOLIcon)

Do:
DoEvents
Loop Until find_MailWin <> 0
Pause (0.1)
lngRecipient = FindWindowEx(find_MailWin, 0, "_AOL_Edit", vbNullString)
lngSubject = FindWindowEx(find_MailWin, lngRecipient, "_AOL_Edit", vbNullString)
lngSubject = FindWindowEx(find_MailWin, lngSubject, "_AOL_Edit", vbNullString)
lngMessage = FindWindowEx(find_MailWin, 0, "RICHCNTL", vbNullString)

Call SendMessage(lngRecipient, WM_SETTEXT, 0, ByVal Recipient)
Call SendMessage(lngSubject, WM_SETTEXT, 0, ByVal subject)
Call SendMessage(lngMessage, WM_SETTEXT, 0, ByVal message)

Call PressIcon(Mail_SendButton)
End Sub
Public Sub PressIcon(AOLIcon As Long)
    Call SendMessage(AOLIcon, WM_LBUTTONDOWN, 0&, 0&)
    'This sends a message to push the icon DOWN
    Call SendMessage(AOLIcon, WM_LBUTTONUP, 0&, 0&)
    'This sends a message to release it

    Call SendMessage(AOLIcon, WM_KEYDOWN, VK_SPACE, 0)
    'presses space on the keyboard on the icon
    'sometimes AOL is bitchy, this is incase the above
    'doesn't work
    Call SendMessage(AOLIcon, WM_KEYUP, VK_SPACE, 0)
    
End Sub

Public Function chat_LineMsg(ChatLine As String) As String
'Seperates ChatMsg from ChatSn
'Example:
'Dim SayWhat As String, ChatText As String
'ChatText = Text1.Text
'SayWhat$ = ChatLineMsg(ChatText$)

If InStr(ChatLine, Chr(9)) = 0 Then
ChatLineMsg = ""
Exit Function
End If
chat_LineMsg = Right(ChatLine, Len(ChatLine) - InStr(ChatLine, Chr(9)))
End Function

Public Function chat_LineSN(ChtLine As String) As String
'Seperates ChatMsg from ChatSn
'Example:
'Dim SN As String, ChatText As String
'ChatText = Text1.Text
'SN$ = ChatLineMsg(ChatText$)
If InStr(ChtLine, ":") = 0 Then
ChatLineSN = ""
Exit Function
End If
chat_LineSN = Left(ChtLine, InStr(ChtLine, ":") - 1)
End Function
Public Function get_ChatText() As String
'gets the chattext without Link of X of Y
'formats chattext for use with textbox
Dim strChat As String
strChat = get_Chat
strChat = ReplaceString(strChat, Chr(13), vbNewLine)

If InStr(strChat, "Link") Then
    strChat = Left(strChat, InStr(strChat, "Link") - 1)
End If

get_ChatText = strChat
End Function

Public Function Pause(Time As Long)
'Call Pause(1)
'duration is in seconds
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Time
        DoEvents
    Loop
End Function
Function textCountLines(Text)
theview$ = Text
Dim c As Integer
c = 0
For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
c = c + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
'If theline = c Then GoTo ex
thechars$ = ""
End If

Next FindChar
textCountLines = c

End Function

Function textLine(Text, theline)
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
'thechatext$ = Mid$(thechatext$, InStr(thechattext$, Chr(13)) + 1)
'If InStr(thechatext$, Chr(10)) Then
'thechatext$ = Mid$(thechatext$, InStr(thechattext$, Chr(10)) + 1)
'End If

textLine = thechatext$


End Function


Function ReplaceText(Text, charfind, charchange)
Dim replace As Integer

If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For replace = 1 To Len(Text)
thechar$ = Mid(Text, replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next replace

ReplaceText = thechars$

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


Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function ReverseText(Text)
For words = Len(Text) To 1 Step -1
ReverseText = ReverseText & Mid(Text, words, 1)
Next words
End Function



Sub form_NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub




Sub Text_Save(txtSave, path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave
    Open path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

Function Text_Load(path As String)
    Dim TextString As String
    On Error Resume Next
    Open path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    Text_Load = TextString$
End Function
Function KTEncrypt(ByVal password, ByVal strng, force)
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
  If force = 0 Then

    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)

    If InStr(chk$, Chr$(1) & "EX" & Chr$(1) & "X" & Chr$(1)) Then
    
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
      EncryptFlag = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag = True

    End If
  Else
    'force flag set, ecrypt string regardless of tag
    EncryptFlag = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(password)
  
  
  'Tack on leading characters to prevent repetative recognition
  password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
  password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
  password = password + Chr$(Asc(Right$(password, 1)) Xor PassMax)
  password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag = True Then
    strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
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
  If EncryptFlag = True Then
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
    strng = Chr$(1) + "EX" + Chr$(1) + strng + Chr$(1) + "EX" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
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
Public Sub Form_Center(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub

Function Text_Scramble(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Scrambles the text
For scrambling = 1 To Len(TheText)
thechar$ = Mid(TheText, scrambling, 1)
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
Text_Scramble = scrambled$

Exit Function
End Function

Function Text_Descramble(TheText)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(TheText, Len(TheText), 1)

If Not findlastspace = " " Then
TheText = TheText & " "
Else
TheText = TheText
End If

'Descrambles the text
For scrambling = 1 To Len(TheText)
thechar$ = Mid(TheText, scrambling, 1)
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
Text_Descramble = scrambled$

End Function



Function file_Exists(ByVal sFileName As String) As Integer
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        file_Exists = False
        Else
            file_Exists = True
    End If

End Function

Public Sub clickToolbar(IconNumber&, letter$)

Dim aolframe As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim AOLIcon As Long
Dim Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(clickToolbar2, AOLIcon, "_AOL_Icon", vbNullString)
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
Public Sub clickToolbar2(IconNumber&, letter$, letter2$)

Dim aolframe As Long
Dim menu As Long
Dim clickToolbar1 As Long
Dim clickToolbar2 As Long
Dim AOLIcon As Long
Dim Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
clickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon = FindWindowEx(clickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
AOLIcon = FindWindowEx(clickToolbar2, AOLIcon, "_AOL_Icon", vbNullString)
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


Function free_Process()
Do: DoEvents
process = process + 1
If process = 50 Then Exit Do
Loop
'frees process of freezes in your program
'and other stuff that makes your program
'slow down.  Works great.

End Function
Function get_Caption(hWnd)
hwndLength = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength, 0)
a = GetWindowText(hWnd, hwndTitle$, (hwndLength + 1))
get_Caption = hwndTitle$
End Function
Public Sub chat_Send(word As String)
'Call chatSend("hax")
Dim lngRICHCNTL As Long
Dim lngAOLIcon As Long
lngRICHCNTL = FindWindowEx(find_Room, 0, "RICHCNTL", vbNullString)
lngAOLIcon = FindWindowEx(find_Room, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(find_Room, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(find_Room, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(find_Room, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(find_Room, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(find_Room, lngAOLIcon, "_AOL_Icon", vbNullString)

Call SendMessage(lngRICHCNTL, WM_SETTEXT, 0, ByVal word)

Call PressIcon(lngAOLIcon)
End Sub
Public Sub IM_Send(Recipient As String, message As String)
'sends instant message
'using buddy list button
Dim lngAOLFrame25 As Long
Dim lngMDIClient As Long
Dim lngAOLChild As Long
Dim lngAOLIcon As Long, findIMbutton As Long
Dim lngRecipient As Long, lngMessage As Long
Dim lngSendButton As Long

lngAOLFrame25 = find_AOL
lngMDIClient = FindWindowEx(lngAOLFrame25, 0, "MDIClient", vbNullString)
lngAOLChild = FindWindowEx(lngMDIClient, 0, "AOL Child", "Buddy List")
lngAOLIcon = FindWindowEx(lngAOLChild, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
findIMbutton = lngAOLIcon

Call PressIcon(findIMbutton)

Do
DoEvents
Loop Until find_IMWindow <> 0

lngRecipient = FindWindowEx(find_IMWindow, 0, "_AOL_Edit", vbNullString)
lngMessage = FindWindowEx(find_IMWindow, 0, "RICHCNTL", vbNullString)

Call SendMessage(lngRecipient, WM_SETTEXT, 0, ByVal Recipient)
Call SendMessage(lngMessage, WM_SETTEXT, 0, ByVal message)

lngAOLChild = find_IMWindow
lngAOLIcon = FindWindowEx(lngAOLChild, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngSendButton = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)

Call PressIcon(lngSendButton)

End Sub

Public Sub IM_SendMenu(Recipient As String, message As String)
'sends instant message using the menu
'instead of buddy list
Dim lngAOLFrame25 As Long
Dim lngMDIClient As Long
Dim lngAOLChild As Long
Dim lngAOLIcon As Long, findIMbutton As Long
Dim lngRecipient As Long, lngMessage As Long
Dim lngSendButton As Long


Call clickToolbar("3", "i")


Do
DoEvents
Loop Until find_IMWindow <> 0

lngRecipient = FindWindowEx(find_IMWindow, 0, "_AOL_Edit", vbNullString)
lngMessage = FindWindowEx(find_IMWindow, 0, "RICHCNTL", vbNullString)

Call SendMessage(lngRecipient, WM_SETTEXT, 0, ByVal Recipient)
Call SendMessage(lngMessage, WM_SETTEXT, 0, ByVal message)

lngAOLChild = find_IMWindow
lngAOLIcon = FindWindowEx(lngAOLChild, 0, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngAOLIcon = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)
lngSendButton = FindWindowEx(lngAOLChild, lngAOLIcon, "_AOL_Icon", vbNullString)

Call PressIcon(lngSendButton)

End Sub

Public Sub List_Save(Directory As String, TheList As ListBox)
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub List_Load(Directory As String, TheList As ListBox)
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
Public Function zVers() As String
'this sub is critical for updater to function properly
'zVers = {13}
End Function


Public Sub window_Hide(hWnd As Long)
'dos32
    Call ShowWindow(hWnd&, SW_HIDE)
End Sub

Public Sub window_Show(hWnd As Long)
    Call ShowWindow(hWnd&, SW_SHOW)
End Sub


