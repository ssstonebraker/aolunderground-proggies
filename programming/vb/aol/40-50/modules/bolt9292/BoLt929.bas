Attribute VB_Name = "BoLt929"
'ßØL†929 version 3.5
'You can also use the public const on the bottom for colors instead of typing the Anyonning color code
'for the fadeforms like fadeform or iceform u always have to have "CALL" before it
'This is for Mostly AOL 4.0 and AIM
'Created by: Uyhs
'E-mail:Oddish0923@aol.com
'Website:   Http://uyhs.cjb.net
'I took some subs from Jaguar32 and NeewPsyche becuz that is the best bas out there!
'I must thank Mokefade for the great Sub's, and Function's
'Aim Screen Name:Bolt50432
'The Color Subs are suppose to make it easier instead of entering the color code
'The Aim Send Chat is AIMCS so its easier than typing
'ChatSend or Chat_send

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Private Const SPI_SCREENSAVERRUNNING = 97

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
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
   right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type


Public Const red = &HFF&
Public Const green = &HFF00&
Public Const blue = &HFF0000
Public Const yellow = &HFFFF&
Public Const white = &HFFFFFF
Public Const black = &H0&
Public Const purple = &HFF00FF
Public Const grey = &HC0C0C0
Public Const pink = &HC0C0FF
Public Const TURQUOISE = &HC0C000
Public Const SEAGREEN = &H80FF80
Public Const LBLUE = &HFFFFC0
Public Const LGREEN = &HFF00&
Public Const brown = &H4080&
Public Const DGREEN = &H8000&
Public Const NAVY = &H800000
Public Const Gold = &H8080&
Public Const BLUEPRPL = &HDA2C68
Public Const YELGRN = &H5EF7B6
Public Const MAGENTA = &H640DE8
Public Const MAROON = &H291F76
Public Const orange = &H80FF&


Type COLORRGB
  red As Long
  green As Long
  blue As Long
End Type





Sub FormOnTop(theform As Form)
SetWinOnTop = SetWindowPos(theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
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
Sub DataSaveList(path As String, Lst As ListBox)
'Ex: Call Save_ListBox("c:\windows\desktop\list.lst", list1)

    Dim Listz As Long
    On Error Resume Next

    Open path$ For Output As #1
    For Listz& = 0 To Lst.ListCount - 1
        Print #1, Lst.List(Listz&)
        Next Listz&
    Close #1
End Sub
Sub DataSaveCaption(path As String, lbl As Object)
free = FreeFile

Open path For Output As free
Print #free, lbl.Caption
Close #free
End Sub
Sub DataSaveEnabled(path As String, lbl As Object)
free = FreeFile

Open path For Output As free
Print #free, lbl.Enabled
Close #free
End Sub
Sub SndClick()
Plywav ("C:\WINDOWS\MEDIA\Utopia Close.wav")

End Sub
Sub SndDing()
Plywav ("C:\WINDOWS\MEDIA\logoff.wav")
End Sub
Sub SndTada()
Plywav ("C:\WINDOWS\MEDIA\Tada.wav")
End Sub
Sub SndExcla()
Plywav ("C:\WINDOWS\MEDIA/Utopia Exclamation.wav")
End Sub
Sub SndStartup()
Plywav ("C:\WINDOWS\MEDIA\utopia windows start.wav")
End Sub
Sub SndStartUp2()
Plywav ("C:\WINDOWS\MEDIA\The Microsoft Sound.wav")
End Sub

Sub DataLoadCaption(path As String, lbl As Object)
On Error GoTo errhand
free = FreeFile
Open path For Input As free
Do While Not EOF(free)
Line Input #free, Blah
lbl.Caption = lbl.Caption & Blah
Loop
Close #free
errhand:
Exit Sub
End Sub
Sub DataLoadEnabled(path As String, lbl As Object)
On Error GoTo errhand
free = FreeFile
Open path For Input As free
Do While Not EOF(free)
Line Input #free, Blah
lbl.Caption = lbl.Enabled & Blah
Loop
Close #free
errhand:
Exit Sub
End Sub
Sub DataLoadCombo(path As String, Combo As ComboBox)
'Call Load_ComboBox("c:\windows\desktop\combo.cmb", Combo1)

    Dim What As String
    On Error Resume Next
    Open path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Combo.AddItem What$
    Wend
    Close #1
End Sub
Sub ListRemove(Lst As ListBox)
'Put this in list_Dbl Click
bolt% = Lst.ListIndex
Lst.RemoveItem (bolt%)
End Sub
Sub DataLoadList(path As String, Lst As ListBox)
'Ex: Call Load_ListBox("c:\windows\desktop\list.lst", list1)

    Dim What As String
    On Error Resume Next

    Open path$ For Input As #1
    While Not EOF(1)
        Input #1, What$
        DoEvents
        Lst.AddItem What$
    Wend
    Close #1
End Sub
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, Wavy)

End Function

Function EncryptText(Text, types)
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

EncryptText = Process$
End Function
Function Encrypt(Text, Number)
'Same thing as the first only a little better
'You hafta do the same coding though except
'if u want to decyrpt have last
'parameter be true encyrpt false
For God = 1 To Len(Text)
If Number = 0 Then
Current$ = Asc(Mid(Text, God, 25)) - 25
Else
Current$ = Asc(Mid(Text, God, 25)) + 25
End If
Process$ = Process$ & Chr(Current$)
Next God

Encrypt = Process$
End Function





Public Sub DisableCTrLALTDEL()
Dim ret  As Long
 Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Sub AIMMassIM(lis As ListBox, Txt As String)
    If lis.ListCount = 0 Then
        Exit Sub
    Else
    End If

    Dim Moo
    For Moo = 0 To lis.ListCount - 1
        Call AIMSendIM(lis.List(Moo), Txt, True)
    Next Moo
End Sub
Public Sub EnableCRTLALTDEL()
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub





Function AIMChatLine(iLine As Integer) As String
'Get any line out of the chat.
'This will only get about the last 10 - 15 lines.
    String1$ = ChatText2$
    ChatLine$ = WinGetLine(String1$, iLine%)

End Function
Public Sub AIMReplyIM(stMessage As String)
'Responds to One Focused Im, then closes it.
    On Error Resume Next

    'Find An Open IM!
    If FindIM& <> 0 Then
    strName$ = ImFromWho$
    Call AIMCloseIM

    'Send the IM!
    Call AIMSendIM(strName$, stMessage$, False)
    strName$ = ""
    End If

End Sub
Function WinGetLine(stText As String, iLine As Integer) As String
'This will get one line of a textbox.
'You can use this for every line except the first!

    On Error Resume Next
    If InStr(1, stText, Chr(13)) = 0 Then
    If iLine = 1 Then
    Win_GetLine$ = stText$
    Exit Function
    End If 'iLine
    End If 'InStr
    
    iStop% = 0
    For intLoop3% = 1 To Len(stText$) Step 1
    stLine1$ = Mid(stText$, intLoop3%, 1)
    
    If stLine1$ = Chr(13) Then
    iStop% = iStop% + 1
        If iStop% >= iLine% Then
            Win_GetLine$ = stLine2$
            Exit Function
        End If 'iStop
        stLine1$ = ""
        stLine2$ = ""
        End If
    stLine2$ = stLine2$ + stLine1$
    Next intLoop3%
    

    Win_GetLine$ = stLine2$
End Function
Public Sub AIMCloseIM()
'Close focused Im if open.
    DoEvents:
    lIm& = FindIM&
    If lIm& > 0 Then 'Close
    Call WinKill(lIm&)
    End If 'lIm&
End Sub
Function AIMUserName() As String
'Gets name of user

    String1$ = GetCaption(FindMain&)
     String2$ = Left(String1$, Len(String1$) - 20)
    
    string01$ = GetCaption(FindMain&)
    lApost& = InStr(1, string01$, "'")
    string02$ = Left(string01$, lApost& - 1)

    If String2$ = string02$ Then
    UserName$ = String2$
    Else
    UserName$ = "unknown"
    End If
End Function



Function AIMChatName() As String
'Get name of Chat
    String1$ = GetCaption(FindChat&)
    ChatName$ = right(String1$, Len(String1$) - 11)
    
End Function
Sub MoveLR(Thing As Object, Left As Boolean, HowMuch As String)
'If you want it left have 2nd to last parimeter say true
'if right false
If Left = True Then
Thing.Left = Val(Thing.Left) - HowMuch
End If
If Left = False Then
Thing.Left = Val(Thing.Left) + HowMuch
End If
End Sub
Sub MoveUD(Thing As Object, Up As Boolean, HowMuch As String)
'If you want it up have 2nd to last parimeter say true
'if down false
If Up = True Then
Thing.Top = Val(Thing.Top) + HowMuch
End If
If Up = False Then
Thing.Top = Val(Thing.Top) - HowMuch
End If
End Sub
Function GetLine(stText As String, iLine As Integer) As String
'This will get one line of a textbox.
'You can use this for every line except the first!

    On Error Resume Next
    If InStr(1, stText, Chr(13)) = 0 Then
    If iLine = 1 Then
    GetLine$ = stText$
    Exit Function
    End If 'iLine
    End If 'InStr
    
    iStop% = 0
    For intLoop3% = 1 To Len(stText$) Step 1
    stLine1$ = Mid(stText$, intLoop3%, 1)
    
    If stLine1$ = Chr(13) Then
    iStop% = iStop% + 1
        If iStop% >= iLine% Then
            GetLine$ = stLine2$
            Exit Function
        End If 'iStop
        stLine1$ = ""
        stLine2$ = ""
        End If
    stLine2$ = stLine2$ + stLine1$
    Next intLoop3%
    

   GetLine$ = stLine2$
End Function
Public Sub AIMChatIgnore(iIndex As Integer)
'Ignore by index
    DoEvents:
    On Error Resume Next
    If FindChat& <> 0 Then
    Do Until lngAllPeop& <> 0
    lngAllPeop& = FindWindowEx(FindChat&, 0, "_Oscar_Tree", vbNullString)
    Loop 'BuddyTree
    End If 'FindChat
    
    lngSetCurs& = SendMessageByNum(lngAllPeop&, LB_SETCURSEL, iIndex% - 1, 0)
    lLbHan& = FindWindowEx(FindChat&, 0, "_Oscar_IconBtn", vbNullString)
    lLbHan2& = FindWindowEx(FindChat&, lLbHan&, "_Oscar_IconBtn", vbNullString)
    Call WinClick(lLbHan2&)
    

End Sub

Public Sub AIMChatX(strSName As String)
'Ignore By Name, this is much Nicer!
    DoEvents:
    On Error Resume Next
    If FindChat& <> 0 Then
    Do Until BuddyTree& <> 0
    BuddyTree& = FindWindowEx(FindChat&, 0, "_Oscar_Tree", vbNullString)
    Loop 'BuddyTree
    End If 'FindChat
    
    Call SendMessageByString(BuddyTree&, LB_SETCURSEL, intLoop%, 0)
    Count& = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)

    For intLoop% = 0 To Count& - 1
    lLength& = SendMessage(BuddyTree&, LB_GETTEXTLEN, intLoop%, 0)
    lEmptyT$ = String(lLength&, " ")
    
    lFullTe$ = SendMessageByString(BuddyTree&, LB_GETTEXT, intLoop%, lEmptyT$)
    lTab& = InStr(lEmptyT$, Chr(9)) 'Find Tab.
    sName$ = right$(lEmptyT$, Len(lEmptyT$) - lTab&)

    lTab& = InStr(sName$, Chr(9))
    sText$ = right$(sName$, Len(sName$) - lTab&)
    sName2$ = sText$
    
    If LCase(sName2$) = LCase(strSName$) Then _
    Call AIMChatIgnore(intLoop% + 1)
Next intLoop%
End Sub
Public Sub WinClick(ByVal lWinHandle As Long)
'This clicks any command button
    DoEvents
    Call SendMessage(lWinHandle, WM_LBUTTONDOWN, 0, vbNullString)
    Call SendMessage(lWinHandle, WM_LBUTTONUP, 0, vbNullString)
    DoEvents
End Sub
Sub AIMAddRoomList(lis As ListBox)
    Dim ChatRoom As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer, BuddyTree As Long

    ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatRoom& <> 0 Then
        Do
            BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
            TabPos = InStr(buffer$, Chr$(9))
            NameText$ = right$(buffer$, (Len(buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            For mooz = 0 To lis.ListCount - 1
                If name$ = lis.List(mooz) Then
                    Well% = 123
                    GoTo endz
                End If
            Next mooz
            If Well% <> 123 Then
                lis.AddItem name$
            Else
            End If
endz:
        Next MooLoo
    End If
End Sub




Sub ListKillDupes(Lst As Control)
On Error Resume Next
For i = 0 To Lst.ListCount - 1
For e = 0 To Lst.ListCount - 1
If LCase(Lst.List(i)) Like LCase(Lst.List(e)) And i <> e Then
Lst.RemoveItem (e)
End If
Next e
Next i
End Sub
Sub ListTransfer(List1 As ListBox, ls As ListBox, Transferwhat As String)
Dim MegaBolt As Long
MegaBolt = ls.ListIndex
ls.Selected(MegaBolt) = True
List1.AddItem (ls.Text)
ls.RemoveItem (MegaBolt)
End Sub
Sub ListTransferCaption(List1 As ListBox, Ob As Object, Transferwhat As String)
Dim MegaBolt As Long
MegaBolt = List1.ListIndex
List1.Selected(MegaBolt) = True
Ob.Caption = List1.Text
List1.Selected(MegaBolt) = True
List1.RemoveItem (MegaBolt)
End Sub


Sub AIMAddRoomCombo(cmb As ComboBox)
    Dim ChatRoom As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer, BuddyTree As Long

    ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)

    If ChatRoom& <> 0 Then
        Do
            BuddyTree& = FindWindowEx(ChatRoom&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
            TabPos = InStr(buffer$, Chr$(9))
            NameText$ = right$(buffer$, (Len(buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            For mooz = 0 To cmb.ListCount - 1
                If name$ = cmb.List(mooz) Then
                    Well% = 123
                    GoTo endz
                End If
            Next mooz
            If Well% <> 123 Then
                cmb.AddItem name$
            Else
            End If
endz:
        Next MooLoo
    End If
End Sub
Sub MsgError(Progname As String)
Call MsgBox("What The Heck are you doing!!", vbExclamation, Progname)

End Sub



Function AOLSNLastChatLine()
chattext$ = ChatLastLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        sn = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = sn
End Function

Sub AIMClearChat()
    Dim ChatWindow As Long, BorderThing As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    BorderThing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Call SendMessageByString(BorderThing&, WM_SETTEXT, 0, "")
End Sub
Sub AOLSendMail(Recipiants, subject, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
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
Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
pUser = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function
Sub AIMCS(SayWhat As String)
    Dim ChatWindow As Long, Thing As Long, Thing2 As Long
    Dim SetChatText As Long, Buttin As Long, Buttin2 As Long, Buttin3 As Long
    Dim SendButtin As Long, Click As Long
    
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Thing& = FindWindowEx(ChatWindow&, 0&, "WndAte32Class", vbNullString)
    Thing2& = FindWindowEx(ChatWindow&, Thing&, "WndAte32Class", vbNullString)
    SetChatText& = SendMessageByString(Thing2&, WM_SETTEXT, 0, SayWhat$)
    Buttin& = FindWindowEx(ChatWindow&, 0, "_Oscar_IconBtn", vbNullString)
    Buttin2& = FindWindowEx(ChatWindow&, Buttin&, "_Oscar_IconBtn", vbNullString)
    Buttin3& = FindWindowEx(ChatWindow&, Buttin2&, "_Oscar_IconBtn", vbNullString)
    SendButtin& = FindWindowEx(ChatWindow&, Buttin3&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
End Sub



Sub FormCenter(f As Form)
f.Top = (Screen.Height * 0.85) / 2 - f.Height / 2
f.Left = Screen.Width / 2 - f.Width / 2
End Sub
Function AOLChatLastLineWithSN()
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
Function AOLChatLastLine()
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
ChatLastLine = ChatTrim$
End Function
Public Sub DragForm(frm As Form)
ReleaseCapture
X = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
'To use this,  put the following code in the "Mousedown"  dec

End Sub


Sub BlackBrownForm(frm As Form)
Call FadeForm(frm, black, brown)
End Sub
Sub AIMGoToRoom(Room As String)
' This sub takes the user to a room without inviteing people
' Kinda like a enter room

    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
    Dim TabThing As Long, ChtIcon As Long
    Dim ChatIcon As Long, ChatInvite As Long, ToWhoBox As Long, SetWho As Long
    Dim MessageBox As Long, RealBox As Long, SetMessage As Long
    Dim MesRoom As Long, EdBox As Long, RoomBox As Long, SetRoom As Long
    Dim SendIcon1 As Long, SendIcon2 As Long, SendIcon As Long, who As String
    Dim Click As Long, MesBox As Long
    who$ = Get_UserSN
    If who$ = "[Could not retrieve]" Then Exit Sub
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabThing& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    ChtIcon& = FindWindowEx(TabThing&, 0, "_Oscar_IconBtn", vbNullString)
    ChatIcon& = FindWindowEx(TabThing&, ChtIcon&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(ChatIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(ChatIcon&, WM_LBUTTONUP, 0, 0&)
    Freeze 0.2
    
    ChatInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    ToWhoBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    SetWho& = SendMessageByString(ToWhoBox&, WM_SETTEXT, 0, who$)
    
    MessageBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    RealBox& = FindWindowEx(ChatInvite&, MessageBox&, "Edit", vbNullString)
    SetMessage& = SendMessageByString(RealBox&, WM_SETTEXT, 0, "")
    
    MesBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    EdBox& = FindWindowEx(ChatInvite&, MesBox&, "Edit", vbNullString)
    RoomBox& = FindWindowEx(ChatInvite&, EdBox&, "Edit", vbNullString)
    SetRoom& = SendMessageByString(RoomBox&, WM_SETTEXT, 0, Room$)
    
    SendIcon1& = FindWindowEx(ChatInvite&, 0, "_Oscar_IconBtn", vbNullString)
    SendIcon2& = FindWindowEx(ChatInvite&, SendIcon1&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(ChatInvite&, SendIcon2&, "_Oscar_IconBtn", vbNullString)
    
    Click& = SendMessage(SendIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendIcon&, WM_LBUTTONUP, 0, 0&)
End Sub



Sub Subtract(l As Label, Huh As String)
l.Caption = Val(l) - Huh
End Sub
Sub Add(l As Label, Huh As String)
l.Caption = Val(l) + Huh
End Sub
Sub WindowHide(THeWindow&)
Call ShowWindow(THeWindow&, SW_HIDE)
End Sub
Sub WindowShow(THeWindow&)
Call ShowWindow(THeWindow&, SW_SHOW)
End Sub
Function AIMUserSN() As String
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      AIMUserSN = "[ not online.]"
      Exit Function
    End If

Start:
    Dim GetIt As String, Clear As String
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    GetIt$ = GetTheCaption(BuddyList&)
    Clear$ = ReplaceAString(GetIt$, "'s Buddy List", "")
    AIMUserSN = Clear$
End Function
Function ReplaceAString(MyString As String, ToFind As String, ReplaceWith As String) As String
' By dos, from dos23.bas He gets all the credit for this one
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
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
    ReplaceAString$ = NewString$
End Function
Function GetCaption(TheWin)
' From Dos32.bas He gets fill credit
    Dim WindowLngth As Integer, WindowTtle As String, Moo As String
    
    WindowLngth% = GetWindowTextLength(TheWin)
    WindowTtle$ = String$(WindowLngth%, 0)
    Moo$ = GetWindowText(TheWin, WindowTtle$, (WindowLngth% + 1))
    GetCaption = WindowTtle$
End Function



Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, Wavy As Boolean)
'by monk-e-god
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen$ = Len(TheText)
    For i = 1 To textlen$
        TextDone$ = Left(TheText, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
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
    Next i
    FadeTwoColor = Faded$
End Function


Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = right(TheText, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
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
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
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
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function
Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.blue = Int(CVal / 65536)
  GetRGB.green = Int((CVal - (65536 * GetRGB.blue)) / 256)
  GetRGB.red = CVal - (65536 * GetRGB.blue + 256 * GetRGB.green)
End Function
Public Function RGBToHex2(rgbvalue As Long) As String
    Dim hexstate As String, hexlen As Long
    Let hexstate$ = Hex(rgbvalue&)
    Let hexlen& = Len(hexstate$)
    Select Case hexlen&
        Case 1&
            Let RGBToHex2$ = "00000" & hexstate$
            Exit Function
        Case 2&
            Let RGBToHex2$ = "0000" & hexstate$
            Exit Function
        Case 3&
            Let RGBToHex2$ = "000" & hexstate$
            Exit Function
        Case 4&
            Let RGBToHex2$ = "00" & hexstate$
            Exit Function
        Case 5&
            Let RGBToHex2$ = "0" & hexstate$
            Exit Function
        Case 6&
            Let RGBToHex2$ = "" & hexstate$
            Exit Function
        Case Else
            Exit Function
    End Select
End Function

Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
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
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Sub UnloadLoad(Currentform As Form, LoadedForm As Form)
LoadedForm.Show
Unload Currentform
End Sub
Sub HideLoad(HiddenForm As Form, LoadedForm As Form)
LoadedForm.Show
HiddenForm.Hide
End Sub




Sub AIMAddBuddyList(lis As ListBox)
    Dim BuddyList As Long, TabGroup As Long
    Dim BuddyTree As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0 Then
        Do
            TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
            BuddyTree& = FindWindowEx(TabGroup&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
            TabPos = InStr(buffer$, Chr$(9))
            NameText$ = right$(buffer$, (Len(buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            If InStr(name$, "(") <> 0 And InStr(name$, ")") <> 0 Then
                GoTo HellNo
            End If
            For mooz = 0 To lis.ListCount - 1
                If name$ = lis.List(mooz) Then
                    Well% = 123
                    GoTo HellNo
                End If
            Next mooz
            If Well% <> 123 Then
                lis.AddItem name$
            Else
            End If
HellNo:
        Next MooLoo
    End If
End Sub
Sub OpenCDRom()
retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub
Sub CloseCDRom()
retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)


End Sub
Sub AIMMassInvite(lis As ListBox, say As String, Room As String)
    Dim ChatWindow As Long, Moo As String

    If lis.ListCount = 0 Then
        Exit Sub
    Else
    End If
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call WinKill(ChatWindow&)

    Moo$ = ListToString(lis)
    Call AIMInviteBuddy(Moo$, say, Room)
End Sub
Sub WinKill(TheWind&)
    Call PostMessage(TheWind&, WM_CLOSE, 0&, 0&)
End Sub
Function ListToString(TheList As ListBox) As String
' by dos
    Dim DoList As Long, MailString As String
    If TheList.List(0) = "" Then Exit Function
    For DoList& = 0 To TheList.ListCount - 1
        MailString$ = MailString$ & TheList.List(DoList&) & ", "
    Next DoList&
    MailString$ = Mid(MailString$, 1, Len(MailString$) - 2)
    ListToString$ = MailString$
End Function
Public Sub AIMClickForIM(stName As String)
'This is like what Aim does.  If you click it
'Opens an instant message with just there name
'Ready for writing.
    DoEvents:
    Call AIMOpenIM
    lPareIm& = FindWindowEx(FindIM&, 0, "_Oscar_PersistantCombo", vbNullString)
    lHandIm& = FindWindowEx(lPareIm&, 0, "Edit", vbNullString)
    Call WinSetTxt(lHandIm&, stName$)

End Sub
Public Sub AIMOpenIM()
'Opens New Im.
    DoEvents:
    lParIm& = FindWindowEx(FindMain&, 0, "_Oscar_TabGroup", vbNullString)
    lHanIm& = FindWindowEx(lParIm&, 0, "_Oscar_IconBtn", vbNullString)
    Call WinClick(lHanIm&)
End Sub
Public Sub WinSetTxt(ByVal lWinHandle As Long, ByVal stText As String)
'This writes text to any object
    If Len(stText) > 0 Then
    Call SendMessageByString(lWinHandle, WM_SETTEXT, 0, stText)
    End If
End Sub


Sub AIMAddBuddyCombo(cmb As ComboBox)
    Dim BuddyList As Long, TabGroup As Long
    Dim BuddyTree As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, buffer As String
    Dim TabPos, NameText As String, Text As String
    Dim mooz, Well As Integer

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0 Then
        Do
            TabGroup& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
            BuddyTree& = FindWindowEx(TabGroup&, 0, "_Oscar_Tree", vbNullString)
        Loop Until BuddyTree& <> 0
        LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
        For MooLoo = 0 To LopGet - 1
            Call SendMessageByString(BuddyTree&, LB_SETCURSEL, MooLoo, 0)
            NameLen = SendMessage(BuddyTree&, LB_GETTEXTLEN, MooLoo, 0)
            buffer$ = String$(NameLen, 0)
            Moo2 = SendMessageByString(BuddyTree&, LB_GETTEXT, MooLoo, buffer$)
            TabPos = InStr(buffer$, Chr$(9))
            NameText$ = right$(buffer$, (Len(buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            Text$ = right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = Text$
            If InStr(name$, "(") <> 0 And InStr(name$, ")") <> 0 Then
                GoTo HellNo
            End If
            For mooz = 0 To cmb.ListCount - 1
                If name$ = cmb.List(mooz) Then
                    Well% = 123
                    GoTo HellNo
                End If
            Next mooz
            If Well% <> 123 Then
                cmb.AddItem name$
            Else
            End If
HellNo:
        Next MooLoo
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
Sub WaitOk()
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
Sub WaitOK2()
Do
DoEvents
okw = FindWindow("_AIM_ChatWnd", vbNullString)
If prog_state$ = "off" Then
Exit Sub
Exit Do
End If
DoEvents
Loop Until okw <> 0
okb = FindChildByTitle(okw, "OK")
 okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)

End Sub
Function GetTheCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetTheCaption = hwndTitle$
End Function

Sub EvilProg(frm As Form)
'This makes it so it does lots of weird stuff to your computer
'but doesnt screw it up
'You hafta pressthe reboot button on your comp to
'get out of the program and even when that happens they
'hafta wait a while for that slow mcafee/norton not shut down
'error.Put this is a button_Click or something like that
FormOnTop frm
DisableCTrLALTDEL
Do
MsgBox "Uyhs has taken over your computer", vbCritical, "!!!"
MsgBox "Uyhs has taken over your computer", vbExclamation, "!!!"
MsgBox "Uyhs has taken over your computer", vbInformation, "!!!"
MsgBox "Uyhs has taken over your computer", vbYesNo, "!!!"
MsgBox "Uyhs has taken over your computer", vbAbortRetryIgnore, "!!!"
MsgBox "Uyhs has taken over your computer", vbCritical, "!!!"
OpenCDRom
Freeze 10
CloseCDRom
Freeze 1.5
Loop
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


Sub FadeForm(FormX As Form, Colr1, Colr2)
    B1 = GetRGB(Colr1).blue
    G1 = GetRGB(Colr1).green
    R1 = GetRGB(Colr1).red
    B2 = GetRGB(Colr2).blue
    G2 = GetRGB(Colr2).green
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
Function FadeByColor2(Colr1, Colr2, TheText$, Wavy As Boolean)
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + right(dacolor2$, 2))
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
Sub CountDownTenSec(StartingCount As String, lbl As Label)
lbl = StartingCount
Do
lbl.Caption = Val(lbl) - 0.01
Freeze 0.01
Loop
End Sub
Sub ShowDate(theDate As Label)
theDate = Date
End Sub

Public Sub Fadepicture(frm As PictureBox, TopColor&, BottomColor&)


Dim SaveScale%, SaveStyle%, SaveRedraw%, ThisColor&
Dim i&, j&, X&, Y&, pixels%
Dim RedDelta As Single, GreenDelta As Single, BlueDelta As Single
Dim aRed As Single, aGreen As Single, aBlue As Single
Dim TopColorRed%, TopColorGreen%, TopColorBlue%
Dim BottomColorRed%, BottomColorGreen%, BottomColorBlue%

SaveScale = frm.ScaleMode
SaveStyle = frm.DrawStyle
SaveRedraw = frm.AutoRedraw

frm.ScaleMode = 3
TopColorRed = TopColor And 255
TopColorGreen = (TopColor And 65280) / 256
TopColorBlue = (TopColor And 16711680) / 65536
BottomColorRed = BottomColor And 255
BottomColorGreen = (BottomColor And 65280) / 256
BottomColorBlue = (BottomColor And 16711680) / 65536
    aRed = TopColorRed
    aGreen = TopColorGreen
    aBlue = TopColorBlue
    pixels = frm.ScaleWidth
    
    If pixels <= 0 Then Exit Sub
    
    ColorDifRed = (BottomColorRed - TopColorRed)
    ColorDifGreen = (BottomColorGreen - TopColorGreen)
    ColorDifBlue = (BottomColorBlue - TopColorBlue)
    
    RedDelta = ColorDifRed / pixels
    GreenDelta = ColorDifGreen / pixels
    BlueDelta = ColorDifBlue / pixels
    
    frm.DrawStyle = 5
    frm.AutoRedraw = True
    
    
 For Y = 0 To pixels
        aRed = aRed + RedDelta
        If aRed < 0 Then aRed = 0
        aGreen = aGreen + GreenDelta
        If aGreen < 0 Then aGreen = 0
        aBlue = aBlue + BlueDelta
        If aBlue < 0 Then aBlue = 0
        ThisColor = RGB(aRed, aGreen, aBlue)
        If ThisColor > -1 Then
        
        frm.Line (Y - 2, -2)-(Y - 2, frm.Height + 2), ThisColor, BF
        End If
    Next Y
  

frm.ScaleMode = SaveScale
frm.DrawStyle = SaveStyle
frm.AutoRedraw = SaveRedraw
End Sub


Sub Freeze(interval) 'AKA Pause or Timeout
Dim time
time = Timer
Do While Timer - time < Val(interval)
DoEvents
Loop
End Sub
Sub CoolFormBegining(Form1 As Form, LeftPostion As String, HeightPostion As String, WidthPostion As String, THETIMER As Timer)
'put this in a timer and in the form put thetimer.enabled=true
' make the timers interval 1

Form1.Enabled = False
Form1.Left = 510
Form1.Left = 550
Form1.Height = 720
Form1.Left = 600
Call SunriseForm(Form1)
Freeze 0.1
Form1.Left = 650
Form1.Left = 700
Form1.Height = 750
Form1.Left = 750
Call IceForm(Form1)
Freeze 0.1
Form1.Left = 800
Form1.Left = 850
Form1.Height = 800
Form1.Left = 900
Call NeonForm(Form1)
Freeze 0.1
Form1.Left = 950
Form1.Left = 1000
Form1.Height = 850
Form1.Left = 1050
Call BlackWhiteForm(Form1)
Freeze 0.1
Form1.Left = 1100
Form1.Left = 1150
Form1.Height = 900
Form1.Left = 1200
Call FadeForm(Form1, Gold, black)
Freeze 0.1
Form1.Left = 1250
Form1.Left = 1300
Form1.Height = 950
Form1.Left = 1350
Call FadeForm(Form1, SEAGREEN, black)
Freeze 0.1
Form1.Left = 510
Form1.Left = 550
Form1.Height = 720
Form1.Left = 600
Call SunriseForm(Form1)
Freeze 0.1
Form1.Left = 650
Form1.Left = 700
Form1.Height = 750
Form1.Left = 750
Call IceForm(Form1)
Freeze 0.1
Form1.Left = 800
Form1.Left = 850
Form1.Height = 800
Form1.Left = 900
Call NeonForm(Form1)
Freeze 0.1
Form1.Left = 950
Form1.Left = 1000
Form1.Height = 850
Form1.Left = 1050
Call BlackWhiteForm(Form1)
Freeze 0.1
Form1.Left = 1100
Form1.Left = 1150
Form1.Height = 900
Form1.Left = 1200
Call FadeForm(Form1, Gold, black)
Freeze 0.1
Form1.Left = 1250
Form1.Left = 1300
Form1.Height = 950
Form1.Left = 1350
Call FadeForm(Form1, SEAGREEN, black)
Freeze 0.1
Form1.Left = 510
Form1.Left = 550
Form1.Height = 720
Form1.Left = 600
Call SunriseForm(Form1)
Freeze 0.1
Form1.Left = 650
Form1.Left = 700
Form1.Height = 750
Form1.Left = 750
Call IceForm(Form1)
Freeze 0.1
Form1.Left = 800
Form1.Left = 850
Form1.Height = 800
Form1.Left = 900
Call NeonForm(Form1)
Freeze 0.1
Form1.Left = 950
Form1.Left = 1000
Form1.Height = 850
Form1.Left = 1050
Call BlackWhiteForm(Form1)
Freeze 0.1
Form1.Left = 1100
Form1.Left = 1150
Form1.Height = 900
Form1.Left = 1200
Call FadeForm(Form1, Gold, black)
Freeze 0.1
Form1.Left = 1250
Form1.Left = 1300
Form1.Height = 950
Form1.Left = 1350
Call FadeForm(Form1, SEAGREEN, black)
Freeze 0.1
Form1.Left = 510
Form1.Left = 550
Form1.Height = 720
Form1.Left = 600
Call SunriseForm(Form1)
Freeze 0.1
Form1.Left = 650
Form1.Left = 700
Form1.Height = 750
Form1.Left = 750
Call IceForm(Form1)
Freeze 0.1
Form1.Left = 800
Form1.Left = 850
Form1.Height = 800
Form1.Left = 900
Call NeonForm(Form1)
Freeze 0.1
Form1.Left = 950
Form1.Left = 1000
Form1.Height = 850
Form1.Left = 1050
Call BlackWhiteForm(Form1)
Freeze 0.1
Form1.Left = 1100
Form1.Left = 1150
Form1.Height = 900
Form1.Left = 1200
Call FadeForm(Form1, Gold, black)
Freeze 0.1
Form1.Left = 1250
Form1.Left = 1300
Form1.Height = 950
Form1.Left = 1350
Call FadeForm(Form1, SEAGREEN, black)
Freeze 0.1
Form1.Left = 510
Form1.Left = 550
Form1.Height = 720
Form1.Left = 600
Call SunriseForm(Form1)
Freeze 0.1
Form1.Left = 650
Form1.Left = 700
Form1.Height = 750
Form1.Left = 750
Call IceForm(Form1)
Freeze 0.1
Form1.Left = 800
Form1.Left = 850
Form1.Height = 800
Form1.Left = 900
Call NeonForm(Form1)
Freeze 0.1
Form1.Left = 950
Form1.Left = 1000
Form1.Height = 850
Form1.Left = 1050
Call BlackWhiteForm(Form1)
Freeze 0.1
Form1.Left = 1100
Form1.Left = 1150
Form1.Height = 900
Form1.Left = 1200
Call FadeForm(Form1, Gold, black)
Freeze 0.1
Form1.Left = 1250
Form1.Left = 1300
Form1.Height = 950
Form1.Left = 1350
Call FadeForm(Form1, SEAGREEN, black)
Freeze 0.1
Form1.Left = LastPostion
Form1.Height = HeightPostion
Form1.Width = WidthPostion
Form1.Enabled = True
THETIMER.Enabled = False
End Sub


Public Sub DataSaveCombo(ByVal Directory As String, Combo As ComboBox)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Combo.ListCount - 1
        Print #1, Combo.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Sub
Sub IceForm(theform As Form)
Call FadeForm(theform, LBLUE, NAVY)
End Sub
Sub SunriseForm(theform As Form)
Call FadeForm(theform, yellow, red)
End Sub
Sub NeonForm(theform As Form)
Call FadeForm(theform, LGREEN, yellow)
End Sub


Function AIMCheckOnline(Dir1 As Boolean, Dir2 As Boolean, UnloadForm As Boolean, LoadForm As Boolean, LoadWhatForm As Form, UnloadWhatForm As Form)
'Checks if your online
'It should be something like a start up screen where they type there name in
'a Textbox
Dim AIMUSER As Long
AIMUSER& = FindWindow("_Oscar_BuddyListWin", vbNullString)
If AIMUSER& <> 0& Then
GoTo Poo
Else
Select Case MsgBox("You are not online...Do u wish to sign on?", vbYesNo, "[Not Online]")
Case vbYes
If Dir1 = True Then Call AIMLoad
If Dir2 = True Then Call AIMLoad2
Case vbNo
Exit Function
End Select
End If
Poo:
MsgBox "Login Success!", vbExclamation, "[Online]"
Beep
Freeze 0.5
Beep
If UnloadForm = True Then Unload UnloadWhatForm
If LoadForm = True Then LoadWhatForm.Show
End Function
Sub WavPlay(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub
Public Sub MidiPlay(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Public Sub MidiStop(MIDIFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(MIDIFile$)
    If SafeFile$ <> "" Then
        Call mciSendString("stop " & MIDIFile$, 0&, 0, 0)
    End If
End Sub
Sub AIMLoad()
    Dim X As Long, NoFreeze As Integer
    
    X& = Shell("C:\Program Files\AIM95\aim.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Sub AIMLoad2()
Dim X As Long, nofree As Integer
X& = Shell("C:\AOL Instant Messenger\AIM.exe", 1): NoFreeze% = DoEvents(): Exit Sub
End Sub


Sub RunMenuByString(Application, StringSearch)
' From Hix he gets full credit

    Dim ToSearch As Integer, MenuCount As Integer, FindString
    Dim ToSearchSub As Integer, MenuItemCount As Integer, GetString
    Dim SubCount As Integer, MenuString As String, GetStringMenu As Integer
    Dim MenuItem As Integer, RunTheMenu As Integer
    
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
Function ListCount(Lst As ListBox)
    Dim Moo As Integer

    Moo% = Lst.ListCount
    ListCount = Moo%
End Function

Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub AIMChatLink(Address As String, Text As String)
    AIMCS "<A HREF=""" + Address$ + """>" + Text$ + ""
End Sub

Public Sub FadeObject(theobject As Object, color1 As Long, color2 As Long)
    Dim index As Long, index2 As Long, lngclrhold(1& To 3&) As Single
    Dim lnghold As Single, obj100th As Double, lngswidth As Long
    Dim strcolor1 As String, strcolor2 As String, strred1 As String, strgreen1 As String
    Dim strblue1 As String, strred2 As String, strgreen2 As String, strblue2 As String
    Dim lngred1 As Long, lnggreen1 As Long, lngblue1 As Long, lngred2 As Long, lnggreen2 As Long
    Dim lngblue2 As Long, strcolor3 As String, strred3 As String, strgreen3 As String
    Dim strblue3 As String, lngred3 As Long, lnggreen3 As Long, lngblue3 As Long
    On Error Resume Next
    ReDim lngcolors(1& To 2&, 3&) As Integer
    Let strcolor1$ = RGBToHex2(color1&)
    Let strcolor2$ = RGBToHex2(color2&)
    Let strred1$ = "&h" & right$(strcolor1$, 2&)
    Let strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
    Let strblue1$ = "&h" & Left$(strcolor1$, 2&)
    Let strred2$ = "&h" & right$(strcolor2$, 2&)
    Let strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
    Let strblue2$ = "&h" & Left$(strcolor2$, 2&)
    Let lngred1& = Val(strred1$)
    Let lnggreen1& = Val(strgreen1$)
    Let lngblue1& = Val(strblue1$)
    Let lngred2& = Val(strred2$)
    Let lnggreen2& = Val(strgreen2$)
    Let lngblue2& = Val(strblue2$)
    Let lngcolors(1&, 1&) = lngred1&
    Let lngcolors(1&, 2&) = lnggreen1&
    Let lngcolors(1&, 3&) = lngblue1&
    Let lngcolors(2&, 1&) = lngred2&
    Let lngcolors(2&, 2&) = lnggreen2&
    Let lngcolors(2&, 3&) = lngblue2&
    ReDim lngcolors2(1& To 2&, 1& To 3&) As Double
    Let obj100th = theobject.ScaleWidth / 100&
    Let lngswidth& = theobject.ScaleHeight
    theobject.Cls
    For index& = 1& To 2&
        For index2& = 1& To 3&
            Let lngcolors2(index&, index2&) = lngcolors(index&, index2&)
        Next index2&
    Next index&
    Let theobject.BackColor = RGB(lngcolors2(2&, 1&), lngcolors2(2&, 2&), lngcolors2(2&, 3&))
    For index& = 1& To (2& - 1&)
        Let lngclrhold(1&) = (lngcolors2(index& + 1&, 1&) - lngcolors2(index&, 1&)) / (100& / (2& - 1&))
        Let lngclrhold(2&) = (lngcolors2(index& + 1&, 2&) - lngcolors2(index&, 2&)) / (100& / (2& - 1&))
        Let lngclrhold(3&) = (lngcolors2(index& + 1&, 3&) - lngcolors2(index&, 3&)) / (100& / (2& - 1&))
        For index2& = 1& To (100& / (2& - 1&))
            theobject.Line (lnghold, 0&)-(lnghold + obj100th, lngswidth&), RGB(lngcolors2(index&, 1&), lngcolors2(index&, 2&), lngcolors2(index&, 3&)), BF
            Let lngcolors2(index&, 1&) = lngcolors2(index&, 1&) + lngclrhold(1&)
            Let lngcolors2(index&, 2&) = lngcolors2(index&, 2&) + lngclrhold(2&)
            Let lngcolors2(index&, 3&) = lngcolors2(index&, 3&) + lngclrhold(3&)
            Let lnghold = lnghold + obj100th
        Next index2&
    Next index&
End Sub

Sub AIMInviteBuddy(who As String, message As String, Room As String)
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabThing As Long, ChtIcon As Long
    Dim ChatIcon As Long, ChatInvite As Long, ToWhoBox As Long, SetWho As Long
    Dim MessageBox As Long, RealBox As Long, SetMessage As Long
    Dim MesRoom As Long, EdBox As Long, RoomBox As Long, SetRoom As Long
    Dim SendIcon1 As Long, SendIcon2 As Long, SendIcon As Long, Click As Long
    Dim MesBox As Long
    
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabThing& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    ChtIcon& = FindWindowEx(TabThing&, 0, "_Oscar_IconBtn", vbNullString)
    ChatIcon& = FindWindowEx(TabThing&, ChtIcon&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(ChatIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(ChatIcon&, WM_LBUTTONUP, 0, 0&)

    Freeze 0.2
    
    ChatInvite& = FindWindow("AIM_ChatInviteSendWnd", "Buddy Chat Invitation ")
    ToWhoBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    SetWho& = SendMessageByString(ToWhoBox&, WM_SETTEXT, 0, who$)
    
    MessageBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    RealBox& = FindWindowEx(ChatInvite&, MessageBox&, "Edit", vbNullString)
    SetMessage& = SendMessageByString(RealBox&, WM_SETTEXT, 0, message$)
    MesBox& = FindWindowEx(ChatInvite&, 0, "Edit", vbNullString)
    EdBox& = FindWindowEx(ChatInvite&, MesBox&, "Edit", vbNullString)
    RoomBox& = FindWindowEx(ChatInvite&, EdBox&, "Edit", vbNullString)
    SetRoom& = SendMessageByString(RoomBox&, WM_SETTEXT, 0, Room$)
    SendIcon1& = FindWindowEx(ChatInvite&, 0, "_Oscar_IconBtn", vbNullString)
    SendIcon2& = FindWindowEx(ChatInvite&, SendIcon1&, "_Oscar_IconBtn", vbNullString)
    SendIcon& = FindWindowEx(ChatInvite&, SendIcon2&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendIcon&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendIcon&, WM_LBUTTONUP, 0, 0&)

End Sub
Sub KillWin(TheWind&)
Call PostMessage(TheWind&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub CloseErrors(lLimit As Long)
'lLimit is 1000 to allow long check.
    
    If FindWindow("#32770", vbNullString) > 0 Then
    Do Until lAimError& = 0 Or lLim& >= lLimit
    lLim = lLim& + 1
    lAimError& = FindWindow("#32770", vbNullString)
    Call WinKill(lAimError&)
    Loop 'lAimError&
    End If
End Sub
Sub AIMSendIM(SendName As String, SayWhat As String, CloseIM As Boolean)
' My send IM comes with a little thing where you can eather close
' it or not close it....
' Ex: Call IM_Send("ThereSn","Sup man",True) <-- that closes the IM
' Put False to not close the IM, All the IM sends have the TRUE FALSE thing
    Dim BuddyList As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If

Start:
 
    Dim TabWin As Long, IMbuttin As Long, IMWin As Long
    Dim ComboBox As Long, TextEditBox As Long, TextSet As Long
    Dim EditThing As Long, TextSet2 As Long, SendButtin As Long, Click As Long

    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    TabWin& = FindWindowEx(BuddyList&, 0, "_Oscar_TabGroup", vbNullString)
    IMbuttin& = FindWindowEx(TabWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(IMbuttin&, WM_LBUTTONUP, 0, 0&)
    Freeze 0.1
  
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    ComboBox& = FindWindowEx(IMWin&, 0, "_Oscar_PersistantCombo", vbNullString)
    TextEditBox& = FindWindowEx(ComboBox&, 0, "Edit", vbNullString)
    TextSet& = SendMessageByString(TextEditBox&, WM_SETTEXT, 0, SendName$)
    Freeze 0.1
    EditThing& = FindWindowEx(IMWin&, 0, "WndAte32Class", vbNullString)
    EditThing& = GetWindow(EditThing&, 2)
    TextSet2& = SendMessageByString(EditThing&, WM_SETTEXT, 0, SayWhat$)
    SendButtin& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(SendButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(SendButtin&, WM_LBUTTONUP, 0, 0&)
    If CloseIM = True Then
        WinKill (IMWin&)
    Else
        Exit Sub
    End If
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
Function ScrambleText2(TheText)
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
ScrambleText2 = scrambled$

Exit Function
End Function



Sub BlackWhiteForm(theform As Form)
Call FadeForm(theform, black, white)
End Sub

