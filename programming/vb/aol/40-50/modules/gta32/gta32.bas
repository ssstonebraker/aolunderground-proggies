Attribute VB_Name = "GTA32"
'GTA32.bas
'70/\/\7h3l30/\/\l3
'Giant
'Testicle
'Association
Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
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

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        Y As Long
End Type
Public Sub Iconc(aIcon As Long)
    Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub MailBox()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Writemail()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Mailcentermnu()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Mailcenterkw()
Dim aolframe&
Dim aoltoolbar&
Dim aolcombobox&
Dim edit&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolcombobox& = FindWindowEx(aoltoolbar&, 0&, "_aol_combobox", vbNullString)
edit& = FindWindowEx(aolcombobox&, 0&, "edit", vbNullString)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, "Mail Center")
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Printmnu()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Myfiles()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub MyAOL()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Favoritesmnu()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Keyword(Whichkeyword As String)

    Dim aol As Long, tool As Long
    Dim Toolbar As Long, Combo As Long
    Dim EditWin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, Whichkeyword$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub

Sub Favoriteskw()
Keyword ("Favorites")
End Sub
Sub Internetmnu()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Channelsmnu()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Internetkw()
Keyword ("Internet")
End Sub
Sub Peoplemnu()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub SignonScreenCaption(Newcaption As String)
Dim aolframe&
Dim mdiclient&
Dim aolchild&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
 Call SendMessageByString(aolchild, WM_SETTEXT, 0&, Newcaption$)
End Sub
Sub changeAOLcaption(Caption As String)
Dim aolframe&
aolframe& = FindWindow("aol frame25", vbNullString)
 Call SendMessageByString(aolframe, WM_SETTEXT, 0&, Caption$)
End Sub
Sub ChangeSignonStatus(statustext As String)
Dim aolframe&
Dim aolmodal&
Dim aolstatic&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolstatic& = FindWindowEx(aolmodal&, 0&, "_aol_static", vbNullString)
Call SendMessageByString(aolstatic, WM_SETTEXT, 0&, statustext$)
End Sub
Sub ChangeErrorMSG(Newerror As String)
'This is cool
'You know how you get signed off and it says "Your connection to aol has been lost
'this will change it
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim richcntl&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
richcntl& = FindWindowEx(aolchild&, 0&, "richcntl", vbNullString)
richcntl& = FindWindowEx(aolchild&, richcntl&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, Newerror$)

End Sub
Sub MailReadNew(Index As Long)
'index should be 0 if you want to read the first mail
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aoltabcontrol&
Dim aoltabpage&
Dim aoltree&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aoltabcontrol& = FindWindowEx(aolchild&, 0&, "_aol_tabcontrol", vbNullString)
aoltabpage& = FindWindowEx(aoltabcontrol&, 0&, "_aol_tabpage", vbNullString)
aoltree& = FindWindowEx(aoltabpage&, 0&, "_aol_tree", vbNullString)
Call SendMessage(aoltree&, LB_SETCURSEL, Index&, 0&)
Call PostMessage(aoltree&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(aoltree&, WM_KEYUP, VK_RETURN, 0&)
End Sub
Sub IDLE()
'put the following code in a timer with an interval of one.
'IDLE
Dim aolframe&
Dim aolmodal&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolicon& = FindWindowEx(aolmodal&, 0&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Public Sub RunMenuByString(SearchString As String)
'Thanx to dos for this
'www.dosfx.com
    Dim aol As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub

Sub Signoff()
RunMenuByString ("Sign Off")
   
End Sub
Sub Newfile()
RunMenuByString ("New")
End Sub
Sub SwitchScreennames()
RunMenuByString ("Switch Screen name")
End Sub
Sub PrivateChatroom(room As String)
Keyword ("aol://2719:2-2-" & room)
End Sub
Sub ChatSendAOL(message As String)
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim richcntl&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
richcntl& = FindWindowEx(aolchild&, 0&, "richcntl", vbNullString)
richcntl& = FindWindowEx(aolchild&, richcntl&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, message$)
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub IM_AOL(who, What As String)
Keyword "im"
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aoledit&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, who)
Dim richcntl&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
richcntl& = FindWindowEx(aolchild&, 0&, "richcntl", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, What$)
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub AIMaddtotext(Text As String)
'this will change the picture of the add to text!"
Dim oscarbuddylistwin&
Dim wndateclass&
Dim ateclass&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, Text$)
    
End Sub
Sub changeBuddylistcaptionAOl(Newcaption As String)
Dim aolframe&
Dim mdiclient&
Dim aolchild&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
Call SendMessageByString(aolchild&, WM_SETTEXT, 0&, Newcaption)

End Sub


Sub XiRCONcaptionchange(Newcaption As String)
Dim owlwindow&
owlwindow& = FindWindow("owl_window", vbNullString)
Call SendMessageByString(owlwindow&, WM_SETTEXT, 0&, Newcaption)
End Sub
Sub ChatUK()
Keyword ("chatuk")

End Sub
Sub sendEmail(who$, subject$, message$)
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
Dim mdiclient&
Dim aolchild&
Dim aoledit&
Dim richcntl&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
timeout (1)
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, who)
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
aoledit& = FindWindowEx(aolchild&, aoledit&, "_aol_edit", vbNullString)
aoledit& = FindWindowEx(aolchild&, aoledit&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, subject)
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
richcntl& = FindWindowEx(aolchild&, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, message)
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Public Sub timeout(Duration As Long)
Dim time_now As Long
time_now = Timer
    Do Until Timer - time_now >= Duration
        DoEvents
    Loop
End Sub
Sub AOLhide()
Dim aolframe&
aolframe& = FindWindow("aol frame25", vbNullString)
Call ShowWindow(aolframe, SW_HIDE)
End Sub
Sub AOLshow()
Dim aolframe&
aolframe& = FindWindow("aol frame25", vbNullString)
Call ShowWindow(aolframe, SW_SHOW)
End Sub
Sub AOLhidetoolbar()
Dim aolframe&
Dim aoltoolbar&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
Call ShowWindow(aoltoolbar, SW_HIDE)
End Sub
Sub AOLshowtoolbar()
Dim aolframe&
Dim aoltoolbar&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
Call ShowWindow(aoltoolbar, SW_SHOW)
End Sub
Sub ShowAOLsymbol()
Dim aolframe&
Dim aoltoolbar&
Dim aolglyph&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolglyph& = FindWindowEx(aoltoolbar&, 0&, "_aol_glyph", vbNullString)
Call ShowWindow(aolglyph, SW_SHOW)
End Sub
Sub HideAOLsymbol()
Dim aolframe&
Dim aoltoolbar&
Dim aolglyph&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolglyph& = FindWindowEx(aoltoolbar&, 0&, "_aol_glyph", vbNullString)
Call ShowWindow(aolglyph, SW_HIDE)
End Sub
Sub ShowPeopleWindow()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Call ShowWindow(aolicon, SW_SHOW)
End Sub
Sub HidePeopleWindow()
Dim aolframe&
Dim aoltoolbar&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aoltoolbar& = FindWindowEx(aolframe&, 0&, "aol toolbar", vbNullString)
aoltoolbar& = FindWindowEx(aoltoolbar&, 0&, "_aol_toolbar", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aoltoolbar&, aolicon&, "_aol_icon", vbNullString)
Call ShowWindow(aolicon, SW_HIDE)
End Sub
Sub ChatBold()
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub Chatitalic()
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub ChatUnderline()
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolchild&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Sub ChangeAMountofpeople(New_number As Long)
'changes where it says "25 user in room"
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aolstatic&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolstatic& = FindWindowEx(aolchild&, 0&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_aol_static", vbNullString)
aolstatic& = FindWindowEx(aolchild&, aolstatic&, "_aol_static", vbNullString)
Call SendMessageByString(aolstatic&, WM_SETTEXT, 0&, New_number)

End Sub
Sub BuddyInvite(who As String, message As String, room As String)
Keyword "buddychat"
timeout 2
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aoledit&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, who)
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
aoledit& = FindWindowEx(aolchild&, aoledit&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, message)
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aoledit& = FindWindowEx(aolchild&, 0&, "_aol_edit", vbNullString)
aoledit& = FindWindowEx(aolchild&, aoledit&, "_aol_edit", vbNullString)
aoledit& = FindWindowEx(aolchild&, aoledit&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, room)
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aolicon& = FindWindowEx(aolchild&, 0&, "_aol_icon", vbNullString)
Iconc aolicon
timeout 1
End Sub
Sub BuddyBomb(who As String, message As String, room As String, howmanytimes As Long)
'this could get you signed off
Do:
BuddyInvite who, message, room
timeout 2
howmanytimes = howmanytimes - 1
Loop Until howmanytimes = 0
End Sub
Sub HideCancel()
'hides the cancel button during signon
Dim aolframe&
Dim aolmodal&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolicon& = FindWindowEx(aolmodal&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolmodal&, aolicon&, "_aol_icon", vbNullString)
Call ShowWindow(aolicon, SW_HIDE)
End Sub
Sub SHowCancel()
'Shows the cancel button during signon
Dim aolframe&
Dim aolmodal&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolicon& = FindWindowEx(aolmodal&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolmodal&, aolicon&, "_aol_icon", vbNullString)
Call ShowWindow(aolicon, SW_SHOW)
End Sub
Sub ClickCancel()
'Click the cancel button during signon
Dim aolframe&
Dim aolmodal&
Dim aolicon&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolicon& = FindWindowEx(aolmodal&, 0&, "_aol_icon", vbNullString)
aolicon& = FindWindowEx(aolmodal&, aolicon&, "_aol_icon", vbNullString)
Iconc aolicon
End Sub
Function GetsignonStatus()
'gets the signon Status
Dim aolframe&
Dim aolmodal&
Dim aolstatic&
aolframe& = FindWindow("aol frame25", vbNullString)
aolmodal& = FindWindow("_aol_modal", vbNullString)
aolstatic& = FindWindowEx(aolmodal&, 0&, "_aol_static", vbNullString)
    Dim buffer As String
  Dim TextLength As Long
    TextLength& = SendMessage(aolstatic&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(aolstatic&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetsignonStatus = buffer
End Function

Function SearchMemberDirBroad(forwhat As String)
'this will search the member directory using a broad
'search and then return the amount of people in the
'search results
'you will probobly need to click more though..
'this is really fast
Keyword ("aol://4950:0000010000|all:" & forwhat)
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aollistbox&
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aollistbox& = FindWindowEx(aolchild&, 0&, "_aol_listbox", vbNullString)
SearchMemberDirBroad = SendMessage(aollistbox, LB_GETCOUNT, 0, 0)
End Function
Function searchmemberdirSpecific(broad As String, name As String, location As String)
'this just searches broaderfor ppl
Keyword ("aol://4950:0000010000|all:" & broad & "|member_name:" & name & "|location:" & location)
Dim aolframe&
Dim mdiclient&
Dim aolchild&
Dim aollistbox&
Pause 4
aolframe& = FindWindow("aol frame25", vbNullString)
mdiclient& = FindWindowEx(aolframe&, 0&, "mdiclient", vbNullString)
aolchild& = FindWindowEx(mdiclient&, 0&, "aol child", vbNullString)
aollistbox& = FindWindowEx(aolchild&, 0&, "_aol_listbox", vbNullString)
searchmemberdirSpecific = SendMessage(aollistbox, LB_GETCOUNT, 0, 0)
End Function
Function ListtoMailrecipient(list As ListBox)
Dim ListRecipients As Long
Dim Recipients As String
If list.list(0) = "" Then Exit Function
For ListRecipients& = 0 To list.ListCount - 1
Recipients$ = Recipients$ & list.list(ListRecipients) & ", "
Next ListRecipients&
Recipients$ = Mid(Recipients, 1, Len(Recipients) - 2)
ListtoMailrecipient = Recipients
End Function
 
