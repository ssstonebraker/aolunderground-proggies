Attribute VB_Name = "FLYMAN02"
'AIM Module Update 2002 Release 2

Option Explicit
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hWndCallback As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Const EM_UNDO = &HC7
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
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
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
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
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
Public Const MF_SEPARATOR = &H800&
Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const ENTA = 13
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const EM_LINESCROLL = &HB6
Private Const SPI_SCREENSAVERRUNNING = 97
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Type POINTAPI
   X As Long
   Y As Long
End Type
Function SendChat(text As String)
'Ex:Call SendChat("Hey")
'or
'Ex:Call SendChat(Text1.text)
Dim Parent As Long, Child1 As Long, Child2 As Long, child3 As Long, child4 As Long, child5 As Long, child6 As Long, Textset As Long
    Parent& = FindWindow("AIM_ChatWnd", vbNullString)
    Child1& = FindWindowEx(Parent&, 0&, "WndAte32Class", vbNullString)
    Child2& = FindWindowEx(Parent&, Child1&, "WndAte32Class", vbNullString)
    Textset& = SendMessageByString(Child2&, WM_SETTEXT, 0, text$)
    child3& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
    child4& = FindWindowEx(Parent&, Child1&, "_Oscar_IconBtn", vbNullString)
    child5& = FindWindowEx(Parent&, Child2&, "_Oscar_IconBtn", vbNullString)
child6& = FindWindowEx(Parent&, child3&, "_Oscar_IconBtn", vbNullString)
Call Click(child5&)
End Function
Public Sub SetText(Window As Long, text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub
Sub RunMenuByString(Application, StringSearch)
ToSearch% = GetMenu(Application)
MenuCount% = GetMenuItemCount(ToSearch%)

For FindString = 0 To MenuCount% - 1
ToSearchSub% = GetSubMenu(ToSearch%, FindString)
menuItemCount% = GetMenuItemCount(ToSearchSub%)

For GetString = 0 To menuItemCount% - 1
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
Function Chat_Clear(Txt As String)
'Ex: Call Chat_Clear("")
ChatRoom& = FindWindow("AIM_ChatWnd", vbNullString)
chatparent& = FindWindowEx(ChatRoom&, 0, "WndAte32Class", vbNullString)
OurHandle& = FindWindowEx(chatparent&, 0, "Ate32Class", vbNullString)
SetText OurHandle, ""
End Function
Function Chat_SendIm()
'Ex: Call Chat_SendIm
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ourparent& = FindWindowEx(AIMwindow&, 0, "_Oscar_TabGroup", vbNullString)
OurHandle& = FindWindowEx(ourparent&, 0, "_Oscar_IconBtn", vbNullString)
Call Icon(OurHandle)
End Function

Function Gotourl(url As String)
'Ex: Call Gotourl("http://www.deadbyte.com/flyman")
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
TheUrl& = FindWindowEx(AIMwindow&, 0, "Edit", vbNullString)
SetText TheUrl, url
DoEvents
Goicon& = FindWindowEx(AIMwindow&, 0, "_Oscar_IconBtn", vbNullString)
AIM_icon Goicon
End Function

Function SendIM(SN As String, message As String)
'Ex: Call SendIM("Flyman","Im Usin Your Bas")
Dim oscarpersistantcombo&
Dim oscarbuddylistwin&
Dim sendbuttonicon1&
Dim oscariconbtn2&
Dim gobuttonicon&
Dim oscariconbtn&
Dim wndateclass&
Dim aimimessage&
Dim ateclass&
Dim editx2&
Dim editx&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call SendMessageByString(editx&, WM_SETTEXT, 0&, "aim:goim")
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscariconbtn& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_iconbtn", vbNullString)
gobuttonicon& = SendMessage(oscariconbtn&, WM_LBUTTONDOWN, 0, 0&)
gobuttonicon& = SendMessage(oscariconbtn&, WM_LBUTTONUP, 0, 0&)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
editx2& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
Call SendMessageByString(editx2&, WM_SETTEXT, 0&, SN)
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call SendMessageByString(ateclass&, WM_SETTEXT, 0&, message$)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn2& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
sendbuttonicon1& = SendMessage(oscariconbtn2&, WM_LBUTTONDOWN, 0, 0&)
sendbuttonicon1& = SendMessage(oscariconbtn2&, WM_LBUTTONUP, 0, 0&)
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
editx& = FindWindowEx(oscarbuddylistwin&, 0&, "edit", vbNullString)
Call SendMessageByString(editx&, WM_SETTEXT, 0&, "*Search the Web*")
End Function
Function Chat_Link(Link As String, message As String)
'Ex: Call Chat_Link("http://www.deadbyte.com/flyman","FLYMAN 2000")
SendChat "<a href=""" + Link + """>" + message + ""
End Function
Function IM_Link(who As String, Link As String, message As String)
'Ex: Call IM_Link("Flyman","http://www.deadbyte.com/flyman",FLYMAN 2000")
SendIM who, "<a href=""" + Link + """>" + message + ""
End Function
Function ShowAim()
'Ex: Call ShowAim
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ShowWindow AIMwindow, SW_SHOW
End Function

Public Function GetText(WindowHandle As Long) As String
Dim buffer As String, TextLength As Long
TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
buffer$ = String(TextLength&, 0&)
Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
GetText$ = buffer$
End Function
Public Sub PlaySound(strFileName As String)
'Ex: C:Windows\Desktop\MyDocuments\Flyman.wav
sndPlaySound strFileName, 1
End Sub
Public Sub Button(mButton As Long)
Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Macro_Tits()
'Ex: Call Macro_Tits
SendChat ("  ;;;;;;;;;;;;;; ;;; ;;;;;;;;;;;;;; ;;;;;;;;;;")
SendChat ("      ;;;;;;     ;;;     ;;;;;;     ;;;¸¸¸¸¸¸¸")
SendChat ("      ;;;;;;     ;;;     ;;;;;;     ´´´´´´´;;;")
SendChat ("      ;;;;;;     ;;;     ;;;;;;    ;;;;;;;;;;; ")
SendChat ("                 By flyman")
End Sub
Sub Macro_Sperm()
'Ex: Call Macro_Sperm
SendChat ("`·.¸.·´¯`·.¸.·O   Sp3rm")
End Sub
Sub Macro_Pussy()
'Ex: Call Macro_Pussy
SendChat ("   ;;;;;;;;;;  ;;;    ;;;  ;;;;;;;;;;  ;;;;;;;;;; ;;;     ;;;")
SendChat ("   ¸¸¸¸¸¸¸;;;  ;;;    ;;;  ;;;¸¸¸¸¸¸¸  ;;;¸¸¸¸¸¸¸ ;;;¸¸¸¸¸;;;")
SendChat ("   ;;;´´´´´´´  ;;;    ;;;  ´´´´´´´;;;  ´´´´´´´;;;  ´´´´´´´´´")
SendChat ("   ;;;         ´;;;;;;;;´ ;;;;;;;;;;; ;;;;;;;;;;;    ;;;;;")
End Sub
Sub Macro_Phish()
'Ex: Call Macro_Phish
SendChat ("<B>></B><>")
End Sub
Sub Macro_Glass()
'Ex: Call Macro_Glass
Call SendChat("    ¸.·²'°'²·.¸_¸.·²'°'²·.¸¸.-·~²°˜¨")
Call SendChat("    `·.,¸¸,.·´  `·.,¸¸,.·´")
End Sub
Sub Macro_Flyman()
'Ex: Call Macro_Flyman
SendChat ("  ¸;;;;;;;;;;;      ;;;   ;;; ¸;;;¸  ¸;;;; ;;;;;;;;¸ ;;;;¸¸ ;;;")
SendChat ("  ;;;;;;;; ;;;      ´;;;;;;;; ;;;;;¸¸;;;;; ;;;;;;;;; ;;;´;;;;;;")
SendChat ("  ;;;      ´;;;;;;;;   ;;;    ;;; ;;;; ;;; ;;;;;;;;; ;;;  ´;;;;")
End Sub
Sub Macro_BRB()
'Ex: Call Macro_BRB
SendChat ("   ;;;;;;;;;;  ;;;;;;;;;;  ;;;;;;;;;;")
SendChat ("   ¸¸¸¸¸¸¸;;´  ¸¸¸¸¸¸¸;;;  ¸¸¸¸¸¸¸;;´")
SendChat ("   ;;;´´´´;;;  ;;;´´;;;;´  ;;;´´´´;;;")
SendChat ("   ;;;;;;;;;;  ;;;   ´;;;¸ ;;;;;;;;;;")
SendChat ("                       ´;;;¸")
SendChat ("                         ´´´´")
End Sub
Sub AddRoom_ToList(lis As ListBox)
'Ex: Call AddRoom_ToList(List1)
'Updated to work with AIM4.7
Dim ChatRoom As Long, LopGet, MooLoo, Moo2
    Dim name As String, NameLen, buffer As String
    Dim TabPos, NameText As String, text As String
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
            NameText$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
            TabPos = InStr(NameText$, Chr$(9))
            text$ = Right$(NameText$, (Len(NameText$) - (TabPos)))
            name$ = text$
            For mooz = 0 To lis.ListCount - 1
                If name$ = lis.List(mooz) Then
                    Well% = 123
                    GoTo Endz
                End If
            Next mooz
            If Well% <> 123 Then
                lis.AddItem name$
            Else
            End If
Endz:
        Next MooLoo
    End If

End Sub
Sub IM_Stamp_On()
'Ex: Call IM_Stamp_On
    Dim IMwin As Long
    IMwin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMwin&, "Timestamp")
End Sub
Sub IM_Stamps_Off()
'Ex: Call IM_Stamps_Off
    Dim IMwin As Long
    IMwin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMwin&, "Timestamp")
End Sub
Sub IM_Talk()
'Ex: Call IM_Talk
    Dim talkb As Long, FullWindow As Long, FullButton As Long, Klick As Long
    talkb& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(talkb&, "Connect to &Talk")
End Sub
Sub Click(TheIcon&)
'This was not coded by me,thanks to
'digitial
    Dim Klick As Long
    Klick& = SendMessage(TheIcon&, WM_LBUTTONDOWN, 0, 0&)
    Klick& = SendMessage(TheIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Clicks(item, clickmode As Integer)
Select Case clickmode
Case 1
Call SendMessageByNum(item, WM_LBUTTONDOWN, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONUP, 0, 0&)
Case 2
Call SendMessageByNum(item, WM_LBUTTONDOWN, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONUP, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONDOWN, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONUP, 0, 0&)
End Select
End Sub
Sub gobar(url$)
'thanks trend
Dim Parent As Long, Child1 As Long, Textset As Long, Child2 As Long, textset2 As Long
Parent& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "Edit", vbNullString)
Textset& = SendMessageByString(Child1&, WM_SETTEXT, 0, url$)
Child2& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call Click(Child2&)
textset2& = SendMessageByString(Child1&, WM_SETTEXT, 0, "Search The Web")
End Sub
Sub IM_GetInfo()
'Ex: Call IM_GetInfo
    Dim IMwin As Long
    IMwin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMwin&, "Info")
End Sub
Sub IM_Warn()
'Ex: Call IM_Warn
    Dim IMwin As Long, some As Long, Warn As Long, Click As Long
    IMwin& = FindWindow("AIM_IMessage", vbNullString)
    some& = FindWindowEx(IMwin&, 0, "_Oscar_IconBtn", vbNullString)
    Warn& = FindWindowEx(IMwin&, some&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(Warn&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(Warn&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub KillAd()
'Forget were this came from
'Ex: Call KillAd
Dim oscarbuddylistwin&
Dim wndateclass&
Dim ateclass&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
Call ShowWindow(ateclass&, SW_HIDE)
End Sub
Sub Show_GoToBar()
'Ex: Call Show_GoToBar
    Dim BuddyList As Long, STWbox As Long, GoButtin As Long
    Dim X  As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    X& = ShowWindow(STWbox&, SW_SHOW)
    X& = ShowWindow(GoButtin&, SW_SHOW)
End Sub
Sub Hide_GoToBar()
'Ex: Call Hide_GoToBar
    Dim BuddyList As Long, STWbox As Long, GoButtin As Long
    Dim X  As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    X& = ShowWindow(STWbox&, SW_HIDE)
    X& = ShowWindow(GoButtin&, SW_HIDE)
End Sub
Sub Form_Hide(frm As Form)
'Ex: Call Form_Hide(Form1)
frm.Hide
End Sub
Sub Form_Show(frm As Form)
'Ex: Call Form_Show
frm.Show
End Sub
Sub Form_Mini(frm As Form)
'Ex: Call Form_Mini(Form1)
frm.WindowState = 1
End Sub
Sub Form_Max(frm As Form)
'Ex: Call Form_Max(Form1)
frm.WindowState = 2
End Sub
Sub IM_Print()
'Ex: Call IM_Print
Dim aimimessage&
aimimessage& = FindWindow("aim_imessage", vbNullString)
Call RunMenuByString(aimimessage&, "&Print")
End Sub
Sub SignOnaFriend()
'Ex: Call SignOnaFriend
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call RunMenuByString(oscarbuddylistwin&, "&Sign On A Friend...")
End Sub
Sub NewsTicker()
'Ex: Call NewsTicker
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call RunMenuByString(oscarbuddylistwin&, "&News Ticker...")
End Sub

Sub HideMouse()
'Ex: Call HideMouse
Dim HideMouse$
HideMouse$ = ShowCursor(False)
End Sub

Sub ShowMouse()
'Ex: Call ShowMouse
Dim ShowMouse$
ShowMouse$ = ShowCursor(True)
End Sub
Sub List_AddItem(Lst As ListBox, TheItem$)
'Ex: Call List_AddItem(List1,"Whatever")
Lst.AddItem (TheItem$)
End Sub
Sub List_Clear(Lst As ListBox)
'Ex: Call List_Clear(List1)
Lst.Clear
End Sub
Sub Blist_Save()
'Ex: Call Blist_Save
Dim oscarbuddylistwin&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
Call RunMenuByString(oscarbuddylistwin&, "&Save Buddy List...")
End Sub
Sub Chat_Close()
'Ex: Call Chat_Close
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Call RunMenuByString(aimchatwnd&, "&Close...")
End Sub
Function Get_Roomname() As String
'Ex: Call Get_Roomname
Dim aimchatwnd&
Dim aimgetcaption$
Dim aimchangetext$
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
aimgetcaption$ = Get_Caption(aimchatwnd&)
aimchangetext$ = ReplaceString(aimgetcaption$, "Chat Room: ", "")
Get_Roomname = aimchangetext$
End Function
Sub Chat_Print()
'Ex: Call Chat_Print
Dim aimchatwnd&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
Call RunMenuByString(aimchatwnd&, "&Print...")
End Sub
Sub Bot_Attention(text$)
'Ex: Call Bot_Attention("I'm Usin FLYMANS Updated Bas!")
Call SendChat("<B>A T T E N T I O N</B>")
Call SendChat(text$)
Call SendChat("<B>A T T E N T I O N</B>")
End Sub
Public Sub Add_Name_To_List(List As ListBox, AddUser As Boolean)
'Ex: Call Add_Name_To_List(List1,"Person")
'Ex: Call Add_Name_To_List(List1,text1)
'Updated Sub
Dim Chat As Long, OTree As Long, OCount As Long
Dim OItem As Integer, OTlen As Long, OText As String
Chat& = FindChat
OTree& = FindWindowEx(Chat&, 0, "_Oscar_Tree", vbNullString)
OCount& = SendMessageByNum(OTree&, LB_GETCOUNT, 0, 0)
For OItem% = 0 To OCount& - 1
    OTlen = SendMessageByNum(OTree&, LB_GETTEXTLEN, OItem%, 0)
    OText$ = String(OTlen, 0)
    Call SendMessageByString(OTree&, LB_GETTEXT, OItem%, OText$)
    If AddUser = False And LCase(UserSN) = LCase(OText$) Then
    Else
    List.AddItem OText$
    End If
    DoEvents
Next
End Sub
Public Sub Form_Center(frm As Form)
'Ex: Call Form_Center
frm.Left = (Screen.Width / 2) - (frm.ScaleWidth / 2)
frm.Top = (Screen.Height / 2) - (frm.ScaleHeight / 2)
End Sub
Public Function Get_Chat_Text()
'I didn't make this.
Dim Atee As Long, Chat As Long, CText As String
Chat& = FindChat
Atee& = FindWindowEx(Chat&, 0, "WndAte32Class", vbNullString)
CText$ = GetText(Atee&)
GetchatText = CText$
End Function
Public Function Chat_Lastline()
'Ex: Call Chat_Lastline
Dim str As String, Fnd As Long
str$ = Get_Chat_Text
Fnd = InStrRev(str$, "<BR>")
If Fnd <> 0 Then
    str$ = Mid$(str$, Fnd + 4)
End If
Chat_Lastline = FilterHTML(str$)
End Function

Sub Chat_Macrokill_Smile()
'Ex: Call Chat_Macrokill_Smile
SendChat (":):-):-(;-):):-(:-):-(;-):):-):-(;-):):-):-(;-):):-):-(:):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):-):-(;-):):-):-(;-):):-):-(;-):):-(:-);-):):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-)")
SendChat (":):-):-(;-):):-(:-):-(;-):):-):-(;-):):-):-(;-):):-):-(:):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):-):-(;-):):-):-(;-):):-):-(;-):):-(:-);-):):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-)")
End Sub
Sub Pause(interval)
'Ex: Pause(1.0)
    Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
    DoEvents
    Loop
End Sub
Sub Chat_Sounds_Off()
'Ex: Call Chat_Sounds_Off
'thanks to digital
    Dim ChatWindow As Long, ZeeWin As Long, PrefWin As Long
    Dim Buttin2 As Long, Buttin As Long, PlayMess As Long
    Dim Buttin1 As Long, Buttin22 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, PlaySend As Long
    Dim OKbuttin As Long
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")
    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin&, "Button", vbNullString)
    PlayMess& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Call SendMessage(PlayMess&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlayMess&, WM_KEYUP, VK_SPACE, 0&)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin22& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin22&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    PlaySend& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Call SendMessage(PlaySend&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlaySend&, WM_KEYUP, VK_SPACE, 0&)
    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Chat_Sounds_On()
'Ex: Call Chat_Sounds_On
'thanks to digital
    Dim ChatWindow As Long, ZeeWin As Long, PrefWin As Long
    Dim Buttin2 As Long, Buttin As Long, PlayMess As Long
    Dim Buttin1 As Long, Buttin22 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, PlaySend As Long
    Dim OKbuttin As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin&, "Button", vbNullString)
    PlayMess& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Call SendMessage(PlayMess&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlayMess&, WM_KEYUP, VK_SPACE, 0&)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin22& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin22&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    PlaySend& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Call SendMessage(PlaySend&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlaySend&, WM_KEYUP, VK_SPACE, 0&)

    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Form_Protection()
'How this works:
'Add 1 TextBox
'Add 2 CommandButtons
'Add 1 Timer

'In timer1 put:
'Text1.Text = "Can't Get in"

'Name Commands 1 Caption: Enter
'In Command1 Put:
'If Text1.Text = "Flyman" Then Secret.Show
'Timer1.Enabled = True

'Name Commands ' Caption: Exit
'In Command2 Put:
'Send_Text("Im A Loser,Im Not Leet,No Secret Are For me")
'MsgBox "Nice Try Buddy"
'End
End Sub
Sub IM_Block()
'Ex: Call Im_Block
    Dim IMwin As Long, some As Long, Warn As Long, Block As Long
    Dim Click As Long
    IMwin& = FindWindow("AIM_IMessage", vbNullString)
    some& = FindWindowEx(IMwin&, 0, "_Oscar_IconBtn", vbNullString)
    Warn& = FindWindowEx(IMwin&, some&, "_Oscar_IconBtn", vbNullString)
    Block& = FindWindowEx(IMwin&, Warn&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(Block&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(Block&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub IM_Bold(SN$, text$)
'Ex: Call Im_Bold("Flyman","<b>What to say?</b>")
'Warning: I could have done it the real way
'but it would take more time
'Spank you trend
Dim Parent As Long, Child1 As Long
Call gobar("aim:goim?screenname=" & SN$ & "&message=" & text$)
Parent& = FindWindow("AIM_IMessage", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call Click(Child1&)
End Sub

Sub Bot_Verrifer(Txt As TextBox)
'you will need 1 Text Box and 2 Command Buttons
'in Command 1 Button (Check Button) put the following:
If Text1.text = "Flyman" Then
MsgBox "Yes You Found Me", vbOKOnly, "Don't Bug me about programming."
Else
MsgBox "Um.." + Text1.text + " Isn't me", vbOKOnly, "Thats A Poser! """
End If
'in the2nd Command Button (Exit Button) put the following:
Unload Me
End Sub
Sub Bot_FakeProg()
'Ex: Make 2 Text box's.
'1Command button
'Text 1 = the Progs name
'Text 2 = fake makers name
'In command 1 put:
'SendChat "-=(`(`" + Text1.Text + ""
'SendChat "-=(`(` By " + text2.Text + ""
'SendChat "-=(`(`Loaded
End Sub
Sub Clock(lbl As Label)
'Ex: Call Clock(Label1)
lbl.Caption = Time
End Sub
Sub Create_Directory(dir)
'Ex:
'Call Create_Directory("C:\WINDOWS\Desktop\File")
MkDir dir
End Sub
Public Sub Disable_Alt_Ctrl_Delete()
'Ex: Call Disable_Alt_Ctrl_Delete
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub Enable_Alt_Ctrl_Del()
'Ex: Call Enable_Alt_Ctrl_Delete
 Dim Ret As Integer
 Dim pOld As Boolean
 Ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub FileOpen_EXE(file$)
'Ex: Call FileOpen_EXE("C:Program Files\file.exe")
Openit! = Shell(file$, 1): NoFreeze% = DoEvents()
End Sub
Sub Not_On_Top(the As Form)
'Ex: Call Not_On_Top(Form1)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub IM_Blank()
'How to do it:
'In your command button put in:
'Call Send_Im(SN," ")
'just like that!
'easyer to do it this way and waste time to make a sub
End Sub
Sub Chat_Ignore(Person As String)
'Ex: Call Chat_Ignore(text1)
'or
'Ex: Call Chat_Ignore("Screenname")
Dim ChatRoom As Long, LopGet, MooLoo, Moo2
Dim name As String, NameLen, buffer As String
Dim TabPos, NameText As String, text As String
Dim mooz, Well As Integer, BuddyTree As Long
Person = LCase(Person)
ChatRoom = FindWindow("AIM_ChatWnd", vbNullString)
If ChatRoom <> 0 Then
Do
BuddyTree = FindWindowEx(ChatRoom, 0, "_Oscar_Tree", vbNullString)
Loop Until BuddyTree& <> 0
LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
For MooLoo = 0 To LopGet - 1
    Call SendMessageByString(BuddyTree, LB_SETCURSEL, MooLoo, 0)
    NameLen = SendMessage(BuddyTree, LB_GETTEXTLEN, MooLoo, 0)
    buffer = String(NameLen, "[]")
    Moo2 = SendMessageByString(BuddyTree, LB_GETTEXT, MooLoo, buffer)
    TabPos = InStr(buffer, Chr$(9))
    NameText = Right$(buffer, (Len(buffer) - (TabPos)))
    TabPos = InStr(NameText, Chr$(9))
    text = Right(NameText, (Len(NameText) - (TabPos)))
    text = Replace(text, " ", "")
    Person = Replace(Person, " ", "")
    If LCase(text) = LCase(Person) Then
     End If
Next MooLoo
End If
End Sub
Function Bot_Lamerizer(Nam As String)
'Ex: Call Bot_Lamerizer(text1)
'You gotta add a text box to enter who to lamerize.
'UPDATED SUB!
Call SendChat("·  LAMER'IZER FOUND " & Nam & "  ·")
Pause (2.6)
Dim X As Integer
Dim lcse As String
Dim letr As String
Dim dis As String
For X = 1 To Len(Nam)
lcse$ = LCase(Nam)
letr$ = Mid(lcse$, X, 1)
If letr$ = "a" Then Let dis$ = "a-is for the animals that you suck": GoTo Dissem
If letr$ = "b" Then Let dis$ = "b-is for all the boys you love": GoTo Dissem
If letr$ = "c" Then Let dis$ = "c-is for the cunt you are": GoTo Dissem
If letr$ = "d" Then Let dis$ = "d-is for all the times your dissed": GoTo Dissem
If letr$ = "e" Then Let dis$ = "e-is for that egghead of yours": GoTo Dissem
If letr$ = "f" Then Let dis$ = "f-is for the friday nights you stay home": GoTo Dissem
If letr$ = "g" Then Let dis$ = "g-is for the girls who hate you": GoTo Dissem
If letr$ = "h" Then Let dis$ = "h-is for the ho your momma is": GoTo Dissem
If letr$ = "i" Then Let dis$ = "i-is for the idiotic dumbass you are": GoTo Dissem
If letr$ = "j" Then Let dis$ = "j-is for all the times you jerkoff to your dog": GoTo Dissem
If letr$ = "k" Then Let dis$ = "k-is for you self esteem that the cool kids killed": GoTo Dissem
If letr$ = "l" Then Let dis$ = "l-is for the lame ass you are": GoTo Dissem
If letr$ = "m" Then Let dis$ = "m-is for the many men you sucked": GoTo Dissem
If letr$ = "n" Then Let dis$ = "n-is for the nights you spent alone": GoTo Dissem
If letr$ = "o" Then Let dis$ = "o-is for the sex operation you had": GoTo Dissem
If letr$ = "p" Then Let dis$ = "p-is for the times people p on you": GoTo Dissem
If letr$ = "q" Then Let dis$ = "q-is for the queer you are": GoTo Dissem
If letr$ = "r" Then Let dis$ = "r-is for all the times i raped your sister": GoTo Dissem
If letr$ = "s" Then Let dis$ = "s-is for your lover Steve Case": GoTo Dissem
If letr$ = "t" Then Let dis$ = "t-is for the tits youll never see": GoTo Dissem
If letr$ = "u" Then Let dis$ = "u-is for your underwear hangin on the flagpole": GoTo Dissem
If letr$ = "v" Then Let dis$ = "v-is for the victories you'll never have": GoTo Dissem
If letr$ = "w" Then Let dis$ = "w-is for the 400 pounds you wiegh":  GoTo Dissem
If letr$ = "x" Then Let dis$ = "x-is for all the lamers who" & Chr(34) & "[x]'ed" & Chr(34) & " you online": GoTo Dissem
If letr$ = "y" Then Let dis$ = "y-is for the question of, y your even alive?": GoTo Dissem
If letr$ = "z" Then Let dis$ = "z-is for zero which is what you are":  GoTo Dissem

If letr$ = "1" Then Let dis$ = "1-is for how many inches your dick is": GoTo Dissem
If letr$ = "2" Then Let dis$ = "2-is for the 2 dollars you make an hour": GoTo Dissem
If letr$ = "3" Then Let dis$ = "3-is for the amount of men your girl takes at once": GoTo Dissem
If letr$ = "4" Then Let dis$ = "4-is for your mom bein a whore":  GoTo Dissem
If letr$ = "5" Then Let dis$ = "5-is for 5 times an hour you whack off": GoTo Dissem
If letr$ = "6" Then Let dis$ = "6-is for the years you been single": GoTo Dissem
If letr$ = "7" Then Let dis$ = "7-is for the times your girl cheated on you..with me": GoTo Dissem
If letr$ = "8" Then Let dis$ = "8-is for how many people beat the hell outta you today": GoTo Dissem
If letr$ = "9" Then Let dis$ = "9-is for how many boyfriends your momma has": GoTo Dissem
If letr$ = "0" Then Let dis$ = "0-is for the amount of girls you get": GoTo Dissem
Dissem:
Call Send_Text(dis$)
Pause (1)
Next X
End Function
Sub About()
'Name of Bas: FLYMANS AIM MODULE 2002 Release 2

'Maker: FLYMAN

'Programed In: Visual Basic 6.0 Pro

'Compatable For: Visual Basic 5, Visual Basic 6 32bit

'AIM Version(s): 2.5-4.7

'My Site: FLYMAN 2000

'Url: http://www.deadbyte.com/flyman

'Relase Date: N/A

'Api Spy Used: Pat Or Jk 4.0,Mav Spy 4.0.

'Can I Use a sub?: Yes but please give me credit!

'Update Purpose: I got bored and people complained and reported
'bugs so here I am fixing the reported bugs.

'Subs: I don't know,maybe more then 114.
End Sub


Function List_Count(the As ListBox)
'Ex: Call List_count(List1)
Dim List$
List$ = the.ListCount
List_Count = List$
End Function
Sub MoveMouse(X1 As Long, Y1 As Long)
'Ex: Call MoveMouse(400,688)
Dim pointless
pointless = SetCursorPos(X1, Y1)
End Sub
Sub Chat_More()
'Ex: Call Chat_More
Dim aimchatwnd As Long
Dim Button As Long
aimchatwnd = FindWindow("aim_chatwnd", vbNullString)
Button = FindWindowEx(aimchatwnd, 0&, "button", vbNullString)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(Button, WM_LBUTTONUP, 0&, 0&)
End Sub
Sub Chat_Scroll(thetext As String)
'Ex: Call Chat_Scroll(Text1.text)
'or
'Ex: Call Chat_Scroll("HeH")
Call SendChat(thetext)
Pause 0.5
Call SendChat(thetext)
Pause 1
Call SendChat(thetext)
Pause 0.5
Call SendChat(thetext)
Pause 1
Call SendChat(thetext)
Pause 1
Call SendChat(thetext)
Pause 1
Call SendChat(thetext)
Pause 0.5
Call SendChat(thetext)
Pause 1
Call SendChat(thetext)
Pause 0.5
Call SendChat(thetext)
End Sub
Sub Chat_Sendblank()
'Ex: Call Chat_Sendblank
Call SendChat("  ")
End Sub
Sub Chat_Punt()
'Ex: Call Chat_Punt
'Note: This will work with old version of aim if their not
'fixed.(Aim examples: Aim3.5)
Call SendChat("&#770;")
End Sub
Sub IM_Distort(SN As String, Txt As String)
'Ex: Call IM_Distort("TheSN","")
Call SendIM(SN, "<a href=meow>:-*:-*:-*:-*:-*")
Call SendIM(SN, "<a href=meow>:-*:-*:-*:-*:-*")
Call SendIM(SN, "<a href=meow>:-*:-*:-*:-*:-*")
End Sub
Sub IM_KissFace(name As String, Txt As String)
'Ex: Call IM_KissFace("TheSN","")
Call SendIM(name, ":-*")
End Sub
Sub IM_Smile(name As String, Txt As String)
'Ex: Call IM_Smile("TheSN","")
Call SendIM(name, ":-)")
End Sub
Sub IM_Footinmouth(name As String, Txt As String)
'Ex: Call IM_Footinmouth("TheSN","")
Call SendIM(name, ":-!")
End Sub
Sub IM_Laughing(name As String, Txt As String)
'Ex: Call IM_Laughing("TheSN","")
Call SendIM(name, ":-D")
End Sub
Sub IM_Crying(name As String, Txt As String)
'Ex: Call IM_Crying("TheSN","")
Call SendIM(name, ":'")
End Sub
Sub IM_Embarassed(name As String, Txt As String)
'Ex: Call IM_Embarassed("TheSN","")
Call SendIM(name, ":-[")
End Sub
Sub IM_Surprise(name As String, Txt As String)
'Ex: Call IM_Surprise("TheSN","")
Call SendIM(name, "=-0")
End Sub
Sub Chat_Embarassed()
'Ex: Call Chat_Embarassed
Call SendChat(":-[")
End Sub
Sub Chat_Crying()
'Ex: Call Chat_Crying
Call SendChat(":'(")
End Sub
Sub Chat_Smile()
'Ex: Call Chat_Smile
Call SendChat(":")
End Sub
Sub Chat_Surprised()
'Ex: Call Chat_Surprised
Call SendChat("=-0")
End Sub
Sub Chat_Laughing()
'Ex: Call Chat_Laughing
Call SendChat(":-D")
End Sub
Sub Chat_Footinmouth()
'Ex: Call Chat_Footinmouth
Call SendChat(":-!")
End Sub
Sub Chat_Kissface()
'Ex: Call Chat_Kissface
Call SendChat(":-*")
End Sub
Sub IM_Hiddenmessage(SN As String, Txt As String, secret As String)
'Ex: Call IM_Hiddenmessage("TheSN","","The Message")
Call SendIM(SN, "<!--secret--!>")
End Sub
Sub Chat_Font(FontName As String)
'Trend is the creator of this, not GRN.
'Ex: Call Chat_Font("Hello!")
Call SendChat("<FONT FACE=FontName")
End Sub
Sub IM_Font(SN As String, FontName As String)
'Ex: Call IM_Font("TheSN","Hello!")
Call SendIM(SN, "<FONT FACE=FontName")
End Sub
Sub Chat_Punt2()
'Ex: Call Chat_Punt2
'Note: This will work with old version of aim if their not
'fixed.(Aim examples: Aim3.5)
Call SendChat("&.¤770;")
End Sub
Public Sub UnloadAllForms()
Dim Form As Form
   For Each Form In Forms
      Unload Form
      Set Form = Nothing
   Next Form
End Sub
Sub Special_Thanks()
' I would like to thank some people but I really don't know
' why I would thank when they should thank me for answers. (LOL)
' Anyways I would love to say I got my baby maryjanes growing
' its started on July 18th 2001. (Hemp Plants (Marijuana)).
End Sub
Sub CDOpen()
'Found this on the net
'Ex: Call CDOpen
Dim OpenCD$
OpenCD$ = mciSendString("set CDAudio door open", vbNullString, 0, 0)
End Sub
Sub CDClose()
'Found this on the net
'Ex: Call CDClose
Dim CloseCD$
CloseCD$ = mciSendString("set CDAudio door closed", vbNullString, 0, 0)
End Sub
Sub Form_Wheel(frm As Form)
Dim cx, cy, radius, Limit
frm.ScaleMode = 3
cx = frm.ScaleWidth / 2
cy = frm.ScaleHeight / 2
If cx > cy Then Limit = cy Else Limit = cx
For radius = 0 To Limit
frm.Circle (cx, cy), radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Next radius
End Sub
Sub DisableXonForm(beer)
'Put the following in the form load
Dim systemmenu%
Dim res%
systemmenu% = GetSystemMenu(beer, 0)
res% = RemoveMenu(systemmenu%, 6, MF_BYPOSITION)
End Sub
Function EnterKey()
EnterKey = CStr(Chr(13) & Chr(10))
End Function
Public Sub FileHidden(TheFile As String)
'This makes a certain file hidden
'Ex: Call FileHidden ("C:\TheProgram.exe")
Dim SafeFile As String
SafeFile$ = dir(TheFile$)
If SafeFile$ <> "" Then
SetAttr TheFile$, vbHidden
End If
End Sub
Function GetCharCount(text As String) As String
'Thought this would be useful
GetCharCount = Len(text)
End Function
Sub Ad_Change(text As String)
'Ex: Call Ad_Change("Changed AD")
'Updated Sub.
Dim Parent As String
Parent& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "WndAte32Class", vbNullString)
Child2& = FindWindowEx(Child1&, 0&, "Ate32Class", vbNullString)
Textset& = SendMessageByString(Child2&, WM_SETTEXT, 0, text$)
End Sub
Function Get_Caption(TheWin)
Dim WindowLngth As Integer, WindowTtle As String, Moo As String
WindowLngth% = GetWindowTextLength(TheWin)
WindowTtle$ = String$(WindowLngth%, 0)
Moo$ = GetWindowText(TheWin, WindowTtle$, (WindowLngth% + 1))
Get_Caption = WindowTtle$
End Function
Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String

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
Function Get_Text(child)
Dim GetTrim As Integer, TrimSpace As String, GetString As String
GetTrim% = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString$ = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
Get_Text = TrimSpace$
End Function
Function HTML_Remove(Stg As String) As String
'This is from Raid's bas
Dim i, beg As String, ends As String, after As String, junk As String, junk2 As String, before As String
Dim TheStrg As String
TheStrg$ = ReplaceString(Stg$, "<BR>", "" & Chr$(13) + Chr$(10))
For i = 1 To Len(Stg)
beg$ = InStr(1, Stg, "<")
If beg$ = 0 Then
GoTo Endz:
Else
ends$ = InStr(1, Stg, ">")
End If
If ends$ = 0 Then
GoTo Endz:
Else
after$ = Mid$(Stg$, ends$ + 1, Len(Stg$) - ends$)
junk$ = Len(Stg$) - (beg$ - 1)
junk2$ = Len(Stg$) - junk$
before$ = Mid$(Stg$, 1, junk2$)
Stg$ = before$ & after$
End If
Next i
Endz:
HTML_Remove = Stg$
End Function

Function Chat_Get_Text()
'Gets text from the chatroom.
Dim aimimessage2$
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
aimimessage2$ = Get_Text(ateclass&)
Chat_Get_Text = aimimessage2$
End Function
Public Sub FormDrag(TheForm As Form)
'Ex: Call FormDrag Me
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Sub ChatCrash_link()
'Ex: Call ChatCrash_link
Call Chat_Link("File:///C:/con/con", "Get Paid 10.00 each referal")
End Sub
Sub ChatCrash_link2()
'Ex: Call ChatCrash_link2
Call Chat_Link("File:///C:/aux/aux", "Get Paid 10.00 each referal")
End Sub
Function New_IM()
'Ex: Call New_IM
Dim aim As Long, Group As Long, Button As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
Click (Button&)
End Function
Sub Form_On_Top(the As Form)
'Updated Sub
SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
