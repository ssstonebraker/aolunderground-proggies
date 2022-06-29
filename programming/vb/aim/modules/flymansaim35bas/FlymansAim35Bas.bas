Attribute VB_Name = "FlymansAim35Bas"
Option Explicit
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFilename As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFilename As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
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
Declare Function iswindowenabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function PutFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
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
Public Const SW_ShowMinimized = 2
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
Public Const ENTA = 13
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const EM_LINESCROLL = &HB6
Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Type POINTAPI
   x As Long
   Y As Long
End Type
Function AboutBas()
'this is for aim3.5...nothing great.
'this was used for small toolz 3.0!
'I'd like to really thank
'Quirk,because of he fact that he
'makes me impressed by his work
'so i decide to do that stuff too
'so thanks to him!
'Also I'd like to thank KnK
'for his tight ass page that keeps
'on rocking!
'Plus Everyone Else!
End Function
Sub Clear_Chat(text$)
'Ex: Call Clear_Chat("")
Dim parent As Long, child1 As Long, child2 As Long, child3 As Long, textset As Long
parent& = FindWindow("AIM_ChatWnd", vbNullString)
child1& = FindWindowEx(parent&, 0&, "WndAte32Class", vbNullString)
child2& = FindWindowEx(child1&, 0&, "Ate32Class", vbNullString)
textset& = SendMessageByString(child2&, WM_SETTEXT, 0, text$)
End Sub
Sub Send_Text(text$)
'this goes to a chat room
'Ex:Call Send_Text(Text1)
'or
'Ex:Call Send_Text("The Text here")
Dim parent As Long, child1 As Long, child2 As Long, child3 As Long, child4 As Long, child5 As Long, child6 As Long, textset As Long
    parent& = FindWindow("AIM_ChatWnd", vbNullString)
    child1& = FindWindowEx(parent&, 0&, "WndAte32Class", vbNullString)
    child2& = FindWindowEx(parent&, child1&, "WndAte32Class", vbNullString)
    textset& = SendMessageByString(child2&, WM_SETTEXT, 0, text$)
    child3& = FindWindowEx(parent&, 0&, "_Oscar_IconBtn", vbNullString)
    child4& = FindWindowEx(parent&, child1&, "_Oscar_IconBtn", vbNullString)
    child5& = FindWindowEx(parent&, child2&, "_Oscar_IconBtn", vbNullString)
child6& = FindWindowEx(parent&, child3&, "_Oscar_IconBtn", vbNullString)
Call click(child5&)
End Sub
Sub click(TheIcon&)
'This was not coded by me,thanks to
'digitial
    Dim Klick As Long
    Klick& = SendMessage(TheIcon&, WM_LBUTTONDOWN, 0, 0&)
    Klick& = SendMessage(TheIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Im_Direct(sn As String)
'Ex: Call Im_Direct(text1)
Dim parent As Long, child1 As Long
Call gobar("aim:goim?screenname=" & sn$)
parent& = FindWindow("AIM_IMessage", vbNullString)
child1& = FindWindowEx(parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call RunMenuByString(parent, "Send IM I&mage")
End Sub
Sub Get_Files(sn$)
'Ex:Call Get_Files(Text1)
'or
'Ex:Call Get_Files("")
Call gobar("aim:getfile?screenname=" & sn$)
End Sub
Sub gobar(url$)
Dim parent As Long, child1 As Long, textset As Long, child2 As Long, textset2 As Long
parent& = FindWindow("_Oscar_BuddyListWin", vbNullString)
child1& = FindWindowEx(parent&, 0&, "Edit", vbNullString)
textset& = SendMessageByString(child1&, WM_SETTEXT, 0, url$)
child2& = FindWindowEx(parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call click(child2&)
textset2& = SendMessageByString(child1&, WM_SETTEXT, 0, "Search The Web")
End Sub
Sub Clear_Im(text$)
'Ex Call Clear_Im("")
Dim parent As Long, child1 As Long, child2 As Long, textset As Long, send As Long
parent& = FindWindow("AIM_IMessage", vbNullString)
child1& = FindWindowEx(parent&, 0&, "WndAte32Class", vbNullString)
child2& = FindWindowEx(child1&, 0&, "Ate32Class", vbNullString)
textset& = SendMessageByString(child2&, WM_SETTEXT, 0, text$)
End Sub
Sub Send_Im(sn$, text$)
'Ex: Call Send_Im(Text1,"Your Text")
'or
'Ex: Call Send_Im(Text1,Text2)
'or
'Ex: Call Send_Im("Tourq","Hey I am Useing your .bas")
'or
'Ex: Call Send_Im("Tourq",Text1)
Dim parent As Long, child1 As Long
Call gobar("aim:goim?screenname=" & sn$ & "&message=" & text$)
parent& = FindWindow("AIM_IMessage", vbNullString)
child1& = FindWindowEx(parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call click(child1&)
End Sub
Sub RunMenuByString(Application, StringSearch)
'I got this from someone!

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
Sub File_Send(sn As String)
'Ex: Call File_Send(Text1)
'or
'Ex: Call File_Send("")
Dim parent As Long, child1 As Long
Call gobar("aim:goim?screenname=" & sn & "")
parent& = FindWindow("AIM_IMessage", vbNullString)
child1& = FindWindowEx(parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call RunMenuByString(parent, "Send &File")
End Sub
Sub DirectIm_Close(sn As String)
'Ex: Call DirectIm_Close(Text1)
'or
'Ex: Call DirectIm_Close("")
Dim parent As Long, child1 As Long
Call gobar("aim:goim?screenname=" & sn$)
parent& = FindWindow("AIM_IMessage", vbNullString)
child1& = FindWindowEx(parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call RunMenuByString(parent, "&Close IM Image")
End Sub
Function Attention(text$)
'Ex:Call Attention(text1)
Send_Text "-=\[A.T.T.E.N.T.I.O.N]/=-_"
Send_Text Message$
Send_Text "-=\[A.T.T.E.N.T.I.O.N]/=-_"
End Function

Function Blue_Link(Link$, Message$)
'Ex:Call Blue_Link("Http://www.deadbyte.com","My Page")
'or
'Ex:Call Blue_Link(Text1,Text2)
Send_Text "<a href=""" + Link$ + """><font color=#0000ff>" + Message$ + ""
End Function
