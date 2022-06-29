Attribute VB_Name = "ItZeD"
'This is ItZeD.bas.  the easyiest bas for AOL programming
'This bas made by Mr Ed
'Mail me at danonutzmeestaed@finalpeace.com OR mredgotnonuts@hotmail.com
'this bas is being used to make the prog/punter  Buh Bye 1 IM
'I took some of the subz and functions out.  if I left them it would be way to easy for you
'people to make a GOOD punter like Buh Bye 1 IM.  like close the IM Reflectoin
'IF you want to IM me I can be IMed at "imredl" or it loox like this "IMreDl"
'I don't sign on much though
'I took the ONLY fader sub OUT.  if you want it rite it YOUR self god damnit!!!
'and for a copy of Buh Bye 1 IM when I FINISH it mail me and I'll send you a copy!
'It will work for all versions of AOL from 2.5 up and have errors for each to.  hopefully
'even AOL 4s new fix, I'm working on it now!!!  I took the ALL AOL sendchat out and added
'a different sendchat for AOL4 only.  this bas is AOL 4 only, sorry ;(
'I added a few "unneeded" subs and functions because I know people will do a lot of editing
'to this bas even though I DON'T want them to I KNOW they WILL.  And I also know that they
'will use this bas with others and EVERYBODY (even me) forgets sometimes, I think it works
'FULLY.  PERFECT.  I think, well, bye
'P.S. I tried to add a little something to tell what the things do but some things don't
'have one, either you probaly know it, or ask somebody what it does

'                              ApI FuncTiOnS!
'             user32.dll
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function GetNextWindow Lib "user32" (ByVal hwnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'              Kernel32

Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'           dwspy32.dll

Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)

'           shell32.dll

Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpsSoundName As String, ByVal uFlags As Long) As Long

'      BM_ and WM_ Public Constants

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_MDIDESTROY = &H221
Public Const WM_NCLBUTTONDOWN = &HA1

'        LB_ and VK_ Public Constants

Public Const LB_GETTEXT = &H189
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_GETCURSEL = &H188
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
Public Const VK_ESCAPE = &H1B
Public Const VK_LCONTROL = &HA2

'       GW_, SW_,MF_, and other Public Constants

Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2

Public Const GW_DELTA = 2
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4
Public Const GW_CHILD = 5
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


Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_APPEND = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_BYCOMMAND = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&


Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000


Type POINTAPI
   X As Long
   Y As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Function GetClass(child)
'get the class
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Function FindChildByTitle(parentw, childhand)
'find a child
ed% = GetWindow(parentw, 5)
If UCase(GetCaption(ed%)) Like UCase(childhand) Then GoTo master
ed% = GetWindow(parentw, GW_CHILD)

While firs%
edme% = GetWindow(parentw, 5)
If UCase(GetCaption(edme%)) Like UCase(childhand) & "*" Then GoTo master
ed% = GetWindow(ed%, 2)
If UCase(GetCaption(ed%)) Like UCase(childhand) & "*" Then GoTo master
Wend
FindChildByTitle = 0

master:
EdRocks% = ed%
FindChildByTitle = EdRocks%
End Function
Sub SendCharNum(win, chars)
E = SendMessageByNum(win, WM_CHAR, chars, 0)
End Sub

Function GetCaption(hwnd)
'get the caption of hWnd
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Function FindChildByClass(parentw, childhand)
'find child by class duh, the name says it all
ed% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(ed%), 1, Len(childhand))) Like UCase(childhand) Then GoTo EdRocks
ed% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(ed%), 1, Len(childhand))) Like UCase(childhand) Then GoTo EdRocks

While ed%
edme% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(edme%), 1, Len(childhand))) Like UCase(childhand) Then GoTo EdRocks
ed% = GetWindow(ed%, 2)
If UCase(Mid(GetClass(ed%), 1, Len(childhand))) Like UCase(childhand) Then GoTo EdRocks
Wend
FindChildByClass = 0

EdRocks:
IMreDl% = firs%
FindChildByClass = IMreDl%

End Function

Sub SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Sub
Sub AddRoomToListBox(ListBox As ListBox)
'Adds the room to the listbox you choose
'i.e. Call AddRoomToListBox (list1)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
TheList.Clear

Room = FindChatroom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
If Person$ = UserSN Then GoTo Na
ListBox.AddItem Person$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function


Function AOLMDI()
'AOLMDI duh
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Sub SetText(win, txt)
'sets the TXT to the win you choose
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub

Sub ClickIcon(icon%)
'clicks a icon, just define it!
'easyier than typing the below out everytime you want to click a icon!
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

'you gotta have the basics first!
Function FindChatroom()
'find the chat room, good for a room buster
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatroom = Room%
Else:
   FindChatroom = 0
End If
End Function

Function TxtFromChat()
'get the text from a chat
FindRoom% = FindChatroom
AOL% = FindChildByClass(FindRoom%, "RICHCNTL")
chattext = GetText(AOL%)
TxtFromChat = chattext
End Function

Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
'I should have put this earlyier, but oh well, it adds from the list box to the combo, so
'make a list box, List1 and a combobox Combo1, then hide the list box and this will work
Call AddRoomToListBox(ListBox)
For abc = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(abc))
Next abc
End Sub

Sub Upchat()
'This allows the user to upload and chat at the same time, kinda nice
AOL% = FindWindow("AOL Frame25", vbNullString)
Modal% = FindChildByClass(AOL%, "_AOL_Modal")
Gauge% = FindChildByClass(Modal%, "_AOL_Gauge")
If Gauge% <> 0 Then Upp% = Modal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Sub UnUpchat()
'Stops the "UpChat" above, stops upload and chat
AOL% = FindWindow("AOL Frame25", vbNullString)
Modal% = FindChildByClass(AOL%, "_AOL_Modal")
Gauge% = FindChildByClass(Modal%, "_AOL_Gauge")
If Gauge% <> 0 Then Upp% = Modal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub
Function UserSN()
'This neeto function gets the users screen name from the welcome menu, good for load and
'unload on your progs
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function



Sub SendChat(txt)
'send text to chat
Room% = FindChatroom
AOL% = FindChildByClass(Room%, "RICHCNTL")


AOL% = GetWindow(AOL%, 2)
AOL% = GetWindow(AOL%, 2)
AOL% = GetWindow(AOL%, 2)
AOL% = GetWindow(AOL%, 2)
AOL% = GetWindow(AOL%, 2)
AOL% = GetWindow(AOL%, 2)

Call SetFocusAPI(AOL%)
Call SendMessageByString(AOL%, WM_SETTEXT, 0, txt)
DoEvents
Call SendMessageByNum(AOL%, WM_CHAR, 13, 0)
End Sub
Sub SendMail(Recipiants, subject, message)
'Sends mail to the person(s) you want with the subject and message you want to
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

Sub SendIM(Person, message)
'send a IM, I used sendkeys because it makes it a LOT easyier on me, it makes it easyier to
'send a IM on all versions of AOL rather than makeing a sub that sends on all that is HUGE
'even if I did make one like that I wouldn't leave it in the bas
AppActivate "America  Online"
SendKeys ("^i")
Do: DoEvents
    MDI% = AOLMDI
    IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
    who% = FindChildByClass(IMWin%, "_AOL_Edit")
    mess% = FindChildByClass(IMWin%, "RICHCNTL")
    Sender% = GetWindow(mess%, 2)
    KWin% = FindChildByTitle(AOLMDI, "Keyword")
    X = SendMessage(KWin%, WM_CLOSE, 0, 0)
 If IMWin% <> 0 And who% <> 0 And mess% <> 0 And Sender% <> 0 And KWin% = 0 Then Exit Do
Loop
Pause 0.1
DoEvents
Call SetText(who%, Person)
Call SetText(who%, Person)
Call SetText(who%, Person)
Pause 0.1
Call SetText(mess%, message)
Call SetText(mess%, message)
Call SetText(mess%, message)
ClickIcon (Sender%)
End Sub

Sub keyword(keyword)
'go to a keyword
AOL% = AOLWindow
MDI% = AOLMDI
Do:
DoEvents
Tool% = FindChildByClass(AOL%, "AOL Toolbar")
Tool2% = FindChildByClass(Tool%, "_AOL_Toolbar")
E% = FindChildByClass(Tool2%, "_AOL_Glyph")
a% = GetWindow(E%, 2)
b% = GetWindow(a%, 2)
c% = GetWindow(b%, 2)
d% = GetWindow(c%, 2)
E2% = GetWindow(d%, 2)
F% = GetWindow(E2%, 2)
If Tool% <> 0 And Tool2% <> 0 And E% <> 0 And a% <> 0 And b% <> 0 And c% <> 0 And d% <> 0 And E2% <> 0 And F% <> 0 Then Exit Do
Loop
ClickIcon (F%)

Do:
DoEvents
wow% = FindChildByTitle(MDI%, "Keyword")
Money% = FindChildByClass(wow%, "_AOL_Edit")
hey% = FindChildByClass(wow%, "_AOL_Icon")
If Cwow% <> 0 And Money% <> 0 And hey% <> 0 Then Exit Do
Loop
TimeOut 0.2
Call SetText(Money%, keyword)
Do: DoEvents
    Chill% = FindChildByTitle(MDI%, "Keyword")
    Call ClickIcon(hey%)
Loop Until Chill% = 0
End Sub


Sub TimeOut(Duration)
'pause your prog for the desired time
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub

Sub StayOnTop(TheForm As Form)
'make your form the top window
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
