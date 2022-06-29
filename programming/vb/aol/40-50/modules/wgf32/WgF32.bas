Attribute VB_Name = "Wgf32"
'                                 For Aol4.o, Made in VB5.o By:\/\/gF
'Sup? This was made with the stuff I like. I think all the stuff works if something dosn't
'e-mail me and I will fix it on the next update of it. I will not be updating it
'a lot because when I do I want to have big changes just not little things.
'This is the last style to my .bas and last name for it I will just keep on adding
'stuff to it like jag.bas* So e-mail me with requests for it at realwgf@hotmail.com
'Also all the main cool chat stuff is at the bottom of the .bas
                                                    'L8ter WgF
                                                    
                                                    
                                                    
                                                    
                                                    
                                                    
'This .bas was made by me anything that has likness to anyone eleses bas was completly accidential. My freinds told me what to put. exept for the exitav which was modified from the nuclear bas. because i liked it a lot
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
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
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

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
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type

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

Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function FindChatRoom()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Room% = FindChildByClass(MDI%, "AOL Child")
STUFF% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If STUFF% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function
Function HyperLink(Txt As String, URL As String)
'Send A Link To The ChatRoom
HyperLink = ("<A HREF=" & Chr$(34) & Text2 & Chr$(34) & ">" & Text1 & "</A>")
End Function
Sub KillDupes(lst As ListBox)
'Killz Dupes In A ListBoX
For X = 0 To lst.ListCount - 1
Current = lst.List(X)
For i = 0 To lst.ListCount - 1
Nower = lst.List(i)
If i = X Then GoTo dontkill
If Nower = Current Then lst.RemoveItem (i)
dontkill:
Next i
Next X
End Sub
Function UserSN()
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

Sub KillWait()

AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function IsUserOnline()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome,")
If welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Sub SendChat(Chat)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub
Sub pause(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop
End Sub

Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Function RGBtoHEX(RGB)
    a = Hex(RGB)
    B = Len(a)
    If B = 5 Then a = "0" & a
    If B = 4 Then a = "00" & a
    If B = 3 Then a = "000" & a
    If B = 2 Then a = "0000" & a
    If B = 1 Then a = "00000" & a
    RGBtoHEX = a
End Function
Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AOLMakeMeParent(frm As Form)
'this makes the form an aol parent
AOL% = FindChildByClass(FindWindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = SetParent(frm.hwnd, AOL%)
End Sub
Function AOLMDI()
'AOL MDI Window
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Function AOLVersion()
'Tells if AOL Version 4 or 3
hMenu% = GetMenu(AOLWindow())
SubMenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(SubMenu%, 8)
MenuString$ = String$(100, " ")
FindString% = GetMenuString(SubMenu%, subitem%, MenuString$, 100, 1)
If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 4
End If
End Function

Sub XScroll15(Txt As TextBox)
'Max of 14 chr or else u get Msg is too long
Call SetFocus
a = String(116, Chr(32))
d = 116 - Len(Txt)
c$ = Left(a, d)
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.8
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.8
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.8
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.8
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.3
SendChat "" + Txt.Text + "" & c$ & "" + Txt.Text + ""
pause 0.8
End Sub
Function XBoldRedBlackRed(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    RedBlackRed = Msg
XBoldSendChat (Msg)
End Function
Function XBoldPurpleBlackPurple(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
    XBoldPurpleBlackPurple = Msg
XBoldSendChat (Msg)
End Function
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Public Function ChangeAOLsCaption(caption As String) As String
    Dim sup As Long
    sup& = FindWindow("AOL Frame25", vbNullString)
    ChangeCaption cap&, caption$
End Function
Sub ClickNext()
'click the next button

mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 5
AOIcon% = GetWindow(AOIcon%, 2)
Next l
ClickIcon (AOIcon%)
End Sub
Sub ClickRead()
'click the read button

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 0
AOIcon% = GetWindow(AOIcon%, 2)
Next l
ClickIcon (AOIcon%)
End Sub
Public Function FileGetAttributes(TheFile As String) As Integer
'gets Attributes Of A File
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function
Sub XSupaclearChat()
'Clears Chat for every1
Call SendChat("<PRE<")
End Sub
Sub SendMail(Recipiants, subject, Message)

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
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

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
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Sub KeyWord(TheKeyWord As String)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
' If you have Killed the Glyph sub then
' the keyword icon is the 19th icon and you must use the
' code below
'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Function WinCaption(win)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
getWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

Sub IMBuddy(Recipiant, Message)
'Sends IM from bud list
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For l = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next l

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, Message)
'Sends IM by keyword
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Call KeyWord("aol://9293:")
Do: DoEvents
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Message)

For X = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next X

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function

Function GetchatText()
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")
chattext = GetText(AORich%)
GetchatText = chattext
End Function

Function LastChatLineWithSN()
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

Function SNFromLastChatLine()
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        sn = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = sn
End Function
Function LastChatLine()
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear

Room = FindChatRoom()
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

Sub XUnderLineSendChat(TheText)
'It underlines chat text.
SendChat ("<u>" & TheText & "</u>")
End Sub
Sub XFormExitUp(Form As Form)
Do
Form.Top = Trim(Str(Int(Form.Top) - 300))
DoEvents
Loop Until Form.Top < -4500
If Form.Top < -4500 Then End
End Sub
Public Sub XFormexplode(Form As Form, Movement As Integer)

    Dim myRect As RECT
    Dim formWidth As Integer, formHeight As Integer, i As Integer
    Dim X As Integer, Y As Integer
    Dim cx As Integer, cy As Integer
    Dim TheScreen As Long, Brush As Long
    GetWindowRect Form.hwnd, myRect
        formWidth% = (myRect.Right - myRect.Left)
        formHeight% = myRect.Bottom - myRect.Top
            TheScreen& = GetDC(0)
            Brush& = CreateSolidBrush(Form.BackColor)
            For i% = 1 To Movement%
                cx% = formWidth * (i% / Movement%)
                cy% = formHeight * (i% / Movement%)
                X% = myRect.Left + (formWidth% - cx%) / 2
                Y = myRect.Top + (formHeight% - cy%) / 2
                Rectangle TheScreen, X%, Y%, X% + cx%, Y% + cy%
            Next i%
                X% = ReleaseDC(0, TheScreen&)
                DeleteObject (Brush&)
'took this from my friend OO7'S .bas, thanks to him
End Sub
Public Sub XFormExitFav(Form As Form)
    Dim doit As Integer, Go As Long
    Go& = Form.Height / 2
    For doit% = 1 To Go&
    DoEvents
        Form.Height = Form.Height - 10
        Form.Top = (Screen.Height - Form.Height) \ 2
        If Form.Height <= 11 Then GoTo Ending
    Next doit%
Ending:
        Form.Height = 30
        Go& = Form.Width / 2
    For doit% = 1 To Go&
    DoEvents
        Form.Width = Form.Width - 10
        Form.Left = (Screen.Width - Form.Width) \ 2
        If Form.Width <= 11 Then Exit Sub
    Next doit%
    Unload Form
'this is the effect I like the best
End Sub
Public Sub XFormMove(Form As Form)
    'Move Form With no title bar
    Call ReleaseCapture
    Call SendMessage(Form.hwnd, &H112, &HF012, 0)
End Sub
Sub XFormExitLeft(Form As Form)
Do
Form.Left = Trim(Str(Int(Form.Left) - 300))
DoEvents
Loop Until Form.Left < -6300
If Form.Left < -6300 Then End
End Sub
Sub XFormExitRight(Form As Form)
Do
Form.Left = Trim(Str(Int(Form.Left) + 300))
DoEvents
Loop Until Form.Left > 9600
If Form.Left > 9600 Then End
End Sub
Sub XFormExitDown(Form As Form)
Do
Form.Top = Trim(Str(Int(Form.Top) + 300))
DoEvents
Loop Until Form.Top > 7200
If Form.Top > 7200 Then End
End Sub


Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
Sub XFormCenter(F As Form)
'CenterS The Form

F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Sub ChangeSNInWelcomeWindow(Wha As TextBox)
'Changes da SN in welcome window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
Call SendMessageByString(welcome%, WM_SETTEXT, 0, "Welcome," & " " & Wha)
End Sub
Sub XClearChat()
'Clear chat for only user
childs% = FindChatRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Sub ClickForward()
'Clicks The Forward Button

mailwin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(mailwin%, "_AOL_Icon")
For l = 1 To 8
AOIcon% = GetWindow(AOIcon%, 2)
NoFreeze% = DoEvents()
Next l
ClickIcon (AOIcon%)
End Sub
Sub ClickKeepAsNew()
'clicks the keep as new button

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For l = 1 To 2
AOIcon% = GetWindow(AOIcon%, 2)
Next l
ClickIcon (AOIcon%)
End Sub

Sub XFormBlueFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
     vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub
Sub XFormPurpleFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
     vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(25, 0, 100 - intLoop), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub

Sub XFormGreenFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
    vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub XFormRedFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
     vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub XFormFireFade(vForm As Object)
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
Sub XFormSilverFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B 'Draw boxes with specified color of loop
    Next intLoop
End Sub
Sub XFormIceFade(vForm As Object)
    'This code works best when called in the paint event
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 255), B 'Draw boxes with specified color of loop
        Next intLoop
End Sub
Sub XWavyBlueBlack(TheText)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
SendChat (P$)
End Sub
Function XTextHacker(word$)

Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    If letter$ = "a" Then Leet$ = "a"
    If letter$ = "b" Then Leet$ = "B"
    If letter$ = "c" Then Leet$ = "C"
    If letter$ = "d" Then Leet$ = "D"
    If letter$ = "e" Then Leet$ = "e"
    If letter$ = "f" Then Leet$ = "F"
    If letter$ = "g" Then Leet$ = "G"
    If letter$ = "h" Then Leet$ = "H"
    If letter$ = "i" Then Leet$ = "i"
    If letter$ = "j" Then Leet$ = "J"
    If letter$ = "k" Then Leet$ = "K"
    If letter$ = "l" Then Leet$ = "L"
    If letter$ = "m" Then Leet$ = "M"
    If letter$ = "n" Then Leet$ = "N"
    If letter$ = "o" Then Leet$ = "o"
    If letter$ = "p" Then Leet$ = "P"
    If letter$ = "q" Then Leet$ = "Q"
    If letter$ = "r" Then Leet$ = "R"
    If letter$ = "s" Then Leet$ = "S"
    If letter$ = "t" Then Leet$ = "T"
    If letter$ = "u" Then Leet$ = "u"
    If letter$ = "v" Then Leet$ = "V"
    If letter$ = "w" Then Leet$ = "W"
    If letter$ = "x" Then Leet$ = "X"
    If letter$ = "y" Then Leet$ = "y"
    If letter$ = "z" Then Leet$ = "Z"
    If letter$ = "A" Then Leet$ = "a"
    If letter$ = "B" Then Leet$ = "B"
    If letter$ = "C" Then Leet$ = "C"
    If letter$ = "D" Then Leet$ = "D"
    If letter$ = "E" Then Leet$ = "e"
    If letter$ = "F" Then Leet$ = "F"
    If letter$ = "G" Then Leet$ = "G"
    If letter$ = "H" Then Leet$ = "H"
    If letter$ = "I" Then Leet$ = "i"
    If letter$ = "J" Then Leet$ = "J"
    If letter$ = "K" Then Leet$ = "K"
    If letter$ = "L" Then Leet$ = "L"
    If letter$ = "M" Then Leet$ = "M"
    If letter$ = "N" Then Leet$ = "N"
    If letter$ = "O" Then Leet$ = "o"
    If letter$ = "P" Then Leet$ = "P"
    If letter$ = "Q" Then Leet$ = "Q"
    If letter$ = "R" Then Leet$ = "R"
    If letter$ = "S" Then Leet$ = "S"
    If letter$ = "T" Then Leet$ = "T"
    If letter$ = "U" Then Leet$ = "u"
    If letter$ = "V" Then Leet$ = "V"
    If letter$ = "W" Then Leet$ = "W"
    If letter$ = "X" Then Leet$ = "X"
    If letter$ = "Y" Then Leet$ = "y"
    If letter$ = "Z" Then Leet$ = "Z"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
XTextHacker = Made$
End Function
Sub XTEliteTalker(word$)
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

Sub IMsOn()
Call IMKeyword("$IM_ON", " ")
End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", " ")
End Sub
Sub Attention(TheText As String)
XBoldSendChat ("XxXxXx ATTENTIONxXxXxX")
Call TimeOut(0.15)
XBoldSendChat (TheText)
Call TimeOut(0.15)
xboldsendchatSendChat ("XxXxXx ATTENTIONxXxXxX")
Call TimeOut(0.15)
End Sub
Sub CATWatchBot()
'CAT watch bot
CATWatchStop = False
sstart:
If CATWatchStop = True Then Exit Sub
Dim sTheSN As String
Dim iDiff As Integer
Dim sChatLine As String
If SNFromLastChatLine() = "OnlineHost" Then
    sChatLine = LastChatLine()
    If Right$(sChatLine, 21) = "has entered the room." Then
        sRoom% = FindChatRoom()
        iDiff = Len(sChatLine) - 22
        sTheSN = Mid$(sChatLine, 1, iDiff)
        If sTheSN = LCase("CAT") Then
            Call SendMessage(sRoom%, WM_CLOSE, 0, 0)
        End If
        GoTo sstart
    End If
Else
    GoTo sstart
End If
End Sub
Sub Availible(Person)
'This Will Check If Someone is Available

Call KeyWord("aol://9293:")
TimeOut 1.7
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
IMSendTo% = FindChildByClass(IM%, "_AOL_Edit")
Call SendMessageByString(IMSendTo%, WM_SETTEXT, 0, Person)
e = FindChildByClass(IM%, "RICHCNTL")
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
ClickIcon (e)
TimeOut 1
OkWin% = FindWindow("#32770", "America Online")
E2 = FindChildByClass(OkWin%, "Static")
E2 = GetWindow(E2, GW_HWNDNEXT)
OkWinMsgMsg$ = GetText(E2)
If OkWinMsgMsg$ = Person & " is online and able to receive Instant Messages." Then
    Msg$ = " Can Be Punted"
    AvailibleYes = True
    GoTo Ending
ElseIf OkWinMsgMsg$ = Person & " is not currently signed on." Then
    Msg$ = " Ain't Online!"
    AvailibleYes = False
Else
    Msg$ = " Has IMs Off"
    AvailibleYes = False
End If
Ending:
    If OkWin% <> 0 Then
        Call SendMessage(OkWin%, WM_CLOSE, 0, 0)
        Call SendMessage(IM%, WM_CLOSE, 0, 0)
    End If
    Call MsgBox(Person & Msg$)
End Sub

Function XNoFadeRedGreen(TheText)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next w
XNoFadeRedGreen = P$
End Function
Function XNoFadeRedBlue(TheText)
G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
XNoFadeRedBlue = P$
End Function

Function TrimTime()
B$ = Left$(Time$, 5)
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime = HourH$ & Right$(B$, 3) & " " & Ap$
End Function
Function TrimTime2()
B$ = Time$
HourH$ = Left$(B$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = Str$(HourA)
TrimTime2 = HourH$ & ":" & Right$(B$, 5) & " " & Ap$
End Function

Function EliteText(word$)
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
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q

EliteText = Made$

End Function

'Sub MyName()
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::               :::       ::::::::::: ")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>      :::    :::::::    :::           :::")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B>:::   :::   :::   :::   :::           :::")
'Call timeout(0.15)
'SendChat ("<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">" & "<B> :::::::     :::::::    :::::::::     :::")
'End Sub

Sub IMIgnore(thelist As ListBox)
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = IM%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Sub XChatStrike(StrikeOutChat)
'This strikes out text in a chat
SendChat ("<S>" & StrikeOutChat & "</S>")
End Sub
Sub XBI(STUFF)
'All of these letter things are chat stuff B=Bold,I=Italic,S=Strike
SendChat ("<B><I>" & STUFF & "</b></I>")
End Sub
Sub XBIS(STUFF)
SendChat ("<B><I><S>" & STUFF & "</s></b></I>")
End Sub
Sub XBS(STUFF)
SendChat ("<B><S>" & STUFF & "</B></s>")
End Sub
Sub XIS(STUFF)
SendChat ("<I><S>" & STUFF & "</s></I>")
End Sub
Sub XI(STUFF)
SendChat ("<I>" & STUFF & "</I>")
End Sub
Function SNfromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function

Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
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

Function XWavY(TheText As String)

G$ = TheText
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<sup>" & r$ & "</sup>" & u$ & "<sub>" & S$ & "</sub>" & T$
Next w
XWavY = P$

End Function


Sub RespondIM(Message)
'This finds an IM sent to you, answers it with a
'message of "message", sends it and then closes the
'IM window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
e = FindChildByClass(IM%, "RICHCNTL")

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
e = GetWindow(e, GW_HWNDNEXT)
E2 = GetWindow(e, GW_HWNDNEXT) 'Send Text
e = GetWindow(E2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(E2, WM_SETTEXT, 0, Message)
ClickIcon (e)
Call TimeOut(0.8)
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
e = FindChildByClass(IM%, "RICHCNTL")
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
e = GetWindow(e, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (e)
End Sub

Function MessageFromIM()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

IM% = FindChildByTitle(MDI%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(MDI%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(imtext%)
sn = SNfromIM()
snlen = Len(SNfromIM()) + 3
Blah = Mid(IMmessage, InStr(IMmessagge, sn) + snlen)
MessageFromIM = Left(Blah, Len(Blah) - 1)
End Function

Sub RunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub
Function RoomBuster(Room As TextBox, counter As Label)
'duh
d = FindChatRoom4
If d Then KillWin (d)

Do: DoEvents
Call KeyWord("aol://2719:2-2-" + Room + "")
waitforok
counter = counter + 1
If FindChatRoom Then Exit Do
If Text2 = 1 Then Exit Do
Loop
End Function
Sub XScroll50Line(Text As TextBox)
'scrolls around 50 linez, i never counted
Line = "Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text"
line2 = "Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text"
Call SendChat("<PRE<" & Line & line2 & ">>")
End Sub
Sub XStartGhosting()
'Makes ya ghost!
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")
If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If
SetupButton% = FindChildByClass(buddy%, "_AOL_Icon")
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
ClickIcon (SetupButton%)
TimeOut 3
PPButton% = FindChildByClass(buddysetup%, "_AOL_Icon")
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
ClickIcon (PPButton%)
TimeOut 3
PPWin% = FindChildByTitle(MDI%, "Privacy Preferences")
BlockAll% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
ClickIcon (BlockAll%)
BlockAll2% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
ClickIcon (BlockAll2%)
TimeOut 0.5
SaveButton% = FindChildByClass(PPWin%, "_AOL_Icon")
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
ClickIcon (SaveButton%)
waitforok
Call SendMessage(buddysetup%, WM_CLOSE, 0, 0)
End Sub
Sub XStopGhosting()
'+~-> This Will Stop You From Ghosting
'+~-> Ex. Call Stop Ghosting
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")
If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If
SetupButton% = FindChildByClass(buddy%, "_AOL_Icon")
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
ClickIcon (SetupButton%)
TimeOut 3
buddysetup% = FindChildByTitle(MDI%, UserSN & "'s Buddy Lists")
PPButton% = FindChildByClass(buddysetup%, "_AOL_Icon")
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
ClickIcon (PPButton%)
TimeOut 3
PPWin% = FindChildByTitle(MDI%, "Privacy Preferences")
BlockAll% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
ClickIcon (BlockAll%)
BlockAll2% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
ClickIcon (BlockAll2%)
TimeOut 0.5
SaveButton% = FindChildByClass(PPWin%, "_AOL_Icon")
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
ClickIcon (SaveButton%)
waitforok
Call SendMessage(buddysetup%, WM_CLOSE, 0, 0)
End Sub
Function XTextbackwards(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextchr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextchr$ & newsent$
Loop
XTextbackwards = newsent$
End Function

Sub RunMenuByString(Application, StringSearch)
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

'Sub Surge()
'G$ = "<FONT FACE=" & Chr$(34) & "ARIAL" & Chr$(34) & ">"
'SendChat (G$ & "<B> ::::::::                                                          :::")
'Call timeout(0.15)
'SendChat (G$ & "<B> :::::::   :::  :::   : :::::    ::::::     ::::::                            " & Chr$(160) & " " & "    :::  :::  :::   :::  :::  :::  :::   :::···´")
'Call timeout(0.15)
'SendChat (G$ & "<B>::::::::    ::::: ::  :::        :::::::    ::::::                             " & Chr$(160) & " " & "                                   :::")
'Call timeout(0.15)
'SendChat (G$ & "<B>                                ::::::::")
'Call timeout(0.5)
'End Sub

Sub Upchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub
Sub UnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub

Sub HideAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 0)
End Sub

Sub ShowAOL()
AOL% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL%, 5)
End Sub

Function XBoldBlackBlueBlack(Text1)
    a = Len(Text1)
    For B = 1 To a
        c = Left(Text1, B)
        d = Right(c, 1)
        e = 510 / a
        F = e * B
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 0)
        H = RGBtoHEX(G)
        Msg = Msg & "<Font Color=#" & H & ">" & d
    Next B
 XBoldBlackBlueBlack = Msg
XBoldSendChat (Msg)
End Function
Sub XBoldSendChat(BoldText)
'It will come out bold on the chat screen.
SendChat ("<b>" & BoldText & "</b>")
End Sub
Sub XAIMim(who As String, Message As String)
'send an im thru aim
aim% = FindWindow("_Oscar_BuddyListWin", vbNullString)
Blah% = FindChildByClass(aim%, "_Oscar_TabGroup")
heh% = FindChildByClass(Blah%, "_Oscar_IconBtn")
lol% = GetWindow(heh%, 2)
ClickIcon (lol%)
aim2% = FindWindow("AIM_IMessage", vbNullString)
bag% = FindChildByClass(aim2%, "_Oscar_PersistantCombo")
ed% = FindChildByClass(bag%, "Edit")
X% = SendMessageByString(ed%, WM_SETTEXT, 0, who$)
TimeOut 0.5
aim% = FindWindow("AIM_IMessage", vbNullString)
ack% = FindChildByClass(aim%, "WndAte32Class")
ack% = GetWindow(ack%, 2)
but% = FindChildByClass(aim%, "_Oscar_IconBtn")
X% = SendMessageByString(ack%, WM_SETTEXT, 0, Message$)
TimeOut 0.3

ClickIcon (but%)
End Sub
