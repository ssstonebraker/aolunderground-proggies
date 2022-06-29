Attribute VB_Name = "Tical2000"
'Hey to all the programmers this is DoDa
'This is My First Bas! - Tical 2000 it is for AOL 4.0
'I plan on making new additions to the bas in
'later months so look for updates at major VB sites
'Like KnK and Hider
'Any questions email me at TicalDoDa@hotmail.com
'Also i hate people saying my name wrong so here
'is how you say it
'"Doe" "Da"



Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function GetNextWindow Lib "user32" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long





Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long



Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)




Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpsSoundName As String, ByVal uFlags As Long) As Long


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

Sub AddRoomToListBox(Lst As ListBox)
Dim Index As Long
Dim i As Integer
For Index = 0 To 25
namez$ = String$(256, " ")
Ret = GetList(Index, namez$) & erb$
namez$ = Left$(Trim$(namez$), Len(Trim(namez$)))

ADDLB namez$, Lst
Next Index
end_addr:
Lst.RemoveItem Lst.ListCount - 1

i = GetListIndex(Lst, UserSN())
If i <> -2 Then Lst.RemoveItem i

End Sub
Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Function AOLOpen()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
If welcome% = 0 Then
AOLOpen = 0 'The 0 means that AOL is not open
Exit Function
End If
AOLOpen = 1 'The 1 means AOL is open
End Function

Sub AOLRunMenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer
AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
RunTheMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)
End Sub

Function AOLWindow()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function
Sub Bold()
'Click the Bold button in chatroom
'might want to put this in a check box
Room% = FindRoom()
bol% = FindChildByClass(Room%, "_AOL_Icon")
bol% = GetWindow(bol%, GW_HWNDNEXT)
Click% = SendMessage(bol%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(bol%, WM_LBUTTONUP, 0, 0&)
End Sub
Sub CenterForm(frm As Form)
'Centers a Form on the Screen
'Example -  In Form_Load put   Call CenterForm(Me)
A% = (Screen.Width - frm.Width) / 2
B% = (Screen.Height - frm.Height) / 2
frm.Move A%, B%
End Sub
Sub ChangeAolCaption(newcaption)
'This changes AOL's Caption
'Example - Call ChangeAolCaption(DoDa's AOL)
Call SetText(AOLWindow(), newcaption)
End Sub
Function Chat_Black()
Chat_Black = "<B><FONT COLOR=""#FFFFFF"">"
End Function
Function Chat_Blue()
Chat_Blue = "<B><FONT COLOR=""#0000FF"">"
End Function
Function Chat_Green()
Chat_Green = "<B><FONT COLOR=""#008000"">"
End Function

Function Chat_Red()
Chat_Red = "<B><FONT COLOR=""#FF0000"">"
End Function

Sub ChatSend(txt)
'Sends it to a chat room
'this goes through two windows so the chat window
'does no have to be on top!!!
    Room% = FindRoom()
    If Room% Then
        hChatEdit% = FindChildByClass(Room%, "RICHCNTL")
        Ret = SendMessageByString(hChatEdit%, WM_SETTEXT, 0, txt)
        Ret = SendMessageByNum(hChatEdit%, WM_CHAR, 13, 0)
    End If
End Sub
Sub ChatSend40(Text)
If FindRoom = 0 Then Exit Sub
AOLListz% = FindChildByClass(FindRoom, "_AOL_Listbox")
If AOLListz% = 0 Then Exit Sub
Room% = FindRoom()
R1% = FindChildByClass(Room%, "RICHCNTL")
R2% = GetWindow(R1%, 2)
R3% = GetWindow(R2%, 2)
R4% = GetWindow(R3%, 2)
R5% = GetWindow(R4%, 2)
R6% = GetWindow(R5%, 2)
R7% = GetWindow(R6%, 2)
DoEvents
sBuffer = Space(92)
sBuffer = GetText(R7%)

Call SetText(R7%, Text)
SendIcon% = GetWindow(R7%, 2)
ClickIcon (SendIcon%)
ClickIcon (SendIcon%)
Call SetText(R7%, sBuffer)
End Sub
Sub Clickbutton(Button%)
but% = SendMessage(Button%, WM_KEYDOWN, VK_SPACE, 0)
but% = SendMessage(Button%, WM_KEYUP, VK_SPACE, 0)
End Sub
Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function ClickList(hWnd, Index)
X = SendMessage(hWnd, LB_SETCURSEL, ByVal CLng(Index), ByVal 0&)
End Function

Function CountMail()
AoL40% = FindWindow("AOL Frame25", vbNullString)
MDI% = AOL4_MDI
TB% = FindChildByClass(AoL40%, "AOL Toolbar")
Toolz% = FindChildByClass(TB%, "_AOL_Toolbar")
AOLRead% = FindChildByClass(Toolz%, "_AOL_Icon")
ClickIcon (AOLRead%)
Pause 0.1
u = GetUser
Do:
DoEvents
MailPar% = FindChildByTitle(MDI%, u + "'s Online Mailbox")
TabControl% = FindChildByClass(MailPar%, "_AOL_TabControl")
TabPage% = FindChildByClass(TabControl%, "_AOL_TabPage")
Tree& = FindChildByClass(TabPage%, "_AOL_Tree")
If MailPar% <> 0 And TabControl% <> 0 And TabPage% <> 0 And Tree& <> 0 Then Exit Do
Loop
Pause 5
Do Until MailNum <> 0 And MailNum = MailNum2 And MailNum2 = MailNum3
    DoEvents
    MailNum = SendMessage(Tree&, LB_GETCOUNT, 0, 0)
    Pause (0.7)
    MailNum2 = SendMessage(Tree&, LB_GETCOUNT, 0, 0)
    Pause (1.5)
    MailNum3 = SendMessage(Tree&, LB_GETCOUNT, 0, 0)
    Pause (0.346)
    Loop

sBuffer = SendMessage(Tree&, LB_GETCOUNT, 0, 0)
If sBuffer > 1 Then
MsgBox "You have " & sBuffer & " New mails in your mail box!!"
End If
If sBuffer = 1 Then
MsgBox "You have one message in your Mailbox."
End If
If sBuffer < 1 Then
MsgBox "Common Sense tells you if the mail box doesn't have letters in it, YOU DON'T HAVE MAIL!!!"
End If

End Function
Function AOL4_MDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOL4_MDI = FindChildByClass(AOL%, "MDIClient")
End Function

Sub Decompile_Protect(Text1 As TextBox)
'This will Shield an exe so "code stealers" can't
'Decompile it
On Error Resume Next
If Text1 = "" Then Exit Sub
fee = FreeFile
Open Text1 For Binary As #fee
cat = "."
Seek #fee, 25
Put #fee, , cat
Close #1
If Err Then Exit Sub
MsgBox Text1 & " has been protected.", 64, "DoDa's AOL"
End Sub

Sub DoubleClick(hWnd As Integer)
o% = SendMessageByNum(hWnd, WM_LBUTTONDBLCLK, 0, 0)
End Sub
Sub ExtractSN(FilePath As String, Status As Label)
'This is for a PWS Detector..
'It will scan for an E-Mail Address
'Made By ZiP
'DON'T COPY IT!
For FindFileName = 1 To Len(FilePath)
TheName = Right(FilePath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName

free = FreeFile
Status.Caption = "Scanning: " & UCase(TheName)
DoEvents
Dim where As Long
Open FilePath For Binary Access Read As #free
filelength = LOF(free)
For X = 1 To filelength Step 32000
Text$ = Space(32000)
Get #free, X, Text$
If InStr(1, LCase(Text$), ".com", 1) = 0 Then GoTo NoAddress

For FindAddress = 1 To filelength
TheAddress = Mid(Text$, FindAddress, 4)

If LCase(TheAddress) = ".com" Then

For FindBegining = 1 To FindAddress
Character = LCase(Mid(Text$, FindAddress - FindBegining, 1))
If Character = "@" Or Character = "_" Then GoTo Skip
If Asc(Character) > 47 And Asc(Character) < 58 Then GoTo Skip
If Asc(Character) < 97 Or Asc(Character) > 122 Then AddressList = AddressList + Mid(Text$, FindAddress - FindBegining + 1, FindBegining + 3) & Chr(13): Exit For
Skip:
Next FindBegining

End If
Next FindAddress
Next X
If Len(AddressList) <> 0 Then Response = MsgBox(AddressList, , "Possible Addresses")
NoAddress:
If Len(AddressList) = 0 Then Response = MsgBox("sorry... no address could be extracted from: " & LCase(TheName) & Chr(13) & Chr(13) & "this may be a microsoft prouduct or a product that edits" & Chr(13) & "the .ini file in order to start at the time windows does...", , "Address Not Found")

Status.Caption = "Scanning Stopped..."
End Sub






Sub FadeFormBlue(vForm As Form)
'Fades the Background of your form Blue
'Put Call FadeFormBlue(Me) in the Paint section
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 0, 255 - intLoop), B
    Next intLoop
End Sub



Public Function FindChatRoom()
 AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    LISTERS% = FindChildByClass(firs%, "RICHCNTL")
    LISTERE% = FindChildByClass(firs%, "_AOL_Icon")
    LISTERB% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (LISTERS% = 0 Or LISTERE% = 0 Or LISTERB% = 0) And (l <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            LISTERS% = FindChildByClass(firs%, "RICHCNTL")
            LISTERE% = FindChildByClass(firs%, "_AOL_Icon")
            LISTERB% = FindChildByClass(firs%, "_AOL_Listbox")
            If LISTERS% And LISTERE% And LISTERB% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
        If InStr(1, GetCaption(firs%), "Welcome, ") Then FindRoom = 0: Exit Function
        FindChatRoom = firs%
        Exit Function
    End If
    
    FindChatRoom = 0
End Function


Function GetCaption(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function FindChildByClass(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Flapjacks
Wend
FindChildByClass = 0

Flapjacks:
Room% = firs%
FindChildByClass = Room%

End Function


Function FindChildByTitle(Parent, child As String) As Integer
C% = GetWindow(Parent, GW_CHILD)
C% = GetWindow(C%, GW_HWNDFIRST)
While C%
    DoEvents
    A% = SendMessage(C%, WM_GETTEXTLENGTH, 0, 0)
    B$ = String$(A%, 0)
    G% = SendMessageByString(C%, WM_GETTEXT, A% + 1, B$)
    If UCase$(B$) Like UCase$(child) & "*" Then
        FindChildByTitle = C%
        Exit Function
    End If
    C% = GetWindow(C%, GW_HWNDNEXT)
Wend
End Function

Function FindRoom()

    AOL% = FindWindow("AOL Frame25", vbNullString)
    MDI% = FindChildByClass(AOL%, "MDIClient")
    firs% = GetWindow(MDI%, 5)
    LISTERS% = FindChildByClass(firs%, "RICHCNTL")
    LISTERE% = FindChildByClass(firs%, "_AOL_Icon")
    LISTERB% = FindChildByClass(firs%, "_AOL_Listbox")
    Do While (LISTERS% = 0 Or LISTERE% = 0 Or LISTERB% = 0) And (l <> 100)
            DoEvents
            firs% = GetWindow(firs%, 2)
            LISTERS% = FindChildByClass(firs%, "RICHCNTL")
            LISTERE% = FindChildByClass(firs%, "_AOL_Icon")
            LISTERB% = FindChildByClass(firs%, "_AOL_Listbox")
            If LISTERS% And LISTERE% And LISTERB% Then Exit Do
            l = l + 1
    Loop
    If (l < 100) Then
        If InStr(1, GetCaption(firs%), "Welcome, ") Then FindRoom = 0: Exit Function
        FindRoom = firs%
        Exit Function
    End If
    
    FindRoom = 0
End Function


Function FindTextOnList(Item As String, list As ListBox)
'This finds text on a List box
Dim A As Long
Cnt = list.ListCount
For A = 0 To Cnt - 1
ToFind = InStr(list.list(A), Item)
If ToFind <= 0 Then FindTextOnList = False: Exit Function
If ToFind > 0 Then FindTextOnList = True: Exit Function
Next A
    
End Function
Sub FormMove(frm As Form)
'This moves a Form that you don't want borders on
'Put Call FormMove(Me) in the Mouse_Down

DoEvents
ReleaseCapture
ReturnVal% = SendMessage(frm.hWnd, &HA1, 2, 0)
End Sub

Function GetAoLDir()
'Gets AOL's Directory.
Buffer$ = Space(255)
f = GetProfileString("WAOL", "AppPath", "C:\AOL30", Buffer$, 255)
TrimSpaces (Buffer$)
TrimNull (Buffer$)
GetAoLDir = Buffer$
End Function
Function GetChat()
ChildS% = FindRoom
child% = FindChildByClass(ChildS%, "RICHCNTL")
Tex$ = GetText(child%)
For i = 1 To Len(Tex$)
Returns = InStr(i, Text, Chr(13), vbTextCompare)
Next i
GetChat = Tex$

End Function


Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim q As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, GW_CHILD)
   

While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
q = q + 1
Wend

GetChildCount = q
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Function GetClass(child)
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function

Sub GetCursor()
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", 0&)
X = SendMessage(FindWindow("_AOL_Modal", 0&), WM_CLOSE, 0, 0)
End Sub

Function GetLineCount(Text)

TheView$ = Text


For FindChar = 1 To Len(TheView$)
thechar$ = Mid(TheView$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function
Public Function GetList(Index As Long, Buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6
person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)
person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If
Buffer$ = person$
End Function



Public Function GetListIndex(oListBox As ListBox, sText As String) As Integer
Dim iIndex As Integer

With oListBox
 For iIndex = 0 To .ListCount - 1
   If .list(iIndex) = sText Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With

GetListIndex = -2
'  if Item isnt found
' I didnt want to use -1 as it evaluates to True

End Function




Function GetMessageFromIM()
'This Gets a Message from an IM..
'Example Text1 = GetMessageFromIM
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")

im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo JOe
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo JOe
Exit Function
JOe:
imtext% = FindChildByClass(im%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = GetSNFromIM()
snlen = Len(GetSNFromIM()) + 3
Blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
GetMessageFromIM = Left(Blah, Len(Blah) - 1)
CloseWindow im%
End Function


Function GetSNFromIM()
'Gets SN from an IM
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient") '

im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo JOe
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo JOe
Exit Function
JOe:
IMCap$ = GetCaption(im%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
GetSNFromIM = theSN$

End Function
Function GetWindowDir()
Buffer$ = String$(255, 0)
u = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Sub HideAOL()
If AOLIsOnline = 0 Then Exit Sub
DeLTa = FindWindow("AOL Frame25", vbNullString)
dkfhjl = ShowWindow(DeLTa, SW_HIDE)
End Sub
Sub HideWelcome()
'Hides annoy welcome window that won't close
AOL% = FindWindow("AOL Frame25", 0&)
welc% = FindChildByTitle(AOL%, "Welcome,")
X = ShowWindow(welc%, SW_HIDE)
X = SetFocusAPI(AOL%)
End Sub
Function IfFileExists(ByVal sFileName As String) As Boolean
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IfFileExists = False
    Else
        IfFileExists = True
    End If

End Function


Sub IMIgnore(thelist As ListBox)
'Ignores IMz from everyone on the list..
'If you dont want to get technical like this..
'Just make a TextBox & a command button
'In Command1 put..
'Call SendIM ("$IM_OFF "& Text1.text &", "IM Ignore")
'Make Text1 where they put the persons name

AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.list(findsn)) = LCase$(SNfromIM) Then
            BadIM% = im%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub


Sub IMRespond(message)
'This responds a message to an Incoming IM
'Great for IM Answering machines!
'Example In a Timer put
'Call IMRespond (" "& Text1.text &"")
'Text1=Thier message

im% = FindChildByTitle(AOL4_MDI(), ">Instant Message From:") 'Finds the title bar of your new IM
If im% Then GoTo OhBoy
im% = FindChildByTitle(AOL4_MDI(), "  Instant Message From:") 'Finds the title bar  where an existing IM is
If im% Then GoTo OhBoy
Exit Sub
OhBoy:
E% = FindChildByClass(im%, "RICHCNTL")

E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
E% = GetWindow(E%, 2)
e2% = GetWindow(E%, 2)
E% = GetWindow(e2%, 2)
Call SetText(e2%, message)
ClickIcon (E%)
Pause 0.2
CloseWindow (im%)
End Sub

Function IMsOff()
Call sendim("$IM_Off", "Gòt Thé Lìƒe ¹·º By DoDa")
Pause (0.5)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
closer = SendMessage(aolcl%, WM_CLOSE, 0, 0)
closer2 = SendMessage(im%, WM_CLOSE, 0, 0)
Loop Until aolcl% = 0 And im% = 0
End Function
Function IMsOn()
Call sendim("$IM_On", "Gòt Thé Lìƒe ¹·º By DoDa")
Pause (0.5)
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
closer = SendMessage(aolcl%, WM_CLOSE, 0, 0)
closer2 = SendMessage(im%, WM_CLOSE, 0, 0)
Loop Until aolcl% = 0 And im% = 0
End Function

Function INI_Read(AppName, KeyName As String, FileName As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    INI_Read = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Sub INI_Write(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)

Dim R As Integer
    R = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

Sub Italic()
'clicks the italic button in a chat room
'best if put in a check box
Room% = FindRoom()
bol% = FindChildByClass(Room%, "_AOL_Icon")
bol% = GetWindow(bol%, GW_HWNDNEXT)
bol% = GetWindow(bol%, GW_HWNDNEXT)
Click% = SendMessage(bol%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(bol%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub KeepMailAsNew()
'This keeps the Mail as new.
'Great for Mass Mailers
Mail% = FindChildByTitle(AOL4_MDI, "New Mail")
but% = FindChildByTitle(Mail%, "Keep As New")
Clickbutton but%
End Sub


Sub KeyWord(KeyWord)
'Does a keyword..
'Room Busters are great using this

AoL40% = AOLWindow
MDI40% = AOL4_MDI
Do:
DoEvents
Tool1% = FindChildByClass(AoL40%, "AOL Toolbar")
tool2% = FindChildByClass(Tool1%, "_AOL_Toolbar")
E% = FindChildByClass(tool2%, "_AOL_Glyph")
e1% = GetWindow(E%, 2)
e2% = GetWindow(e1%, 2)
e3% = GetWindow(e2%, 2)
e4% = GetWindow(e3%, 2)
e5% = GetWindow(e4%, 2)
e6% = GetWindow(e5%, 2)
If Tool1% <> 0 And tool2% <> 0 And E% <> 0 And e1% <> 0 And e2% <> 0 And e3% <> 0 And e4% <> 0 And e5% <> 0 And e6% <> 0 Then Exit Do
Loop
ClickIcon (e6%)

Do:
DoEvents
Chill% = FindChildByTitle(MDI40%, "Keyword")
Edit% = FindChildByClass(Chill%, "_AOL_Edit")
ICcD% = FindChildByClass(Chill%, "_AOL_Icon")
If Chill% <> 0 And Edit% <> 0 And ICcD% <> 0 Then Exit Do
Loop
Pause 0.1
Call SetText(Edit%, KeyWord)
Do: DoEvents
    Chill% = FindChildByTitle(MDI40%, "Keyword")
    Call ClickIcon(ICcD%)
Loop Until Chill% = 0
End Sub


Sub Keyword40(KeyWord)
'Same as above
AoL40% = AOLWindow
MDI40% = AOL4_MDI
Do:
DoEvents
Tool1% = FindChildByClass(AoL40%, "AOL Toolbar")
tool2% = FindChildByClass(Tool1%, "_AOL_Toolbar")
E% = FindChildByClass(tool2%, "_AOL_Glyph")
e1% = GetWindow(E%, 2)
e2% = GetWindow(e1%, 2)
e3% = GetWindow(e2%, 2)
e4% = GetWindow(e3%, 2)
e5% = GetWindow(e4%, 2)
e6% = GetWindow(e5%, 2)
If Tool1% <> 0 And tool2% <> 0 And E% <> 0 And e1% <> 0 And e2% <> 0 And e3% <> 0 And e4% <> 0 And e5% <> 0 And e6% <> 0 Then Exit Do
Loop
ClickIcon (e6%)

Do:
DoEvents
Chill% = FindChildByTitle(MDI40%, "Keyword")
Edit% = FindChildByClass(Chill%, "_AOL_Edit")
ICcD% = FindChildByClass(Chill%, "_AOL_Icon")
If Chill% <> 0 And Edit% <> 0 And ICcD% <> 0 Then Exit Do
Loop
Pause 0.1
Call SetText(Edit%, KeyWord)
Do: DoEvents
    Chill% = FindChildByTitle(MDI40%, "Keyword")
    Call ClickIcon(ICcD%)
Loop Until Chill% = 0
End Sub


Sub KillAdvertise()
'This will kill those annoying Advertisements
'In Mail, Upload window, Chat, Ect..
AOL% = FindWindow("AOL Frame25", 0&)
X% = FindChildByClass(AOL%, "_AOL_Image")
If X% = 0 Then Exit Sub
X% = FindChildByClass(AOL%, "_AOL_Image")
q% = ShowWindow(X%, SW_MINIMIZE)
q% = ShowWindow(X%, SW_HIDE)
End Sub

Sub KillAnnoyingGlyph()
'This kills that Aol Icon up in the right corner
'Word off caution IM's won't work right if used
AoL40% = AOLWindow
TB% = FindChildByClass(AoL40%, "AOL Toolbar")
Toolz% = FindChildByClass(TB%, "_AOL_Toolbar")
Glyph% = FindChildByClass(Toolz%, "_AOL_Glyph")
Call CloseWindow(Glyph%)
End Sub
Sub KillModals()
'This kills Modal windows
Dim the As String
Modal% = FindWindow("_AOL_Modal", vbNullString)
the = -1
Do:
If Modal% = 0 Then Exit Do
Modal% = FindWindow("_AOL_Modal", vbNullString)
CloseWindow (Modal%)
the = the + 1
Loop
If the < 1 Then
MsgBox "You have no Modal Windows open", vbExclamation
Exit Sub
End If
If the = 1 Then
MsgBox "1 Modal Window has been destroyed!", vbInformation
Exit Sub
End If
If the > 1 Then
MsgBox the + " Modal Windows have been destroyed!", vbInformation
Exit Sub
End If
End Sub

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

Sub KillWait1()
'Kill Hourglass
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")
For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 9)
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


Function LastChatLine()
'Gets last chat line of text...

getpar% = FindRoom()
child = FindChildByClass(getpar%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)

TheView$ = TrimSpace$


For FindChar = 1 To Len(TheView$)
thechar$ = Mid(TheView$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
LastLineo = Mid(TheView$, lastlen + 1, Len(thechars$) - 1)
LastChatLine = LastLineo
End Function

Function LinkSender(txt As String, URL As String)
Hyperlink = ("<A HREF=" & Chr$(34) & Text2 & Chr$(34) & ">" & Text1 & "</A>")
End Function


Public Sub ListDeleteItem(Lst As ListBox, Item$)
Do:
NoFreeze% = DoEvents()
If LCase$(Lst.list(A)) = LCase$(Item$) Then Lst.RemoveItem (A)
A = 1 + A
Loop Until A >= Lst.ListCount
End Sub

Sub LocateMember(theSN As String)
Call KeyWord("aol://3548:" & theSN)
End Sub
Sub MacroScroll(Text As String)
'This will scroll a Macro that you make
'Great for macro programs
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then
    Text$ = Text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(Text$, Chr$(13)) <> 0)
    Counter = Counter + 1
    Sendit Mid(Text$, 1, InStr(Text$, Chr(13)) - 1)
    If Counter = 4 Then
        TimeOut (2.9)
        Counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub

Sub Mail_Forward(SN, message)
'Forwards a mail
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDI Client")
FwdWin% = FindChildByTitle(AOL%, "Fwd: ")

Edit% = FindChildByClass(FwdWin%, "_AOL_Edit")
Rich% = FindChildByClass(FwdWin%, "RICHCNTL")
icon% = FindChildByClass(FwdWin%, "_AOL_ICON")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, SN)
Call SendMessageByString(Rich%, WM_SETTEXT, 0, message)
Call ClickIcon(icon%)
End Sub
Sub Maillist(Lst)
'List your Emails on a ListBox
Dim Z As String
AOL4_MDI
MailBox% = FindChildByTitle(AOL4_MDI(), "New Mail")
Tree% = FindChildByClass(MailBox%, "_AOL_TREE")
Z = 0
For i = 0 To SendMessageByNum(Tree%, LB_GETCOUNT, 0, 0&) - 1
MailStr$ = String$(255, " ")
    q% = SendMessageByString(Tree%, LB_GETTEXT, i, MailStr$)
    NoDate$ = Mid$(MailStr$, InStr(MailStr$, "/") + 4)
    NoSN$ = Mid$(NoDate$, InStr(NoDate$, Chr(9)) + 1)
    Lst.AddItem Z + ") " + Trim(NoSN$)
    Z = Z + 1
       Next i

End Sub



Sub MailSomething(person, subject, message)
'Fills out an E-Mail and Sends it
AoL40% = AOLWindow
TB% = FindChildByClass(AoL40%, "AOL Toolbar")
Toolz% = FindChildByClass(TB%, "_AOL_Toolbar")
AOLRead% = FindChildByClass(Toolz%, "_AOL_Icon")
AoLWrite% = GetWindow(AOLRead%, 2)
ClickIcon (AoLWrite%)
Pause (0.1)

Do:
DoEvents
MailWin% = FindChildByTitle(AOL4_MDI(), "Write Mail")
SendTo% = FindChildByClass(MailWin%, "_AOL_Edit")
StaCopyTo% = GetWindow(SendTo%, 2)
CopyTo% = GetWindow(StaCopyTo%, 2)
StaSubj% = GetWindow(CopyTo%, 2)
Subj% = GetWindow(StaSubj%, 2)
Msgg% = FindChildByClass(MailWin%, "RICHCNTL")
s1% = FindChildByClass(MailWin%, "_AOL_Icon")
s2% = GetWindow(s1%, 2)
s3% = GetWindow(s2%, 2)
s4% = GetWindow(s3%, 2)
s5% = GetWindow(s4%, 2)
s6% = GetWindow(s5%, 2)
s7% = GetWindow(s6%, 2)
s8% = GetWindow(s7%, 2)
s9% = GetWindow(s8%, 2)
s10% = GetWindow(s9%, 2)
s11% = GetWindow(s10%, 2)
s12% = GetWindow(s11%, 2)
s13% = GetWindow(s12%, 2)
s14% = GetWindow(s13%, 2)
s15% = GetWindow(s14%, 2)
s16% = GetWindow(s15%, 2)
s17% = GetWindow(s16%, 2)
s18% = GetWindow(s17%, 2)
SendNow% = GetWindow(s18%, 2)
If SendTo% <> 0 And CopyTo% <> 0 And Subj% <> 0 And Msgg% <> 0 And SendNow% <> 0 Then Exit Do
Loop
Call SetText(SendTo%, person)
Call SetText(Subj%, subject)
Call SetText(Msgg%, message)
Pause 0.1
ClickIcon (SendNow%)

Do:
DoEvents
Modd% = FindWindow("_AOL_Modal", vbNullString)
Icc% = FindChildByClass(Modd%, "_AOL_Icon")
If Icc% <> 0 Then Exit Do
Loop
Pause 0.1
ClickIcon (Icc%)
CloseWindow (MailWin%)
End Sub

Sub MassIM(Lst, txt)
'Mass IMer!
'Don't you just want to kick yourself
'For how easy it is
'Make 1 list and 1 Text Box..
'List = People on the Mass IM
'Text = What to send to them
Dim A As Long
For A = 0 To List1.ListCount - 1
Call sendim(Lst.list(A), " " & txt.txt & "")
Next A
End Sub
Sub MinUpchat()
'Minimizes Upload Window
AppActivate "America  Online"
aom = FindWindow("_AOL_MODAL", 0&)
If aom <> 0 Then
    X = ShowWindow(aom, SW_MINIMIZE)
End If
End Sub


'massmailing codes!!!!!!!!!
Function mm1()
u$ = UserSN
MDI& = AOL4_MDI
AOL% = FindChildByTitle("AOL Frames25", vbNullString)
gargantuA = FindChildByTitle(MDI&, u$ + "'s Online Mailbox")
TabControl& = FindChildByClass(gargantuA, "_AOL_TabControl")
TabPage& = FindChildByClass(TabControl&, "_AOL_TabPage")
Tree& = FindChildByClass(TabPage&, "_AOL_Tree")
Readbutton$ = FindChildByTitle(gargantuA, "Read")
ClickIcon (Readbutton$)
Call mm2
End Function
Function mm2()
TimeOut (4)
download% = FindChildByTitle(0, "Download Now")
fwdform% = GetParent(download%)
forwardbutton = FindChildByClass(fwdform%, "_AOL_Icon")
forwardbutton2 = FindChildByClass(forwardbutton, GW_HWNDNEXT)
ClickIcon (forwardbutton)
KillWindow Window%
Call KeepMailAsNew
Call ClickList(Tree&)
Call Enter(Tree&)
TimeOut (1)
End Function
Function mm3()
u$ = UserSN
MDI& = AOL4_MDI
gargantuA = FindChildByTitle(MDI&, u$ + "'s Online Mailbox")
TabControl& = FindChildByClass(gargantuA, "_AOL_TabControl")
TabPage& = FindChildByClass(TabControl&, "_AOL_TabPage")
Tree& = FindChildByClass(TabPage&, "_AOL_Tree")
Readbutton$ = FindChildByTitle(gargantuA, "Read")
ClickIcon (Readbutton$)
TimeOut (5)
Window% = FindChildByClass(MDI&, "_AOL_ICON")
AOL% = AOLWindow
download% = FindChildByTitle(AOL%, "Download Now")
fwdform% = GetParent(download%)
fwdform2% = FindChildByClass(GW_HWNDNEXT, fwdform%)
fwdform3% = FindChildByClass(GW_HWNDNEXT, fwdform2%)
fwdform4% = FindChildByClass(GW_HWNDNEXT, fwdform3%)
fwdform5% = FindChildByClass(GW_HWNDNEXT, fwdform4%)
fwdform6% = FindChildByClass(GW_HWNDNEXT, fwdform5%)
ClickIcon (fwdform6%)
KillWindow Window%
Call KeepMailAsNew
TimeOut (1)
End Function

Public Function NotOnList(thing As String, list As Control)

For i = 0 To list.ListCount
If LCase(TrimNull(TrimSpaces(thing$))) = LCase(TrimNull(TrimSpaces(list.list(i)))) Then NotOnList = False: Exit Function
Next i
NotOnList = True

End Function
Sub NotOnTop(the As Form)
SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Function OnList(thing As String, list As Control)

For i = 0 To list.ListCount
If LCase(TrimNull(TrimSpaces(thing$))) = LCase(TrimNull(TrimSpaces(list.list(i)))) Then OnList = True: Exit Function
Next i

OnList = False

End Function

Sub Pause(interval)
'Same as Timeout
'Some people use Pause and some use Timeout.
'So put best of both worlds
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub
Sub Playwav(File As String)
'Plays a Wav with no Freeze!!!
'Yeah Baby
i% = sndPlaySound(File, 1)
End Sub
Function RandomNumber(finished)
'Chooses a Random Number.
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function
Sub ReadNewMail()
'Opens your Mail Box
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
Toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(Toolbar%, "_AOL_Icon")
Call ClickIcon(icon%)
End Sub
Function ReverseText(Text)
For Werds = Len(Text) To 1 Step -1
    ReverseText = ReverseText & Mid(Text, Werds, 1)
Next Werds
End Function

Sub RoomBust()
'This is just an example..
'Make 1Text Box, 1 Timer, & 1 Label
'Label.Caption = 0
'Then do this in a timer

'Do
'Call Keyword40("aol://2719:2-2-" & text1.Text & " ")
'WaitForOK
'Label1.Caption = Label1.Caption + 1
'If FindChatRoom Then Exit Do
'Loop Until Timer1.Enabled = False

'In the Stop button put..
'Sendtext("Name of Prog")
'Sendtext (" Busted Into " & text1.Text & " In " & Label1.Caption & " Try(s)")
End Sub

Function RoomCount()
thechild% = FindRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox") 'Finds the AoL Lisbox in a chat room

GetCount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
RoomCount = GetCount
End Function
Sub RunMenu(menu1 As Integer, menu2 As Integer)

'Menu1 is Vertical and starts at 0
'Menu2 is Horizontal and starts at 0
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub
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

Function Scanfile(FilePath As String, Status As Label, SingleScan As Boolean) As Long
'This will scan for a PassWord Stealer
'Written by ZiP
'DONT COPY!

For FindFileName = 1 To Len(FilePath)
TheName = Right(FilePath, FindFileName)
If Left(TheName, 1) = "\" Then TheName = Right(TheName, FindFileName - 1): Exit For
Next FindFileName

free = FreeFile
Status.Caption = "Scanning: " & UCase(TheName)
DoEvents
Dim where As Long
Open FilePath For Binary Access Read As #free
filelength = LOF(free)
For X = 1 To filelength Step 32000

Text$ = Space(32000)
Get #free, X, Text$

If InStr(1, LCase(Text$), "win.ini", 1) Then PWSFound = True: Close #free: Exit For
PWSFound = False
Next X
Status.Caption = "Scanning Stopped..."
DoEvents
Close #free

If PWSFound = True Then
Response = MsgBox(LCase(TheName) & Chr(13) & Chr(13) & "Possible PWS Detected!!!" & Chr(13) & Chr(13) & "do you want to try to extract an e-mial address?", vbYesNo, "Possible PWS")
If Response = vbYes Then
Call ExtractSN(FilePath, Status)
End If
End If
If SingleScan = True Then If PWSFound = False Then Response = MsgBox(LCase(TheName) & Chr(13) & Chr(13) & "Is NOT A PWS")
End Function







Sub SendBuddyInvitation(person, message, Room)
'Sends buddy Invitations
MDI% = AOL4_MDI
o% = FindChildByTitle(MDI%, "Buddy List Window")
If o% <> 0 Then GoTo JOe
Keyword40 ("bv")
Do:
DoEvents
JOe:
BuddyBox% = FindChildByTitle(MDI%, "Buddy List Window")
Ico1% = FindChildByClass(BuddyBox%, "_AOL_Icon")
Ico2% = GetWindow(Ico1%, 2)
Ico3% = GetWindow(Ico2%, 2)
Ico4% = GetWindow(Ico3%, 2)
Ico5% = GetWindow(Ico4%, 2)
Ico6% = GetWindow(Ico5%, 2)
Ico7% = GetWindow(Ico6%, 2)
Pause 0.1
ClickIcon (Ico7%)
If BuddyBox% <> 0 And Ico7% <> 0 Then Exit Do
Loop
Do:
DoEvents
BuddyChat% = FindChildByTitle(MDI%, "Buddy Chat")
WhoToInvite% = FindChildByClass(BuddyChat%, "_AOL_Edit")
MsgToSend% = GetWindow(WhoToInvite%, 2)
w1% = GetWindow(MsgToSend%, 2)
w2% = GetWindow(w1%, 2)
w3% = GetWindow(w2%, 2)
w4% = GetWindow(w3%, 2)
Roomer% = GetWindow(w4%, 2)

If BuddyChat% <> 0 And WhoToInvite% <> 0 And MsgToSend% <> 0 And Roomer% <> 0 Then Exit Do
Loop

Call SetText(WhoToInvite%, person)
Call SetText(MsgToSend%, message)
Call SetText(Roomer%, Room)
Pause 0.1
SendBut% = FindChildByClass(BuddyChat%, "_AOL_Icon")
ClickIcon (SendBut%)
'Closewindow(BuddyChat%)

End Sub


Sub sendim(person, message)
'Sends an IM to a person
AppActivate "America  Online"
SendKeys ("^i")
Do: DoEvents
    MDI% = AOL4_MDI
    IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
    who% = FindChildByClass(IMWin%, "_AOL_Edit")
    mess% = FindChildByClass(IMWin%, "RICHCNTL")
    Sender% = GetWindow(mess%, 2)
    KWin% = FindChildByTitle(AOL4_MDI, "Keyword")
    X = SendMessage(KWin%, WM_CLOSE, 0, 0)
 If IMWin% <> 0 And who% <> 0 And mess% <> 0 And Sender% <> 0 And KWin% = 0 Then Exit Do
Loop
Pause 0.1
DoEvents
For i = 1 To 3
Call SetText(who%, person)
Next i
Pause 0.1
For i = 1 To 3
Call SetText(mess%, message)
Next i
ClickIcon (Sender%)
End Sub

Sub Sendit(Text)
'This is a simple way to send text to the chat room without
If FindRoom = 0 Then Exit Sub
AOLListz% = FindChildByClass(FindRoom, "_AOL_Listbox")
If AOLListz% = 0 Then Exit Sub
Room% = FindRoom()
R1% = FindChildByClass(Room%, "RICHCNTL")
R2% = GetWindow(R1%, 2)
R3% = GetWindow(R2%, 2)
R4% = GetWindow(R3%, 2)
R5% = GetWindow(R4%, 2)
R6% = GetWindow(R5%, 2)
R7% = GetWindow(R6%, 2)
DoEvents
sBuffer = Space(92)
sBuffer = GetText(R7%)

Call SetText(R7%, Text)
SendIcon% = GetWindow(R7%, 2)
ClickIcon (SendIcon%)
ClickIcon (SendIcon%)
Call SetText(R7%, sBuffer)
End Sub

Function sendlink(txt As TextBox, txt1 As TextBox)
sendlink = "<a href=" & Chr(34) & Chr(34) & "><a href=" & Chr(34) & Chr(34) & "><a href=" & Chr(34) & txt & "><font color=#0000ff><u>" & txt1 & "<font color=#fffeff></a>"
End Function




Sub Sendtext(messg)
'THIS SENDS THE TEXT TO THE CHAT ROOM WITH THE
'Written by SpaM
'DONT Copy
  C1BAK = C1
   C2BAK = C2
   C3BAK = C3
   C4BAK = C4
   C = 0
   o = 0
   o2 = 0
   q = 1
   Q2 = 1

   For X = 1 To Len(messg)
            BVAL1 = 0 - 256
            BVAL2 = 0 - 0
            BVAL3 = 0 - 0
            VAL1 = (BVAL1 / Len(messg) * X) + 1
            VAL2 = (BVAL2 / Len(messg) * X) + 6
            VAL3 = (BVAL3 / Len(messg) * X) + 15
            C1 = RGB2HEX(VAL1, VAL2, VAL3)
            C2 = RGB2HEX(VAL1, VAL2, VAL3)
            C3 = RGB2HEX(VAL1, VAL2, VAL3)
            C4 = RGB2HEX(VAL1, VAL2, VAL3)
            If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then C = 1: msg = msg & "<FONT COLOR=#" + C1 + ">"
            If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
            If C <> 1 Then
            If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
            End If
            If o2 = 1 Then msg = msg + "<sub>"
            If o2 = 3 Then msg = msg + "<sup>"
          msg = msg + Mid$(messg, X, 1)
          If o2 = 1 Then msg = msg + "</sub>"
          If o2 = 3 Then msg = msg + "</sup>"
          If Q2 = 2 Then
            q = 1
            Q2 = 1
      If o2 = 1 Then msg = msg + "<FONT COLOR=#" + C1 + ">"
      If o2 = 2 Then msg = msg + "<FONT COLOR=#" + C2 + ">"
      If o2 = 3 Then msg = msg + "<FONT COLOR=#" + C3 + ">"
      If o2 = 4 Then msg = msg + "<FONT COLOR=#" + C4 + ">"
   End If
nc: Next X
C1 = C1BAK
C2 = C2BAK
C3 = C3BAK
C4 = C4BAK
ChatSend40 "<b>" & msg
End Sub

Function RGB2HEX(R, G, B)

Dim X&
Dim xx&
Dim Color&
Dim Divide
Dim Answer&
Dim Remainder&
Dim Configuring$
For X& = 1 To 3
    If X& = 1 Then Color& = B
    If X& = 2 Then Color& = G
    If X& = 3 Then Color& = R
    For xx& = 1 To 2
        Divide = Color& / 16 'Red, Green, or Blue Divided by 16
        Answer& = Int(Divide) 'Turn Divide into an Integer
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
Configuring$ = TrimSpaces(Configuring$)
RGB2HEX = Configuring$
End Function


Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function
Sub SetText(Window, Text)
hWnd1% = FindWindow("ThunderRT5Form")
hWnd2% = FindChildByClass(hWnd1%, "ThunderRT5TextBox")

eltext% = SendMessageByString(Window, WM_SETTEXT, 0, "")
eltext% = SendMessageByString(Window, WM_SETTEXT, 0, Text)
End Sub


Sub ShowAOL()
B = FindChildByClass(AOLWindow(), "AOL TOOLBAR")
p = ShowWindow(AOLWindow(), SW_SHOW)
p = ShowWindow(B, SW_SHOW)
End Sub
Sub ShowAoLToolBar()
DeLTa = FindWindow("AOL Frame25", vbNullString)
SocK = FindChildByClass(DeLTa, "AOL Toolbar")
PLoP = ShowWindow(SocK, SW_SHOW)
End Sub

Sub signoff()

Call SendMessage(AOLWindow, WM_CLOSE, 0&, 0&)
End Sub

Function SpiralText(sBuffer)
K = sBuffer
DeLTa:
Dim MYLEN As Integer
MYSTRING = sBuffer
MYLEN = Len(MYSTRING)
mystr = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
sBuffer = mystr
Pause 1
If sBuffer = K Then
SpiralText = K: Exit Function
End If
GoTo DeLTa
End Function

Public Function Status(SN As String) As Integer
ToolKeyword40 ("aol://9293:" & SN$)
Do: DoEvents
    MDI% = AOL4_MDI
    IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
    who% = FindChildByClass(IMWin%, "_AOL_Edit")
    mess% = FindChildByClass(IMWin%, "RICHCNTL")
    Sender% = GetWindow(mess%, 2)
    KWin% = FindChildByTitle(AOL4_MDI, "Keyword")
    X = SendMessage(KWin%, WM_CLOSE, 0, 0)
 If IMWin% <> 0 And who% <> 0 And mess% <> 0 And Sender% <> 0 And KWin% = 0 Then Exit Do
Loop
Pause 0.1
DoEvents
stat% = GetWindow(Sender%, 2)
ClickIcon (stat%)
While stat_win% = 0
stat_win% = FindWindow("#32770", "America Online")
Wend
stat_text% = FindChildByClass(stat_win%, "Static")
stat_info$ = GetCaption(stat_text%)
Buffer$ = Space$(255)
stat_info$ = GetWindowText(stat_text%, Buffer$, -1)
Debug.Print stat_info$
stat_ok% = FindChildByTitle(stat_win%, "Ok")
Call ClickIcon(stat_ok%)
Call CloseWindow(IMWin%)
If InStr(stat_info$, "off") Then
Status% = stat_imoff
ElseIf InStr(stat_info$, "on") Then
Status% = stat_imon
Else
Status% = stat_offline
End If

End Function

Sub StayOnTop(frm As Form)
'In Form_Load put
'StayOnTop Me
'It doesnt work until the program is In an exe
Dim ontop%
ontop% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub StopWAV()
'It don't play a wav so don't put one in there
Playwav (" ")
End Sub
Function StringToInteger(OneToChange As String) As Integer
StringToInteger = OneToChange
End Function


Sub TimeOut(Duration)
'Same as Pause
'Some people use Pause and some use Timeout.
'So i put both in here
starttime = Timer
Do While Timer - starttime < Duration
X = DoEvents()
Loop
End Sub
Sub ToolKeyword40(Text)
AOL% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(AOL%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Do:
DoEvents
AOEdit% = FindChildByClass(AOTool2%, "_AOL_ComboBox")
AOedit2% = FindChildByClass(AOEdit%, "Edit")
If AOTool2% <> 0 And AOEdit% <> 0 And AOedit2% <> 0 Then Exit Do
Loop
Call SendMessageByString(AOedit2%, WM_SETTEXT, 0, Text)
Call SendMessageByNum(AOedit2%, WM_CHAR, VK_SPACE, 0)
Call SendMessageByNum(AOedit2%, WM_CHAR, VK_RETURN, 0)
End Sub
Function TrimCharacter(TheText, chars)
TrimCharacter = ReplaceText(TheText, chars, "")
End Function

Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For Replace = 1 To Len(Text)
thechar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function


Function TrimNull(What As String)

wstr = Trim(What)
Do Until xx = Len(What)
Let xx = xx + 1
Let this_chr = Asc(Mid$(What, xx, 1))
If this_chr > 31 And this_chr <> 256 Then Let wordd = wordd & Mid$(What, xx, 1)
Loop
TrimNull = wordd
End Function


Function TrimReturns(Text)
takechr13 = ReplaceText(Text, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function


Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For TrimSpace = 1 To Len(Text)
thechar$ = Mid(Text, TrimSpace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next TrimSpace

TrimSpaces = thechars$
End Function


Sub Underline()
'will click the underline button in a chat room
'best if put in a check box
Room% = FindRoom()
bol% = FindChildByClass(Room%, "_AOL_Icon")
bol% = GetWindow(bol%, GW_HWNDNEXT)
bol% = GetWindow(bol%, GW_HWNDNEXT)
bol% = GetWindow(bol%, GW_HWNDNEXT)
Click% = SendMessage(bol%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(bol%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub UpChatOff()
aom% = FindWindow("_AOL_Modal", 0&)
DoEvents
X = ShowWindow(aom%, SW_SHOW)
X = SetFocusAPI(aom%)

End Sub


Sub UpchatOn()
AOL% = FindWindow("AOL Frame25", 0&)
aom% = FindWindow("_AOL_Modal", 0&)
DoEvents
X = ShowWindow(aom%, SW_HIDE)
X = SetFocusAPI(AOL%)

End Sub

Function UserSN()
'Finds the User Screen Name
'Example Text1 = "My Screen Name is " & UserSN & ""
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
A% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
If Len(UserSN) = 0 Then UserSN = "User"
End Function

Sub waitforok()
'This will click the OK Button on AOL
'Example When a room is full it will push it for u
Do:
DoEvents
clikr% = FindWindow("#32770", "America Online")
okbtn% = FindChildByTitle(clikr%, "OK")
If clikr% <> 0 And okbtn% <> 0 Then Exit Do
Loop
Pause 0.1
Clickbutton (okbtn%)
n% = SendMessage(okbtn%, WM_KEYUP, VK_SPACE, 0)
End Sub

Function WelcomeMenu()
Welcomez% = FindChildByTitle(AOL4_MDI, "Welcome, ")
If Welcomez% = 0 Then
    WelcomeMenu = 0
Else
    WelcomeMenu = Welcomez%
End If
End Function

Function WhichAoL()
If AOLWindow = 0 Then Exit Function
ToolBar30% = FindChildByClass(AOLWindow, "AOL Toolbar")
ToolBar40% = FindChildByClass(ToolBar30%, "_AOL_Toolbar")

If ToolBar40% <> 0 Then
    WhichAoL = 4 'Means they are using Aol 4
Else
    WhichAoL = 3 'Means they are using Aol 3
End If
End Function

Sub Window_Close(win)
' This is like killwin
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub

Sub Window_Hide(hWnd)
X = ShowWindow(hWnd, SW_HIDE)
End Sub

Sub Window_Maximize(win)

X = ShowWindow(win, SW_MAXIMIZE)
End Sub

Sub Window_Minimize(win)

X = ShowWindow(win, SW_MINIMIZE)
End Sub

Sub Window_Show(hWnd)

X = ShowWindow(hWnd, SW_SHOW)
End Sub

Function WindowCaption(hWnd As Integer)

Dim WindowText As String * 255
Dim GetWinText As Integer
GetWinText = GetWindowText(hWnd, WindowText, 255)
WindowCaption = (WindowText)
End Function


Sub WriteToLog(What As String, FilePath As String)
'Writes text to a Txt
If FilePath = "" Then Exit Sub
f% = FreeFile
Open FilePath For Binary Access Write As f%
p$ = What & Chr(10)
Put #1, LOF(1) + 1, p$
Close f%
End Sub

Sub addroom(Lst As Control)
'Need a List Box
Room = FindRoom()
ListX% = FindChildByClass(Room, "_AOL_List")
Count% = SendMessage(ListX%, LB_GETCOUNT, 0&, 0&)
Buffer$ = Space(255)
For Counter% = 0 To Count% - 1
getit = AOLGetList(ListX%, Counter%, Buffer$)
For E = 0 To Lst.ListCount
    If Lst.list(E) = Buffer$ Then
    GoTo Here:
    End If
Next E
If Buffer$ = GetUser Then GoTo Here
Lst.AddItem (Buffer$)
Here:
Next Counter%
End Sub


Function AOLGetList(LBHandle, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
   
AOLThread = GetWindowThreadProcessId(LBHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(LBHandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal person$, 4)
ListPersonHold = ListPersonHold + 6

person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, person$, Len(person$), ReadBytes)

person$ = Left$(person$, InStr(person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = person$
AOLGetList = Buffer$
End Function


Function AddMailToList(Which As Integer, list As ListBox)
AoL40% = FindWindow("AOL Frame25", vbNullString)
MDI% = AOL4_MDI
TB% = FindChildByClass(AoL40%, "AOL Toolbar")
Toolz% = FindChildByClass(TB%, "_AOL_Toolbar")
AOLRead% = FindChildByClass(Toolz%, "_AOL_Icon")
ClickIcon (AOLRead%)
Pause 0.1
u = GetUser
Do:
DoEvents
MailPar% = FindChildByTitle(MDI%, u + "'s Online Mailbox")
TabControl% = FindChildByClass(MailPar%, "_AOL_TabControl")
TabPage% = FindChildByClass(TabControl%, "_AOL_TabPage")
If Which = 2 Then TabPage% = GetWindow(TabPage%, GW_HWNDNEXT)
If Which = 3 Then TabPage% = GetWindow(TabPage%, GW_HWNDNEXT): TabPage% = GetWindow(TabPage%, GW_HWNDNEXT)
Tree& = FindChildByClass(TabPage%, "_AOL_Tree")
If MailPar% <> 0 And TabControl% <> 0 And TabPage% <> 0 And Tree& <> 0 Then Exit Do
Loop

sBuffer = SendMessage(Tree&, LB_GETCOUNT, 0, 0)

For MailNum = 0 To sBuffer
TxtLen% = SendMessageByNum(Tree&, LB_GETTEXTLEN, MailNum, 0&)
txt$ = String(TxtLen% + 1, 0&)
GTTXT% = SendMessageByString(Tree&, LB_GETTEXT, MailNum, txt$)
NewMail = RTrim(txt$)
list.AddItem (NewMail)
Next MailNum

End Function

Sub ADDLB(itm As String, Lst As ListBox)
If Lst.ListCount = 0 Then
Lst.AddItem itm
Exit Sub
End If
Do Until xx = (Lst.ListCount)
Let diss_itm$ = Lst.list(xx)
If Trim(LCase(diss_itm$)) = Trim(LCase(itm)) Then Let do_it = "NO"
Let xx = xx + 1
Loop
If do_it = "NO" Then Exit Sub
Lst.AddItem itm
End Sub
Sub CloseWindow(Window As Integer)
win% = SendMessageByNum(Window, WM_CLOSE, 0, 0)
End Sub
