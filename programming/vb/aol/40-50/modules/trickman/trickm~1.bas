Attribute VB_Name = "TrickmaNN"
'TríçkmáNN.BAS
'Coded By TríçkmáNN and dedicated to my grewp
'The Godfathers of Warez ( GoW )

'Disclaimer
'I nor anyone else but the user/users of this
'BAS can be held for the actions taken.
'By downloadin this BAS the user/users takes full
'responsibility for the actions made.


'If you would like to contribute to this BAS,
'Mail me at "TrickmaNN_579@Juno.com"
'I would be greatful for the donation and add
'you to the BAS creditz.

'Shoutz
'Id like to say Su[]D to TripeD, PhrEE, SoY, WinnieDM,
'DevilGA, Cholo, Twitch, Dominique Moceanu, Arizona,
'Sun Devils!, Leftover79, Gee Q, MaSe, My grewp
'The Godfathers of Warez and God.

'Creditz
'Programmed and coded by: TríçkmáNN a.k.a. Lucky Luciano


' That Cool Sound Stuff
Declare Function sndPlaySoundA Lib "c:\WINDOWS\SYSTEM\WINMM.DLL" (ByVal lpszSoundName$, ByVal ValueFlags As Long) As Long

   Global Const SND_SYNC = &H0
   Global Const SND_ASYNC = &H1
   Global Const SND_NODEFAULT = &H2
   Global Const SND_LOOP = &H8
   Global Const SND_NOSTOP = &H10
' End of the Cool Sound Stuff
'All that Fun Declaration stuff

'Global Const CB_GETCOUNT = (WM_USER + 6)
Declare Sub releaseCapture Lib "User" ()
Declare Function getnextwindow1 Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function getwindowtext Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer

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
   x As Long
   y As Long
End Type

Public Function GetChildCount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim i As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, GW_CHILD)
   

While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function


Sub Ao4Click (Button%)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
End Sub

Sub Ao4InstantMessage (Person, message)

AA% = FindWindow("Aol Frame25", 0&)
A% = FindChildByTitle(AA%, "Buddy List Window")
b% = FindChildByClass(A%, "_Aol_Icon")
If A% = 0 Then Ao4KW "Buddy View"
Do
A% = FindChildByTitle(AA%, "Buddy List Window")
b% = FindChildByClass(A%, "_Aol_Icon")
Call Pause(.001)
Loop Until A% <> 0
C% = GetWindow(b%, GW_HWNDNEXT)
D% = GetWindow(C%, GW_HWNDNEXT)
e% = GetWindow(D%, GW_HWNDNEXT)
Ao4Click D%
Do

'instant message part
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(IM%, "_AOL_Edit")
aolrich% = FindChildByClass(IM%, "RICHCNTL")
imsend% = FindChildByClass(IM%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call AOLSetText(aoledit%, Person)
Call AOLSetText(aolrich%, message)
imsend% = FindChildByClass(IM%, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, 2)
Next sends

AOLIcon (imsend%)

Do: DoEvents
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IM% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(IM%, WM_CLOSE, 0, 0): Exit Do
If IM% = 0 Then Exit Do
Loop
End Sub

Sub Ao4Kill45min ()
Do
A% = FindWindow("_Aol_Palette", 0&)
b% = FindChildByClass(A%, "_Aol_Icon")
Call Pause(.001)
Loop Until b% <> 0
Ao4Click (b%)
End Sub

Sub Ao4KW (where$)
b% = FindWindow("Aol Frame25", 0&)
A% = FindChildByClass(b%, "_Aol_Toolbar")
C% = FindChildByClass(A%, "_Aol_Icon")
D% = GetWindow(C%, GW_HWNDNEXT)
e% = GetWindow(D%, GW_HWNDNEXT)
F% = GetWindow(e%, GW_HWNDNEXT)
G% = GetWindow(F%, GW_HWNDNEXT)
H% = GetWindow(G%, GW_HWNDNEXT)
i% = GetWindow(H%, GW_HWNDNEXT)
J% = GetWindow(i%, GW_HWNDNEXT)
k% = GetWindow(J%, GW_HWNDNEXT)
L% = GetWindow(k%, GW_HWNDNEXT)
M% = GetWindow(L%, GW_HWNDNEXT)
N% = GetWindow(M%, GW_HWNDNEXT)
O% = GetWindow(N%, GW_HWNDNEXT)
P% = GetWindow(O%, GW_HWNDNEXT)
Q% = GetWindow(P%, GW_HWNDNEXT)
R% = GetWindow(Q%, GW_HWNDNEXT)
S% = GetWindow(R%, GW_HWNDNEXT)
T% = GetWindow(S%, GW_HWNDNEXT)
u% = GetWindow(T%, GW_HWNDNEXT)
V% = GetWindow(u%, GW_HWNDNEXT)
W% = GetWindow(u%, GW_HWNDNEXT)
y% = GetWindow(W%, GW_HWNDNEXT)
Ao4Click y%
Do
Aol = FindWindow("AOL Frame25", 0&)
bah = FindChildByTitle(Aol, "Keyword")
daedit = FindChildByClass(bah, "_AOL_Edit")
Pause (.001)
Loop Until daedit <> 0
daedit = FindChildByClass(bah, "_AOL_Edit")
Send daedit, where$
ico% = FindChildByClass(bah, "_AOL_Icon")
Ao4Click ico%
End Sub

Sub Ao4mailcenter ()
Aol% = FindWindow("AOL Frame25", 0&)
toolbar% = FindChildByClass(Aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
MCenter% = GetWindow(TooLBaRB%, 2)
Ao4Click MCenter%


End Sub

Sub Ao4MailSend (SendTo$, subject$, text$)
Aol% = FindWindow("AOL Frame25", 0&)
toolbar% = FindChildByClass(Aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
Ao4Click TooLBaRB%

Do
x% = DoEvents()
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
hideit = ShowWindow(chatlist%, SW_HIDE)
Loop Until chatedit% <> 0
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
hideit = ShowWindow(chatlist%, SW_HIDE)
chatwin% = GetParent(chatlist%)
Button% = FindChildByClass(chatlist%, "_AOL_Icon")
chatedit% = FindChildByClass(chatlist%, "_AOL_Edit")
sndtext% = SendMessageByString(chatedit%, WM_SETTEXT, 0, SendTo$)
blah% = GetWindow(chatedit%, GW_HWNDNEXT)
good% = GetWindow(blah%, GW_HWNDNEXT)
bad% = GetWindow(good%, GW_HWNDNEXT)
Sad% = GetWindow(bad%, GW_HWNDNEXT)
sndtext% = SendMessageByString(Sad%, WM_SETTEXT, 0, subject$)
rich = FindChildByClass(chatlist%, "RICHCNTL")
sndtext% = SendMessageByString(rich, WM_SETTEXT, 0, text$ & " ")
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
chatlist% = FindChildByTitle(FindWindow("AOL Frame25", 0&), "Compose Mail")
Button% = FindChildByClass(chatlist%, "_AOL_Icon")
SendNow% = SendMessageByNum(Button%, WM_LBUTTONDOWN, &HD, 0)
SendNow% = SendMessageByNum(Button%, WM_LBUTTONUP, &HD, 0)
Pause .2

End Sub

Sub Ao4openmail ()
Aol% = FindWindow("AOL Frame25", 0&)
toolbar% = FindChildByClass(Aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
Ao4Click TooLBaRB%


End Sub

Sub Ao4Sendtext (TextToSend$)

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
FIRS% = GetWindow(MDI%, 5)
LISTERS% = FindChildByClass(FIRS%, "RICHCNTL")
LISTERE% = getnextwindow1(LISTERS%, 2)
NXT% = getnextwindow1(LISTERE%, 2)
NXT1% = getnextwindow1(NXT%, 2)
NXT2% = getnextwindow1(NXT1%, 2)
NXT3% = getnextwindow1(NXT2%, 2)
NXT4% = getnextwindow1(NXT3%, 2)
LISTERB% = FindChildByClass(FIRS%, "_AOL_Listbox")
LISTER1% = FindChildByClass(FIRS%, "_AOL_Combobox")
DoEvents
DoEvents
sndtext% = SendMessageByString(NXT4%, WM_SETTEXT, 0, TextToSend$)
DoEvents
DoEvents
SendNow% = SendMessageByNum(NXT4%, WM_CHAR, &HD, 0)
DoEvents
End Sub

Sub Ao4Title (NewTitle$)
Aol% = FindWindow("AOL Frame25", 0&)
'textset Aol%, NewTitle$
End Sub

Sub Ao4writemail ()
Aol% = FindWindow("AOL Frame25", 0&)
toolbar% = FindChildByClass(Aol%, "AOL Toolbar")
ToolBarChild% = FindChildByClass(toolbar%, "_AOL_Toolbar")
TooLBaRB% = FindChildByClass(ToolBarChild%, "_AOL_Icon")
TooLBaRB% = GetWindow(TooLBaRB%, 2)
Ao4Click TooLBaRB%


End Sub

Sub AOLIcon (icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AOLSetText (win, Txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, Txt)
End Sub

Sub CenterForm (frm As Form)
frm.Left = Screen.Width / 2 - frm.Width / 2
frm.Top = Screen.Height / 2 - frm.Height / 2
End Sub

Function FindChildByClass (parentw, childhand)
FIRS% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(FIRS%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
FIRS% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(FIRS%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While FIRS%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
FIRS% = GetWindow(FIRS%, 2)
If UCase(Mid(GetClass(FIRS%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
room% = FIRS%
FindChildByClass = room%

End Function

Function FindChildByTitle (parentw, childhand)
FIRS% = GetWindow(parentw, 5)
If UCase(GetCaption(FIRS%)) Like UCase(childhand) Then GoTo bone
FIRS% = GetWindow(parentw, GW_CHILD)

While FIRS%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
FIRS% = GetWindow(FIRS%, 2)
If UCase(GetCaption(FIRS%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
room% = FIRS%
FindChildByTitle = room%
End Function

Function FreeProcess ()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
'frees process of freezes in your program
'and other stuff that makes your program
'slow down.  Works great.

End Function

Function GetCaption (hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
A% = getwindowtext(hWnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Function GetClass (child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function GetLineCount (text)

theview$ = text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(text, Len(text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Function GetWindowDir ()
buffer$ = String$(255, 0)
x = GetWindowsDirectory(buffer$, 255)
If Right$(buffer$, 1) <> "\" Then buffer$ = buffer$ + "\"
GetWindowDir = buffer$
End Function

Function IntegerToString (tochange As Integer) As String
IntegerToString = Str$(tochange)
End Function

Sub killwin (windo)
x = SendMessageByNum(windo, WM_CLOSE, 0, 0)
End Sub

Sub ParentChange (Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub

Sub Pause (interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub PlaySound (file)
'if you havent guessed this plays a wav file
'call playwave (app.path + "\yourwave.wav")

SoundName$ = file
   ValueFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndPlaySoundA(SoundName$, ValueFlags%)

End Sub

Function ReplaceText (text, charfind, charchange)
If InStr(text, charfind) = 0 Then
ReplaceText = text
Exit Function
End If

For Replace = 1 To Len(text)
thechar$ = Mid(text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function

Sub RunMenuByString (Application, StringSearch)
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

Sub Send (chatedit, sill$)
sndtext = SendMessageByString(chatedit, WM_SETTEXT, 0, sill$)
End Sub

Sub StayOnTop (frm As Form)
'  call stayontop(me)
SetWindowPos frm.hWnd, -1, 0, 0, 0, 0, &H50

End Sub

Function StringToInteger (tochange As String) As Integer
StringToInteger = tochange
End Function

Function TrimCharacter (thetext, chars)
'TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function TrimReturns (thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces (text)
If InStr(text, " ") = 0 Then
TrimSpaces = text
Exit Function
End If

For TrimSpace = 1 To Len(text)
thechar$ = Mid(text, TrimSpace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next TrimSpace

TrimSpaces = thechars$
End Function

Function UntilWindowClass (parentw, childhand)
GoBack:
DoEvents
FIRS% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(FIRS%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
FIRS% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(FIRS%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While FIRS%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
FIRS% = GetWindow(FIRS%, 2)
If UCase(Mid(GetClass(FIRS%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBack
FindClassLike = 0

bone:
room% = FIRS%
UntilWindowClass = room%
End Function

Function UntilWindowTitle (parentw, childhand)
GoBac:
DoEvents
FIRS% = GetWindow(parentw, 5)
If UCase(GetCaption(FIRS%)) Like UCase(childhand) Then GoTo bone
FIRS% = GetWindow(parentw, GW_CHILD)

While FIRS%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) Then GoTo bone
FIRS% = GetWindow(FIRS%, 2)
If UCase(GetCaption(FIRS%)) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBac
FindWindowLike = 0

bone:
room% = FIRS%
UntilWindowTitle = room%

End Function

Sub WaitWindow ()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
topmdi% = GetWindow(MDI%, 5)

Do: DoEvents
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
topmdi2% = GetWindow(MDI%, 5)
If Not topmdi2% = topmdi% Then Exit Do
Loop

End Sub

