Attribute VB_Name = "Guess"
'This .bas was not made entirely by me
'but a lot of it was made by me
'All of the AIM stuff was made by me
'but please D/L my AbbotAIM.bas if you want a .bas from me
'you can get it at KNK4.o (http://knk.tierranet.com/knk4o/index2.html)
'This example was made by me (ThaAbbot)
'Dedicate to ProGs InC
'Please Visit The ProGs InC website!
'http://come.to/progsinc
'later

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
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
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
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
        X As Long
        Y As Long
End Type





Function AOL_Hide()
AOL& = FindWindow("AOL Frame25", vbNullString)
X = ShowWindow(AOL&, SW_HIDE)

End Function

Function AIMAttention(message$)
'Gets peoples attention
SendChat "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
SendChat message$
SendChat "`·.·¨'¹l|[Attention]|l¹¨'·.·´"
End Function
Sub SendChat2(Chat)
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
Function SetChildFocus(child)
setchild& = SetFocusAPI(child)
End Function

Function mail_scroll()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "'s Online Mailbox")
Control& = FindWindowEx(Box&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(Control&, 0&, "_AOL_TabPage", vbNullString)
MailList& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
For d = 0 To SendMessageByNum(MailList&, LB_GETCOUNT, 0, 0&) - 1
S = SendMessageByNum(MailList&, LB_SETCURSEL, d, 0&)
Timeout (0.6)
Next d
End Function


Function Text_Spaced(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Spaced = Newsent$
SendChat "" + Text_Spaced + ""
End Function
Function IM_spaced(who, strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let Newsent$ = Newsent$ + NextChr$
Loop
IM_spaced = Newsent$
Call SendIM(who, "" + IM_spaced + "")
End Function
Function Text_lag(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "<HTML></HTML><HTML></HTML><HTML></HTML>"
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_lag = Newsent$
SendChat "</B></I></U><font color=#000000>" + Text_lag + ""
End Function
Function IM_lag(who, strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "<HTML></HTML><HTML></HTML><HTML></HTML>"
Let Newsent$ = Newsent$ + NextChr$
Loop
IM_lag = Newsent$
Call SendIM(who, "</B></I></U><font color=#000000>" + IM_lag + "")
End Function
Function Text_Hacker(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Hacker = Newsent$
SendChat "" + Text_Hacker + ""
End Function
Function IM_Hacker(who, strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
If NextChr$ = "A" Then Let NextChr$ = "a"
If NextChr$ = "E" Then Let NextChr$ = "e"
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "O" Then Let NextChr$ = "o"
If NextChr$ = "U" Then Let NextChr$ = "u"
If NextChr$ = "b" Then Let NextChr$ = "B"
If NextChr$ = "c" Then Let NextChr$ = "C"
If NextChr$ = "d" Then Let NextChr$ = "D"
If NextChr$ = "z" Then Let NextChr$ = "Z"
If NextChr$ = "f" Then Let NextChr$ = "F"
If NextChr$ = "g" Then Let NextChr$ = "G"
If NextChr$ = "h" Then Let NextChr$ = "H"
If NextChr$ = "y" Then Let NextChr$ = "Y"
If NextChr$ = "j" Then Let NextChr$ = "J"
If NextChr$ = "k" Then Let NextChr$ = "K"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "m" Then Let NextChr$ = "M"
If NextChr$ = "n" Then Let NextChr$ = "N"
If NextChr$ = "x" Then Let NextChr$ = "X"
If NextChr$ = "p" Then Let NextChr$ = "P"
If NextChr$ = "q" Then Let NextChr$ = "Q"
If NextChr$ = "r" Then Let NextChr$ = "R"
If NextChr$ = "s" Then Let NextChr$ = "S"
If NextChr$ = "t" Then Let NextChr$ = "T"
If NextChr$ = "w" Then Let NextChr$ = "W"
If NextChr$ = "v" Then Let NextChr$ = "V"
If NextChr$ = " " Then Let NextChr$ = " "
Let Newsent$ = Newsent$ + NextChr$
Loop
IM_Hacker = Newsent$
Call SendIM(who, "" + IM_Hacker + "")
End Function
Function AIMChangeAIMCaption(Caption$)
'changes AIM's caption
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function
Public Function FindRoom() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindRoom& = child&
End Function
Public Sub MoveForm(Frm As Form)
ReleaseCapture
X = SendMessage(Frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'To use this,  put the following code in the "Mousedown"  dec
'of a label or picture box *Replace frm with your formname.
'MoveForm(frm)

End Sub
Function AIMChangeChatCaption(Caption$)
'changes the chatroom's caption
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_ChatWnd", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function
Public Function GetUser() As String
    Dim AOL As Long, MDI As Long, welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    GetUser$ = ""
End Function


Sub SendChat3(Chat)
' Replace the old send chat in the .bas witht this new one
' In improved sendchat, it takes any original text out, puts your text in
' waits till the text is sent (not manually), then puts the old text back in
Room& = FindChatRoom
AORich& = FindChildByClass(Room&, "RICHCNTL")

AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)

old = GetText(AORich&)
Call SendMessageByString(AORich&, WM_SETTEXT, 0, "")
Call SendMessageByString(AORich&, WM_SETTEXT, 0, Chat)
Call SendMessageByNum(AORich&, WM_CHAR, 13, 0)
Do Until GetText(AORich&) = ""
Timeout (1E-99)
Loop
Call SendMessageByString(AORich&, WM_SETTEXT, 0, old)
End Sub
Function ClickButton(Button)
Click& = SendMessageByNum(Button, WM_LBUTTONDOWN, &HD, 0)
Click& = SendMessageByNum(Button, WM_LBUTTONUP, &HD, 0)
End Function

Function SendMail30(Person, subject, message)
Call RunMenuByString("Compose Mail")
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "Compose Mail")
Timeout (5)
Check:
If Box& <> 0 Then
GoTo send
Else
Timeout (2)
GoTo Check
End If
send:
Edit& = FindWindowEx(Box&, 0&, "_AOL_Edit", vbNullString)
Edit& = FindWindowEx(Box&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(Edit&, WM_SETTEXT, 0, Person)
Edit& = FindWindowEx(Box&, Edit&, "_AOL_Edit", vbNullString)
Edit& = FindWindowEx(Box&, Edit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(Edit&, WM_SETTEXT, 0, subject)
send& = FindWindowEx(Box&, 0&, "_AOL_Icon", vbNullString)
Edit2& = FindWindowEx(Box&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(Edit2&, WM_SETTEXT, 0, message)
ClickIcon (send&)
End Function
Function AIMChangeIMCaption(Caption$)
'changes the caption of an IM
Dim caption1 As Long, captionchange As Long
caption1& = FindWindow("AIM_IMessage", vbNullString)
captionchange& = SendMessageByString(caption1&, WM_SETTEXT, 0, Caption$)
End Function
Function AIMclearChat()
'clears the chat text
Dim Chat As Long, Box As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
Call SendMessageByString(Box&, WM_SETTEXT, 0, "")

End Function
Function text_elite(message$)
'sends your message using ascii characters
On Error Resume Next
Do
If InStr(message$, "A") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "A") - 1) + "Ã" + Right$(message$, Len(message$) - InStr(message$, "A"))
message$ = macstringz
Loop Until InStr(message$, "A") = 0
Do
If InStr(message$, "a") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "a") - 1) + "å" + Right$(message$, Len(message$) - InStr(message$, "a"))
message$ = macstringz
Loop Until InStr(message$, "a") = 0
Do
If InStr(message$, "B") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "B") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "B"))
message$ = macstringz
Loop Until InStr(message$, "B") = 0
Do
If InStr(message$, "b") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "b") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "b"))
message$ = macstringz
Loop Until InStr(message$, "b") = 0
Do
If InStr(message$, "C") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "C") - 1) + "Ç" + Right$(message$, Len(message$) - InStr(message$, "C"))
message$ = macstringz
Loop Until InStr(message$, "C") = 0
Do
If InStr(message$, "c") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "c") - 1) + "ç" + Right$(message$, Len(message$) - InStr(message$, "c"))
message$ = macstringz
Loop Until InStr(message$, "c") = 0
Do
If InStr(message$, "D") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "D") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "D"))
message$ = macstringz
Loop Until InStr(message$, "D") = 0
Do
If InStr(message$, "d") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "d") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "d"))
message$ = macstringz
Loop Until InStr(message$, "d") = 0
Do
If InStr(message$, "E") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "E") - 1) + "Ê" + Right$(message$, Len(message$) - InStr(message$, "E"))
message$ = macstringz
Loop Until InStr(message$, "E") = 0
Do
If InStr(message$, "e") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "e") - 1) + "ê" + Right$(message$, Len(message$) - InStr(message$, "e"))
message$ = macstringz
Loop Until InStr(message$, "e") = 0
Do
If InStr(message$, "F") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "F") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "F"))
message$ = macstringz
Loop Until InStr(message$, "F") = 0
Do
If InStr(message$, "f") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "f") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "f"))
message$ = macstringz
Loop Until InStr(message$, "f") = 0
Do
If InStr(message$, "I") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "I") - 1) + "Ï" + Right$(message$, Len(message$) - InStr(message$, "I"))
message$ = macstringz
Loop Until InStr(message$, "I") = 0
Do
If InStr(message$, "i") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "i") - 1) + "ï" + Right$(message$, Len(message$) - InStr(message$, "i"))
message$ = macstringz
Loop Until InStr(message$, "i") = 0
Do
If InStr(message$, "L") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "L") - 1) + "£" + Right$(message$, Len(message$) - InStr(message$, "L"))
message$ = macstringz
Loop Until InStr(message$, "L") = 0
Do
If InStr(message$, "N") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "N") - 1) + "Ñ" + Right$(message$, Len(message$) - InStr(message$, "N"))
message$ = macstringz
Loop Until InStr(message$, "N") = 0
Do
If InStr(message$, "n") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "n") - 1) + "ñ" + Right$(message$, Len(message$) - InStr(message$, "n"))
message$ = macstringz
Loop Until InStr(message$, "n") = 0
Do
If InStr(message$, "O") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "O") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "O"))
message$ = macstringz
Loop Until InStr(message$, "O") = 0
Do
If InStr(message$, "o") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "o") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "o"))
message$ = macstringz
Loop Until InStr(message$, "o") = 0
Do
If InStr(message$, "R") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "R") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "R"))
message$ = macstringz
Loop Until InStr(message$, "R") = 0
Do
If InStr(message$, "r") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "r") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "r"))
message$ = macstringz
Loop Until InStr(message$, "r") = 0
Do
If InStr(message$, "S") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "S") - 1) + "§" + Right$(message$, Len(message$) - InStr(message$, "S"))
message$ = macstringz
Loop Until InStr(message$, "S") = 0
Do
If InStr(message$, "s") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "s") - 1) + "š" + Right$(message$, Len(message$) - InStr(message$, "s"))
message$ = macstringz
Loop Until InStr(message$, "s") = 0
Do
If InStr(message$, "T") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "T") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "T"))
message$ = macstringz
Loop Until InStr(message$, "T") = 0
Do
If InStr(message$, "t") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "t") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "t"))
message$ = macstringz
Loop Until InStr(message$, "t") = 0
Do
If InStr(message$, "U") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "U") - 1) + "Ü" + Right$(message$, Len(message$) - InStr(message$, "U"))
message$ = macstringz
Loop Until InStr(message$, "U") = 0
Do
If InStr(message$, "u") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "u") - 1) + "ú" + Right$(message$, Len(message$) - InStr(message$, "u"))
message$ = macstringz
Loop Until InStr(message$, "u") = 0
Do
If InStr(message$, "X") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "X") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "X"))
message$ = macstringz
Loop Until InStr(message$, "X") = 0
Do
If InStr(message$, "x") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "x") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "x"))
message$ = macstringz
Loop Until InStr(message$, "x") = 0
Do
If InStr(message$, "Y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "Y") - 1) + "¥" + Right$(message$, Len(message$) - InStr(message$, "Y"))
message$ = macstringz
Loop Until InStr(message$, "Y") = 0
Do
If InStr(message$, "y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "y") - 1) + "ÿ" + Right$(message$, Len(message$) - InStr(message$, "y"))
message$ = macstringz
Loop Until InStr(message$, "y") = 0
SendChat message$
End Function
Function IM_elite(who, message$)
'sends your message using ascii characters
On Error Resume Next
Do
If InStr(message$, "A") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "A") - 1) + "Ã" + Right$(message$, Len(message$) - InStr(message$, "A"))
message$ = macstringz
Loop Until InStr(message$, "A") = 0
Do
If InStr(message$, "a") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "a") - 1) + "å" + Right$(message$, Len(message$) - InStr(message$, "a"))
message$ = macstringz
Loop Until InStr(message$, "a") = 0
Do
If InStr(message$, "B") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "B") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "B"))
message$ = macstringz
Loop Until InStr(message$, "B") = 0
Do
If InStr(message$, "b") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "b") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "b"))
message$ = macstringz
Loop Until InStr(message$, "b") = 0
Do
If InStr(message$, "C") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "C") - 1) + "Ç" + Right$(message$, Len(message$) - InStr(message$, "C"))
message$ = macstringz
Loop Until InStr(message$, "C") = 0
Do
If InStr(message$, "c") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "c") - 1) + "ç" + Right$(message$, Len(message$) - InStr(message$, "c"))
message$ = macstringz
Loop Until InStr(message$, "c") = 0
Do
If InStr(message$, "D") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "D") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "D"))
message$ = macstringz
Loop Until InStr(message$, "D") = 0
Do
If InStr(message$, "d") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "d") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "d"))
message$ = macstringz
Loop Until InStr(message$, "d") = 0
Do
If InStr(message$, "E") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "E") - 1) + "Ê" + Right$(message$, Len(message$) - InStr(message$, "E"))
message$ = macstringz
Loop Until InStr(message$, "E") = 0
Do
If InStr(message$, "e") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "e") - 1) + "ê" + Right$(message$, Len(message$) - InStr(message$, "e"))
message$ = macstringz
Loop Until InStr(message$, "e") = 0
Do
If InStr(message$, "F") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "F") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "F"))
message$ = macstringz
Loop Until InStr(message$, "F") = 0
Do
If InStr(message$, "f") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "f") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "f"))
message$ = macstringz
Loop Until InStr(message$, "f") = 0
Do
If InStr(message$, "I") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "I") - 1) + "Ï" + Right$(message$, Len(message$) - InStr(message$, "I"))
message$ = macstringz
Loop Until InStr(message$, "I") = 0
Do
If InStr(message$, "i") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "i") - 1) + "ï" + Right$(message$, Len(message$) - InStr(message$, "i"))
message$ = macstringz
Loop Until InStr(message$, "i") = 0
Do
If InStr(message$, "L") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "L") - 1) + "£" + Right$(message$, Len(message$) - InStr(message$, "L"))
message$ = macstringz
Loop Until InStr(message$, "L") = 0
Do
If InStr(message$, "N") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "N") - 1) + "Ñ" + Right$(message$, Len(message$) - InStr(message$, "N"))
message$ = macstringz
Loop Until InStr(message$, "N") = 0
Do
If InStr(message$, "n") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "n") - 1) + "ñ" + Right$(message$, Len(message$) - InStr(message$, "n"))
message$ = macstringz
Loop Until InStr(message$, "n") = 0
Do
If InStr(message$, "O") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "O") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "O"))
message$ = macstringz
Loop Until InStr(message$, "O") = 0
Do
If InStr(message$, "o") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "o") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "o"))
message$ = macstringz
Loop Until InStr(message$, "o") = 0
Do
If InStr(message$, "R") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "R") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "R"))
message$ = macstringz
Loop Until InStr(message$, "R") = 0
Do
If InStr(message$, "r") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "r") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "r"))
message$ = macstringz
Loop Until InStr(message$, "r") = 0
Do
If InStr(message$, "S") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "S") - 1) + "§" + Right$(message$, Len(message$) - InStr(message$, "S"))
message$ = macstringz
Loop Until InStr(message$, "S") = 0
Do
If InStr(message$, "s") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "s") - 1) + "š" + Right$(message$, Len(message$) - InStr(message$, "s"))
message$ = macstringz
Loop Until InStr(message$, "s") = 0
Do
If InStr(message$, "T") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "T") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "T"))
message$ = macstringz
Loop Until InStr(message$, "T") = 0
Do
If InStr(message$, "t") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "t") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "t"))
message$ = macstringz
Loop Until InStr(message$, "t") = 0
Do
If InStr(message$, "U") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "U") - 1) + "Ü" + Right$(message$, Len(message$) - InStr(message$, "U"))
message$ = macstringz
Loop Until InStr(message$, "U") = 0
Do
If InStr(message$, "u") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "u") - 1) + "ú" + Right$(message$, Len(message$) - InStr(message$, "u"))
message$ = macstringz
Loop Until InStr(message$, "u") = 0
Do
If InStr(message$, "X") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "X") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "X"))
message$ = macstringz
Loop Until InStr(message$, "X") = 0
Do
If InStr(message$, "x") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "x") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "x"))
message$ = macstringz
Loop Until InStr(message$, "x") = 0
Do
If InStr(message$, "Y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "Y") - 1) + "¥" + Right$(message$, Len(message$) - InStr(message$, "Y"))
message$ = macstringz
Loop Until InStr(message$, "Y") = 0
Do
If InStr(message$, "y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "y") - 1) + "ÿ" + Right$(message$, Len(message$) - InStr(message$, "y"))
message$ = macstringz
Loop Until InStr(message$, "y") = 0
Call SendIM(who, message$)
End Function

Function AIMEliteTalker(message$)
'sends your message using ascii characters
On Error Resume Next
Do
If InStr(message$, "A") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "A") - 1) + "Ã" + Right$(message$, Len(message$) - InStr(message$, "A"))
message$ = macstringz
Loop Until InStr(message$, "A") = 0
Do
If InStr(message$, "a") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "a") - 1) + "å" + Right$(message$, Len(message$) - InStr(message$, "a"))
message$ = macstringz
Loop Until InStr(message$, "a") = 0
Do
If InStr(message$, "B") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "B") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "B"))
message$ = macstringz
Loop Until InStr(message$, "B") = 0
Do
If InStr(message$, "b") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "b") - 1) + "ß" + Right$(message$, Len(message$) - InStr(message$, "b"))
message$ = macstringz
Loop Until InStr(message$, "b") = 0
Do
If InStr(message$, "C") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "C") - 1) + "Ç" + Right$(message$, Len(message$) - InStr(message$, "C"))
message$ = macstringz
Loop Until InStr(message$, "C") = 0
Do
If InStr(message$, "c") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "c") - 1) + "ç" + Right$(message$, Len(message$) - InStr(message$, "c"))
message$ = macstringz
Loop Until InStr(message$, "c") = 0
Do
If InStr(message$, "D") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "D") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "D"))
message$ = macstringz
Loop Until InStr(message$, "D") = 0
Do
If InStr(message$, "d") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "d") - 1) + "Ð" + Right$(message$, Len(message$) - InStr(message$, "d"))
message$ = macstringz
Loop Until InStr(message$, "d") = 0
Do
If InStr(message$, "E") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "E") - 1) + "Ê" + Right$(message$, Len(message$) - InStr(message$, "E"))
message$ = macstringz
Loop Until InStr(message$, "E") = 0
Do
If InStr(message$, "e") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "e") - 1) + "ê" + Right$(message$, Len(message$) - InStr(message$, "e"))
message$ = macstringz
Loop Until InStr(message$, "e") = 0
Do
If InStr(message$, "F") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "F") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "F"))
message$ = macstringz
Loop Until InStr(message$, "F") = 0
Do
If InStr(message$, "f") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "f") - 1) + "ƒ" + Right$(message$, Len(message$) - InStr(message$, "f"))
message$ = macstringz
Loop Until InStr(message$, "f") = 0
Do
If InStr(message$, "I") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "I") - 1) + "Ï" + Right$(message$, Len(message$) - InStr(message$, "I"))
message$ = macstringz
Loop Until InStr(message$, "I") = 0
Do
If InStr(message$, "i") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "i") - 1) + "ï" + Right$(message$, Len(message$) - InStr(message$, "i"))
message$ = macstringz
Loop Until InStr(message$, "i") = 0
Do
If InStr(message$, "L") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "L") - 1) + "£" + Right$(message$, Len(message$) - InStr(message$, "L"))
message$ = macstringz
Loop Until InStr(message$, "L") = 0
Do
If InStr(message$, "N") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "N") - 1) + "Ñ" + Right$(message$, Len(message$) - InStr(message$, "N"))
message$ = macstringz
Loop Until InStr(message$, "N") = 0
Do
If InStr(message$, "n") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "n") - 1) + "ñ" + Right$(message$, Len(message$) - InStr(message$, "n"))
message$ = macstringz
Loop Until InStr(message$, "n") = 0
Do
If InStr(message$, "O") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "O") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "O"))
message$ = macstringz
Loop Until InStr(message$, "O") = 0
Do
If InStr(message$, "o") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "o") - 1) + "Ø" + Right$(message$, Len(message$) - InStr(message$, "o"))
message$ = macstringz
Loop Until InStr(message$, "o") = 0
Do
If InStr(message$, "R") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "R") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "R"))
message$ = macstringz
Loop Until InStr(message$, "R") = 0
Do
If InStr(message$, "r") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "r") - 1) + "®" + Right$(message$, Len(message$) - InStr(message$, "r"))
message$ = macstringz
Loop Until InStr(message$, "r") = 0
Do
If InStr(message$, "S") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "S") - 1) + "§" + Right$(message$, Len(message$) - InStr(message$, "S"))
message$ = macstringz
Loop Until InStr(message$, "S") = 0
Do
If InStr(message$, "s") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "s") - 1) + "š" + Right$(message$, Len(message$) - InStr(message$, "s"))
message$ = macstringz
Loop Until InStr(message$, "s") = 0
Do
If InStr(message$, "T") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "T") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "T"))
message$ = macstringz
Loop Until InStr(message$, "T") = 0
Do
If InStr(message$, "t") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "t") - 1) + "†" + Right$(message$, Len(message$) - InStr(message$, "t"))
message$ = macstringz
Loop Until InStr(message$, "t") = 0
Do
If InStr(message$, "U") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "U") - 1) + "Ü" + Right$(message$, Len(message$) - InStr(message$, "U"))
message$ = macstringz
Loop Until InStr(message$, "U") = 0
Do
If InStr(message$, "u") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "u") - 1) + "ú" + Right$(message$, Len(message$) - InStr(message$, "u"))
message$ = macstringz
Loop Until InStr(message$, "u") = 0
Do
If InStr(message$, "X") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "X") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "X"))
message$ = macstringz
Loop Until InStr(message$, "X") = 0
Do
If InStr(message$, "x") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "x") - 1) + "×" + Right$(message$, Len(message$) - InStr(message$, "x"))
message$ = macstringz
Loop Until InStr(message$, "x") = 0
Do
If InStr(message$, "Y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "Y") - 1) + "¥" + Right$(message$, Len(message$) - InStr(message$, "Y"))
message$ = macstringz
Loop Until InStr(message$, "Y") = 0
Do
If InStr(message$, "y") = 0 Then Exit Do
macstringz = Left$(message$, InStr(message$, "y") - 1) + "ÿ" + Right$(message$, Len(message$) - InStr(message$, "y"))
message$ = macstringz
Loop Until InStr(message$, "y") = 0
AIMSendChat message$
End Function



Function AIMFindChatRoom()
'looks for a chatroom
Dim Chat As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
FindChatRoom = Chat&
End Function

Function AIMGetChatName()
'gets the name of the chat room
'This won't work if you've changed the chat caption!
Dim Chat As Long
On Error Resume Next
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
RoomName$ = GetCaption(Chat&)
Room$ = Mid(RoomName$, InStr(RoomName$, ":") + 2)
GetChatName = Room$
End Function
Function AIMGetchatText()
'gets the text from a chat
'sorry, but it gets the text with full HTML
'so it's kinda useless, but I put in here anyway
Dim Chat As Long, ChatWin As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
ChatWin& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
it = GetText(ChatWin&)
GetChatText = it
End Function
Function AIMWebPage(page$)
'Goes to a webpage using AIM
Dim aim As Long, Box As Long, Go As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Box& = FindWindowEx(aim&, 0&, "Edit", vbNullString)
Call SendMessageByString(Box&, WM_SETTEXT, 0, page$)
Go& = FindWindowEx(aim&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Go&)
End Function

Function AIMHide()
'Hides the AIM window
Dim aim As Long, HideIt As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
HideIt& = ShowWindow(aim&, SW_HIDE)
End Function
Sub Mail_ForwardMail(Recipiants, message)

AOL% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
MDI% = FindChildByClass(AOL%, "MDIClient")
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 14
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon


AOLIcon (AOIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(MDI%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Sub AOLIcon(Icon%)
Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub AIMHideAIMAd()
'Hides that annoying advertisement
Dim part1 As Long, part2 As Long, HideIt As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
HideIt& = ShowWindow(part2&, SW_HIDE)
End Sub

Function GetChatName()
Room& = FindChatRoom
RoomName$ = GetCaption(Room&)
GetChatName = RoomName$
End Function

Function IsAIMOnline()
'sees if AIM is online
Dim aim As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
IsAIMOnline = aim&
End Function
Function MassIM(List As ListBox, message As String)
'Mass IM the people on your list
'You must have a listbox with screen names
'and a message to send!
If List.ListCount = 0 Then
Do: DoEvents: Loop
End If
List.Enabled = False
i = List.ListCount - 1
List.ListIndex = 0
For X = 0 To i
List.ListIndex = X
Call SendIM(List.Text, message)
Timeout (2)
Next X
List.Enabled = True
End Function
Function AIMMassIM(List As ListBox, message As String)
'Mass IM the people on your list
'You must have a listbox with screen names
'and a message to send!
If List.ListCount = 0 Then
Do: DoEvents: Loop
End If
List.Enabled = False
i = List.ListCount - 1
List.ListIndex = 0
For X = 0 To i
List.ListIndex = X
Call AIMSendIM(List.Text, message)
Timeout (0.8)
Next X
List.Enabled = True
End Function
Function AIMopenIM()
'Opens an IM
Dim aim As Long, Group As Long, Button As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
End Function

Function AIMSayRoomName()
'Says the room name
'Won't work if you've changed the chat caption!
AIMSendChat "You have just entered room """ + GetChatName + "."""
End Function

Function AIMScroller(message$)
'Scrolls 10 lines of your message
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
SendChat message$
Timeout (0.25)
AIMSendChat message$
End Function
Function AIMSendChat(message$)
'Sends you're message to a chat
Dim Chat As Long, Box As Long, Box2 As Long, Box3 As Long, send As Long, send2 As Long, send3 As Long, send4 As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
Box& = FindWindowEx(Chat&, 0&, "WndAte32Class", vbNullString)
Box2& = FindWindowEx(Chat&, Box&, "WndAte32Class", vbNullString)
Box3& = SendMessageByString(Box2&, WM_SETTEXT, 0, message$)
send& = FindWindowEx(Chat&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Chat&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Chat&, send2&, "_Oscar_IconBtn", vbNullString)
send4& = FindWindowEx(Chat&, send3&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send4&)
End Function
Function AIMSendIM(who$, message$)
'Sends an IM to someone
Dim aim As Long, Group As Long, Button As Long, IM As Long, IMcombo As Long, IMto As Long, IMmessage As Long, send As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button&)
IM& = FindWindow("AIM_IMessage", vbNullString)
IMcombo& = FindWindowEx(IM&, 0&, "_Oscar_PersistantCombo", vbNullString)
IMto& = FindWindowEx(IMcombo&, 0&, "Edit", vbNullString)
Call SendMessageByString(IMto&, WM_SETTEXT, 0, who$)
IMmessage& = FindWindowEx(IM&, 0&, "WndAte32Class", vbNullString)
IMmessage& = GetWindow(IMmessage&, 2)
Call SendMessageByString(IMmessage&, WM_SETTEXT, 0, message$)
send& = FindWindowEx(IM&, 0&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send&)
End Function

Function AIMSendInvite(who$, message$, Chat$)
'Sends an chat invitation to someone
Dim aim As Long, Group As Long, Button As Long, Button2 As Long, Invite As Long, Edit1 As Long, Edit2 As Long, Edit3 As Long, send As Long, send2 As Long, send3 As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Group& = FindWindowEx(aim&, 0&, "_Oscar_TabGroup", vbNullString)
Button& = FindWindowEx(Group&, 0&, "_Oscar_IconBtn", vbNullString)
Button2& = FindWindowEx(Group&, Button&, "_Oscar_IconBtn", vbNullString)
ClickIcon (Button2&)
Invite& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Edit1& = FindWindowEx(Invite&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit1&, WM_SETTEXT, 0, who$)
Edit2& = FindWindowEx(Invite&, Edit1&, "Edit", vbNullString)
Call SendMessageByString(Edit2&, WM_SETTEXT, 0, message$)
Edit3& = FindWindowEx(Invite&, Edit2&, "Edit", vbNullString)
Call SendMessageByString(Edit3&, WM_SETTEXT, 0, Chat$)
send& = FindWindowEx(Invite&, 0&, "_Oscar_IconBtn", vbNullString)
send2& = FindWindowEx(Invite&, send&, "_Oscar_IconBtn", vbNullString)
send3& = FindWindowEx(Invite&, send2&, "_Oscar_IconBtn", vbNullString)
ClickIcon (send3&)
End Function
Function AIMSendLink(link$, message$)
'Sends a link in a chat
SendChat "<a href=""" + link$ + """>" + message$ + ""
End Function


Function AIMShow()
'Shows the AIM window
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Showit& = ShowWindow(aim&, SW_SHOW)
End Function

Function AIMShowAd()
'Shows that gay advertisement
Dim part1 As Long, part2 As Long, Showit As Long
part1& = FindWindow("_Oscar_BuddyListWin", vbNullString)
part2& = FindWindowEx(part1&, 0&, "WndAte32Class", vbNullString)
Showit& = ShowWindow(part2&, SW_SHOW)
End Function
Function SendLink(link, message)
SendChat "<font face=arial><B></I><U><a href=""""><a href=""""><a href=""" + link + """>" + message + "<font color=#fefefe></a>"
End Function

Function signoffAIM()
'signs off AIM
End Function

Function AIMSNfromIM()
'gets the SN of the sender of an IM
'won't work if you've changed the IM caption!
Dim IM As Long
On Error Resume Next
IM& = FindWindow("AIM_IMessage", vbNullString)
Name$ = GetCaption(IM&)
If InStr(Name$, "- Instant Message") <> 0 Then
Text% = GetWindowTextLength(IM&)
Text% = (Text%) - 19
SN$ = Left$(Name$, InStr(Name$, "" + Name$ + "") + Text%)
SNfromIM = SN$
Else
SNfromIM = "(Unknown)"
End If
End Function
Function AIMUserSN()
'Get's the User's SN
'Won't work if you've changed the AIM caption!!!
Dim aim As Long
On Error Resume Next
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Name$ = GetCaption(aim&)
If InStr(Name$, "Buddy List") <> 0 Then
Text% = GetWindowTextLength(aim&)
Text% = (Text%) - 14
SN$ = Left$(Name$, InStr(Name$, "" + Name$ + "") + Text%)
UserSN = SN$
Else
UserSN = "(Unknown)"
End If
End Function
Function mail_count()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "'s Online Mailbox")
Control& = FindWindowEx(Box&, 0&, "_AOL_TabControl", vbNullString)
TabPage& = FindWindowEx(Control&, 0&, "_AOL_TabPage", vbNullString)
MailList& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
mail_count = SendMessage(MailList&, LB_GETCOUNT, 0, 0)
End Function
Function mail_findbox()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "'s Online Mailbox")
mail_findbox = Box&
End Function

Function mail_Read()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "'s Online Mailbox")
Read& = FindWindowEx(Box&, 0&, "_AOL_Icon", vbNullString)
ClickButton (Read&)
End Function

Sub mail_ClickRead()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
MailBox% = FindChildByTitle(MDI%, UserSN & "'s Online Mailbox")
AOIcon% = FindChildByClass(MailBox%, "_AOL_Icon")
For L = 1 To 0
AOIcon% = GetWindow(AOIcon%, 2)
Next L
ClickIcon (AOIcon%)
End Sub
Sub Mail_ClickForward()
MailWin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(MailWin%, "_AOL_Icon")
For L = 1 To 8
AOIcon% = GetWindow(AOIcon%, 2)
NoFreeze% = DoEvents()
Next L

AOLIcon (AOIcon%)
End Sub
Function AOLMDI()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function
Function mail_Status()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "'s Online Mailbox")
Icon& = FindWindowEx(Box&, 0&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Box&, Icon&, "_AOL_Icon", vbNullString)
ClickButton (Icon&)
End Function
Function mail_Delete()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "'s Online Mailbox")
Icon& = FindWindowEx(Box&, 0&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Box&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Box&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Box&, Icon&, "_AOL_Icon", vbNullString)
ClickButton (Icon&)
End Function
Function Mail_KeepAsNew()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindChildByTitle(MDI&, "'s Online Mailbox")
Icon& = FindWindowEx(Box&, 0&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Box&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Box&, Icon&, "_AOL_Icon", vbNullString)
ClickButton (Icon&)
End Function
Function MessageFromIM()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
IM& = FindChildByTitle(MDI&, ">Instant Message From:")
If IM& Then GoTo Greed
IM& = FindChildByTitle(MDI&, "  Instant Message From:")
If IM& Then GoTo Greed
Exit Function
Greed:
IMText& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
IMmessage = GetText(IMText&)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
message = Mid(IMmessage, InStr(SN, ":") + snlen)
MessageFromIM = message
End Function
Function AOL_Show()
AOL& = FindWindow("AOL Frame25", vbNullString)
X = ShowWindow(AOL&, SW_SHOW)
End Function
Function SNfromIM()

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)

IM& = FindChildByTitle(MDI&, ">Instant Message From:")
If IM& Then GoTo Greed
IM& = FindChildByTitle(MDI&, "  Instant Message From:")
If IM& Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM&)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function
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


Function Toolbar_Hide()
AOL& = FindWindow("AOL Frame25", vbNullString)
    bar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    X = ShowWindow(bar&, SW_HIDE)
End Function

Function Toolbar_Show()
AOL& = FindWindow("AOL Frame25", vbNullString)
    bar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    X = ShowWindow(bar&, SW_SHOW)
End Function
Function Upchat()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
MailBox& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
upload% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(upload%, SW_HIDE)
X = ShowWindow(upload%, SW_MINIMIZE)
X = ShowWindow(MailBox&, SW_MINIMIZE)
Call AOLSetFocus
End Function
Function FindChatRoom()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Room& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Stuff& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
MoreStuff& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
If Stuff& <> 0 And MoreStuff& <> 0 Then
   FindChatRoom = Room&
Else:
   FindChatRoom = 0
End If
End Function

Function Keyword(txt)
    AOL& = FindWindow("AOL Frame25", vbNullString)
    bar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    bar& = FindWindowEx(bar&, 0&, "_AOL_Toolbar", vbNullString)
    bar& = FindWindowEx(bar&, 0&, "_AOL_Combobox", vbNullString)
    Box& = FindWindowEx(bar&, 0&, "Edit", vbNullString)
    Call SendMessageByString(Box&, WM_SETTEXT, 0, txt)
    Call SendMessageByNum(Box&, WM_CHAR, VK_SPACE, 0)
    Call SendMessageByNum(Box&, WM_CHAR, VK_RETURN, 0)
End Function
Function FindChildByTitle(Parent, child As String) As Integer
childfocus% = GetWindow(Parent, 5)

While childfocus%
hwndLength% = GetWindowTextLength(childfocus%)
buffer$ = String$(hwndLength%, 0)
WindowText% = GetWindowText(childfocus%, buffer$, (hwndLength% + 1))

If InStr(UCase(buffer$), UCase(child)) Then FindChildByTitle = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend

End Function
Function KillListDupes(List)
For i = 0 To List.ListCount - 1
For E = 0 To List.ListCount - 1

If LCase(List.List(i)) Like LCase(List.List(E)) And i <> E Then

List.RemoveItem (E)
End If

Next E
Next i
End Function

Sub SendIM(who, message)
Call Keyword("aol://9293:" + who + "")
Timeout (2)
Do
DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
send& = FindWindowEx(IM&, 0&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(IM&, send&, "_AOL_Icon", vbNullString)
Loop Until (IM& <> 0 And Rich& <> 0 And send& <> 0)
Call SendMessageByString(Rich&, WM_SETTEXT, 0, message)
imsend& = GetWindow(send&, GW_HWNDNEXT)
ClickIcon (send&)
If IM& Then KillWindow (IM&)

End Sub


Function SignOff()
AOL& = FindWindow("AOL Frame25", "America  Online")
Call RunMenuByString("Sign Off")
End Function

Sub Timeout(Duration)
StartTime = Timer
Do While Timer - StartTime < Duration
DoEvents
Loop

End Sub
Sub StayOnTop(the As Form)
'keeps the form on top of everything
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub


Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function


Function GetClass(child)
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function
Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function
Sub ClickIcon(Icon%)
Click% = SendMessage(Icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(Icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function LastChatLineWithSN()
ChatText$ = GetChatText

For FindChar = 1 To Len(ChatText$)

thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(ChatText$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function
Function LastIMLineWithSN()
ChatText$ = MessageFromIM

For FindChar = 1 To Len(ChatText$)

thechar$ = Mid(ChatText$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(ChatText$, lastlen, Len(thechars$))

LastIMLineWithSN = lastline
End Function

Function GetChatText()
Room& = FindChatRoom
AORich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
ChatText = GetText(AORich&)
GetChatText = ChatText
End Function
Function LastChatLine()
ChatText = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Function SNFromLastChatLine()
ChatText$ = LastChatLineWithSN
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function

Public Sub AddRoomToList(TheList As ListBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindChatRoom
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                TheList.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
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
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub
Sub WaitOKorChat()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0 Or FindRoom <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub
Function mail_send(Recipiants, subject, message)

AOL& = FindWindow("AOL Frame25", vbNullString)
AOTooL& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
AOTool2& = FindWindowEx(AOTooL&, 0&, "_AOL_Toolbar", vbNullString)
AOIcon& = FindWindowEx(AOTool2&, 0&, "_AOL_Icon", vbNullString)

AOIcon& = GetWindow(AOIcon&, 2)

ClickIcon (AOIcon&)

Do: DoEvents
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AOMail& = FindChildByTitle(MDI&, "Write Mail")
AOEdit& = FindWindowEx(AOMail&, 0&, "_AOL_Edit", vbNullString)
AORich& = FindWindowEx(AOMail&, 0&, "RICHCNTL", vbNullString)
AOIcon& = FindWindowEx(AOMail&, 0&, "_AOL_Icon", vbNullString)
Loop Until AOMail& <> 0 And AOEdit& <> 0 And AORich& <> 0 And AOIcon& <> 0

Call SendMessageByString(AOEdit&, WM_SETTEXT, 0, Recipiants)

AOEdit& = GetWindow(AOEdit&, 2)
AOEdit& = GetWindow(AOEdit&, 2)
AOEdit& = GetWindow(AOEdit&, 2)
AOEdit& = GetWindow(AOEdit&, 2)
Call SendMessageByString(AOEdit&, WM_SETTEXT, 0, subject)
Call SendMessageByString(AORich&, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon& = GetWindow(AOIcon&, 2)
Next GetIcon

ClickIcon (AOIcon&)

Do: DoEvents
AOError& = FindChildByTitle(MDI&, "Error")
AOModal& = FindWindow("_AOL_Modal", vbNullString)
If AOMail& = 0 Then Exit Do
If AOModal& <> 0 Then
AOIcon& = FindWindowEx(AOModal&, 0&, "_AOL_Icon", vbNullString)
ClickIcon (AOIcon&)
Call SendMessage(AOMail&, WM_CLOSE, 0, 0)
Exit Function
End If
If AOError& <> 0 Then
Call SendMessage(AOError&, WM_CLOSE, 0, 0)
Call SendMessage(AOMail&, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Function
Public Sub AddRoomToCombo(TheCombo As ComboBox, AddUser As Boolean)
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindChatRoom
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser$ Or AddUser = True Then
                TheCombo.AddItem ScreenName$
            End If
        Next index&
        Call CloseHandle(mThread)
    End If
    If TheCombo.ListCount > 0 Then
       TheCombo.Text = TheCombo.List(0)
    End If
End Sub
Public Sub RunMenuByString(SearchString As String)
    Dim AOL As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(AOL&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(AOL&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Sub SendChat(Text)
Room& = FindChatRoom
AORich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)

AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
AORich& = GetWindow(AORich&, 2)
Call SendMessageByString(AORich&, WM_SETTEXT, 0, Text)
DoEvents
Call SendMessageByNum(AORich&, WM_CHAR, 13, 0)
End Sub

Function IdleBot()
IdleWin& = FindWindow("_AOL_Modal", vbNullString)
Button& = FindWindowEx(IdleWin&, 0&, "_AOL_Icon", vbNullString)
If IdleWin& <> 0 Then
ClickIcon (Button&)
End If
End Function

Function AOLWin()
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLWin = AOL&
End Function

Function UnUpchat()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
MailBox& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
upload% = FindWindow("_AOL_MODAL", vbNullString)
X = ShowWindow(upload%, SW_SHOW)
X = ShowWindow(upload%, SW_RESTORE)
X = ShowWindow(MailBox&, SW_RESTORE)
Call AOLSetFocus
End Function
Function Mail_Open()
AOL& = FindWindow("AOL Frame25", vbNullString)
    bar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    bar& = FindWindowEx(bar&, 0&, "_AOL_Toolbar", vbNullString)
    mailicon& = FindWindowEx(bar&, 0&, "_AOL_Icon", vbNullString)
    ClickIcon (mailicon&)
End Function
Sub AOLSetFocus()
X = GetCaption(AOLWin)
AppActivate X
End Sub
Function UserSN()
On Error Resume Next
AOL% = FindWindow("AOL Frame25", "America  Online")
MDI% = FindChildByClass(AOL%, "MDIClient")
welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(welcome%, WelcomeTitle$, (WelcomeLength% + 1))
user = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = user
If UserSN = "" Then
UserSN = Unknown
End If
End Function


Function KillWindow(Window%)
KillWindow = SendMessageByNum(Window%, WM_CLOSE, 0, 0)
End Function

Sub WaitOKorChat2()
Do
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
FULL% = FindChildByTitle(MDI%, "Sorry...")
DoEvents
If FindRoom <> 0 Then
Exit Sub
End If
Loop Until FULL% <> 0
      KillWindow (FULL%)
      KillWindow (FULL%)
      KillWindow (FULL%)
End Sub


