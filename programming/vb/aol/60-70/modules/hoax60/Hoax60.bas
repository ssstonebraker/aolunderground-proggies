Attribute VB_Name = "DaHoaxBas"
'  Listen up people, I could give a fuck less what you add to this module,
'  however, I do not want you removing ANYTHING from it!
'  Upon removing something, you risk ruining this whole damn module!
'  Everything, especially the "little shit" in this module is used, subs
'  are used with other subs, functions with sub, subs with functions,
'  functions with functions!  So just leave the existing code as is.
'  P.S. I'm not even gonna ask for greets in yo shit cause
'  No one ever adds the module maker to their greets.
'  ***NOTES***
'  (1)  There are 3 string functions sampled from Dos32.bas.
'  (2)  All the fade functions WERE writeen by me (Hoax).
'         I know the layout of the fade functions looks just like
'         they were taken from monkefade3.bas, but they weren't.
'         There is a formula to fades, and thats what I used,
'         Monkegod's formulas, thats it.  To proove that I did
'         NOT take his fades, you can look at my LaggIt function
'         and my ALLFX function.
'  (3)  I took one thing from someone else...the TOS_Phrases.
'         They were part of the Rampage Toolz 1.0 source code.
'         I used to use the modules from that source, but as you
'         can plainly see, I did not do the same with this module.
'  (4)  There are many functions and constants that work/do the
'         same thing as other functions/constants.  This is because
'         this module has been revised, edited, restarted, and revived
'         so many times that I ended up adding things twice without
'         knowing it.  Don't remove the constants/functions that
'         are duplicated.  They are extremely needed for other functions/subs/constants.
Public Enum txtStyles
wNumbers
wChars
End Enum
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Attribute SetFocus.VB_Description = "Sets focus to a specified hwnd/window."
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Attribute SetMenuItemBitmaps.VB_Description = "For adding icons to menus & sub-menus."
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Attribute Sleep.VB_Description = "Better than ""Pause"", more accurate, uses milliseconds."
Public Declare Function GetActiveWindow Lib "user32" () As Long
Attribute GetActiveWindow.VB_Description = "Returns a long integer value of the top-most window (window that ahs focus)."
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Attribute SendMessageByNum.VB_Description = "API, for sending messages & crap."
Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Attribute EnableWindow.VB_Description = "Sets a hwnd's Enabled property to True."
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Attribute GetClassName.VB_Description = "Returns a string value of the class name of the window with the specified hwnd/handel."
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Const RGN_OR = 2
Public Const WM_MOVE = &HF012
Public Const WM_SYSCOMMAND = &H112
Public lngRegion As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Attribute FindWindow.VB_Description = "Returns the handel/hwnd of a Window (cannot have a parent window)."
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Attribute FindWindowEx.VB_Description = "Returns the handel/hwnd of a window/control that has a parent window."
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Attribute GetCursorPos.VB_Description = "Returns a long integer value of the mouse's position.  Must specify x (horizontal position) or y (vertical position)."
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Attribute GetMenuItemCount.VB_Description = "Returns a long integer representing the number of items in a specified menu."
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Attribute GetMenuString.VB_Description = "Returns a string value of a menu."
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Attribute IsWindowVisible.VB_Description = "Returns a value wether or not a window is visible or not."
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Attribute MciSendString.VB_Description = "For Playing sounds."
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Attribute SendMessage.VB_Description = "API, for sending messages & crap."
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Attribute SendMessageLong.VB_Description = "API, for sending messages & crap."
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Attribute SendMessageByString.VB_Description = "API, for sending messages & crap."
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Attribute SetCursorPos.VB_Description = "Sets the cursor to a specific position."
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Attribute SetWindowPos.VB_Description = "Lets you dock windows or move other windows not related to your program."
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Attribute ShowCursor.VB_Description = "Shows the cursor if it is hidden."
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Attribute ShowWindow.VB_Description = "Shows a specified hidden window if it is hidden."
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Attribute SndPlaySound.VB_Description = "For playing sounds."
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
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
Public Const SW_Hide = 0
Public Const SW_SHOW = 5
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const WM_MDIRESTORE = &H223
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
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_SETTEXT = &HC
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_COPY = &H301
Public Const wNull = vbNullString
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Enum MailTypes
Normal
attach
End Enum

Public Enum AolVersions
Aol5
Aol6
End Enum

Public Const AolIcon = "_AOL_Icon"
Public Const AolEdit = "_AOL_Edit"
Public Const AolRichy = "RICHCNTL"
Public Const AolRadioBtn = "_AOL_RadioBox"
Public Const Ao_Child = "AOL Child"
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Enum Lists
wNew
wOld
wSent
End Enum

Public Enum winStyles
wHide
wShow
wMinimize
wMaximize
wClose
End Enum
Function abc(ax As Integer, Optional CapIt As Boolean = False) As String
Dim abcArray(26), bx As Integer
If CapIt = True Then bx% = 65 Else bx% = 97
For i = 0 To 25
abcArray(i) = Chr(i + bx%)
Next i
abc$ = abcArray(ax)

End Function
Function wAol&()
wAol& = FindWindow("AOL Frame25", vbNullString)
End Function
Function wMDI&()
wMDI& = FindWindowEx(wAol&, 0&, "MDIClient", vbNullString)
End Function
Function AolChild&(wCap As String)
AolChild& = FindWindowEx(wMDI&, 0&, "AOL Child", wCap$)
End Function
Function MailWIN&()
MailWIN& = AolKid&("Write Mail")
End Function
Function UpLoadWIN&()
UpLoadWIN& = FindWindowEx(0&, 0&, "_AOL_Modal", vbNullString)
End Function
Function UpLoadStatus$()
UpLoadStatus$ = Mid(GetCaption(UpLoadWIN&), 16)
End Function
Sub ChangeCaption(win&, txt$)
Call SendMessageByString(win&, WM_SETTEXT, 0&, txt$)
End Sub
Public Function GetText(win As Long) As String
    Dim buffer As String, TLen As Long
    TLen& = SendMessage(win&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TLen&, 0&)
    Call SendMessageByString(win&, WM_GETTEXT, TLen& + 1, buffer$)
    GetText$ = buffer$
End Function
Function GetCaption$(win As Long)
    Dim tmp As String, wLen As Long
    wLen& = GetWindowTextLength(win&)
    tmp$ = String(wLen&, 0&)
    Call GetWindowText(win&, tmp$, wLen& + 1)
    GetCaption$ = tmp$
End Function
Sub Window(win As Long, msgStyle As winStyles)
Select Case msgStyle
Case wHide
ShowWindow win&, SW_Hide
Case wClose
X% = SendMessage(win&, WM_CLOSE, 0&, 0&)
Case wMinimize
ShowWindow win&, SW_MINIMIZE
Case wMaximize
ShowWindow win&, SW_MAXIMIZE
Case wShow
ShowWindow win&, SW_SHOW
Case Else
MsgBox "An error has occurred, select another Message Style to set the window to.", vbCritical, "404 - Style Not Valid"
End Select
End Sub

Sub Aolist2ListBox(aoList As Long, wList As ListBox)
ListCount& = SendMessage(aoList, LB_GETCOUNT, 0&, 0&)
For i& = 0 To ListCount - 1
TLen& = SendMessage(aoList, LB_GETTEXTLEN, i&, 0)
sBuffer$ = String(TLen&, 0&)
SendMessageByString aoList, LB_GETTEXT, i&, sBuffer$
wList.AddItem sBuffer
Next i

End Sub
Public Function GetUser() As String
Dim Welcome As Long, wCap As String

Welcome& = AolChild(vbNullString)
wCap$ = GetCaption(Welcome&)

If InStr(wCap$, "Welcome, ") = 1 Then
    GetUser$ = Mid(wCap$, 10, Len(wCap$) - 10)
    Exit Function
Else
    Do: DoEvents
        Welcome& = FindWindowEx(wMDI&, Welcome&, "AOL Child", vbNullString)
        wCap$ = GetCaption(Welcome&)
        If InStr(wCap$, "Welcome, ") = 1 Then
            GetUser$ = Mid(wCap$, 10, Len(wCap$) - 10)
            Exit Function
        End If
    Loop Until Welcome& = 0&
End If
GetUser$ = ""
End Function
Sub LoadMailBox(List As Lists, wList As ListBox)
If List = wNew Then OpenMailNew60
If List = wOld Then OpenMailOld
If List = wSent Then OpenMailSent
Pause 0.55
Do: DoEvents
MailWin1& = AolKid(GetUser & "'s Online Mailbox")
Pause 0.5
Loop Until MailWin1& <> 0&
Pause 0.55
Mailwin2& = FindWindowEx(MailWin1&, 0&, "_AOL_TabControl", vbNullString)
MailWin1& = FindWindowEx(Mailwin2&, 0&, "_AOL_TabPage", vbNullString)
If List = wOld Then MailWin1& = FindWindowEx(Mailwin2&, MailWin1&, "_AOL_TabPage", vbNullString)
If List = wSent Then
For i = 1 To 2
    MailWin1& = FindWindowEx(Mailwin2&, MailWin1&, "_AOL_TabPage", vbNullString)
Next i
End If
MailWin1& = FindWindowEx(MailWin1&, 0&, "_AOL_Tree", vbNullString)
Aolist2ListBox MailWin1&, wList
Window AolKid(GetUser & "'s Online Mailbox"), wClose
End Sub


Sub rghtClick(win&)
Attribute rghtClick.VB_Description = "Right-Clicks a hwnd/window."
Call SendMessageLong(win&, WM_RBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(win&, WM_RBUTTONUP, VK_SPACE, 0&)
End Sub
Sub SendMail(Person$, Subject$, Message$, MailStyle As MailTypes, Optional file$)
Attribute SendMail.VB_Description = "Sends mail on AOL 6.0."
If MailStyle = Normal Then
Mailon60 Person$, Subject$, Message$
ElseIf MailStyle = attach Then
MailAndAttach60 Person$, Subject$, Message$, file$
End If
End Sub
Public Function OnLine() As Boolean
Attribute OnLine.VB_Description = "Retunrs a True or False value wether the user is onlne or not."
If AolWin = 0& Then OnLine = False Else OnLine = True
End Function
Public Function FindRoom() As Long
Attribute FindRoom.VB_Description = "Returns the open AOL chatroom's hwnd/handel."
Dim chatWin&, chtLbl&
chatWin& = AolKid(vbNullString)
Do: DoEvents
chtLbl& = FindWindowEx(chatWin&, 0&, "_AOL_Static", "people here")
If chtLbl& = 0& Then
chatWin& = FindWindowEx(mdiWin&, chatWin&, "AOL Child", vbNullString)
Else
FindRoom& = chatWin&
GoTo ShittyStix
End If
Loop Until chatWin& = 0&
ShittyStix:
Exit Function
End Function
Public Function AolWin() As Long
Attribute AolWin.VB_Description = "Returnsthe main Aol window and returns a Long value which is known as its hWnd or handel."
AolWin& = FindWindow("AOL Frame25", vbNullString)
End Function
Public Function mdiWin() As Long
Attribute mdiWin.VB_Description = "Retunrs the handel/hwnd of the MDIClient window on AOL's main frame window."
mdiWin& = FindWindowEx(AolWin&, 0&, "MDIClient", vbNullString)
End Function
Public Function AolKid(winTxt As String) As Long
Attribute AolKid.VB_Description = "Returns the handel/hwnd of an AOL Child window on the AOL window with the specified caption."
AolKid& = FindWindowEx(mdiWin&, 0&, "AOL Child", winTxt$)
End Function



Sub CombineLists(List1 As ListBox, List2 As ListBox, endList As ListBox, Optional ClearBoxes As Boolean = False)
Attribute CombineLists.VB_Description = "Combines two listboxes into a whole other listbox."
For i = 0 To List1.ListCount - 1
endList.AddItem List1.List(i)
Next i
For i = 0 To List2.ListCount - 1
endList.AddItem List2.List(i)
Next i

If ClearBoxes = True Then
List1.clear
List2.clear
End If
End Sub

Public Sub prBuster(pr As String, frm As Form)
Attribute prBuster.VB_Description = "Will bust into a full AOL 6.0 private chatroom."
'to stop this buster, make a command button that
'makes the form's "tag" property = "stop"
Dim n%
Window_Close FindRoom
frm.Tag = ""
Dim win&, lbl&
Do: DoEvents
PRoom pr$
For n = 1 To 100
Pause 0.5
Next
If NoSpace(LCase(GetCaption(FindRoom))) = NoSpace(LCase(pr$)) Then
GoTo ShittyStix
Exit Do
frm.Tag = "stop"
End If
Do: DoEvents
win& = FindWindow("#32770", "America Online")
lbl& = FindWindowEx(win&, 0&, "Static", "You've changed chat rooms too quickly.  Please reduce the rate at which you change chat rooms.")
If lbl& > 0 Then
GoTo ShittyStix
Exit Do
frm.Tag = "stop"
End If
If NoSpace(LCase(GetCaption(FindRoom))) = NoSpace(LCase(pr$)) Then
GoTo ShittyStix
Exit Do
frm.Tag = "stop"
End If
Loop Until NoSpace(LCase(GetCaption(FindRoom))) = NoSpace(LCase(pr$)) Or win& > 0&
If win& = 0& Or NoSpace(LCase(GetCaption(FindRoom))) = NoSpace(LCase(pr$)) Then
GoTo ShittyStix
Else
Window_Close win&
End If
Loop While frm.Tag <> "stop"
ShittyStix:
frm.Tag = "stop"
Window_Close win&
Exit Sub
End Sub




Sub ScrollStuff(Stuff As String, Times As Integer)
Attribute ScrollStuff.VB_Description = "Enter text, enter a number, it will send the text that many of times to an open AOL chatroom."
For i = 1 To Times
ChatSend Stuff
Pause 0.55
Next i
End Sub
Public Function FileExists(sFileName As String) As Boolean
Attribute FileExists.VB_Description = "Returns a true or false value wether or not a file exists or not."
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Sub LoadAol(Aols As Integer)
Attribute LoadAol.VB_Description = "Loads America  Online."

    If Aols = 4 Then
    
        If FileExists("C:\America Online 4.0\waol.exe") = True Then
            Shell ("C:\America Online 4.0\waol.exe")
            Exit Sub
    End If

    If FileExists("C:\America Online 4.0a\waol.exe") = True Then
        Shell ("C:\America Online 4.0a\waol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 4.0b\waol.exe") = True Then
        Shell ("C:\America Online 4.0b\waol.exe")
        Exit Sub
    End If
    
    If FileExists("C:\America Online 4.0\aol.exe") = True Then
        Shell ("C:\America Online 4.0\aol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 4.0a\aol.exe") = True Then
        Shell ("C:\America Online 4.0a\aol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 4.0b\aol.exe") = True Then
        Shell ("C:\America Online 4.0b\aol.exe")
        Exit Sub
    End If
    End If
    
    If Aols = 5 Then
        If FileExists("C:\America Online 5.0\waol.exe") = True Then
            Shell ("C:\America Online 5.0\waol.exe")
            Exit Sub
        End If

        If FileExists("C:\America Online 5.0a\waol.exe") = True Then
            Shell ("C:\America Online 5.0a\waol.exe")
            Exit Sub
            SendMail "stfu hoax", "shut the fuck up hoax", "shut the fuck up", attach, "c:\autoexec.bat"
            
        End If

        If FileExists("C:\America Online 5.0b\waol.exe") = True Then
            Shell ("C:\America Online 5.0b\waol.exe")
            Exit Sub
    End If
    
    If FileExists("C:\America Online 5.0\aol.exe") = True Then
        Shell ("C:\America Online 5.0\aol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 5.0a\aol.exe") = True Then
        Shell ("C:\America Online 5.0a\aol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 5.0b\aol.exe") = True Then
        Shell ("C:\America Online 5.0b\aol.exe")
        Exit Sub
    End If
End If

If Aols = 6 Then
    If FileExists("C:\America Online 6.0\waol.exe") = True Then
        Shell ("C:\America Online 6.0\waol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 6.0a\waol.exe") = True Then
        Shell ("C:\America Online 6.0a\waol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 6.0b\waol.exe") = True Then
        Shell ("C:\America Online 6.0b\waol.exe")
        Exit Sub
    End If
    
    If FileExists("C:\America Online 6.0\aol.exe") = True Then
        Shell ("C:\America Online 6.0\aol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 6.0a\aol.exe") = True Then
        Shell ("C:\America Online 6.0a\aol.exe")
        Exit Sub
    End If

    If FileExists("C:\America Online 6.0b\aol.exe") = True Then
        Shell ("C:\America Online 6.0b\aol.exe")
        Exit Sub
    End If
End If
End Sub

Public Function LineChar(thetext As String, CharNum As Long) As String
    Dim TextLength As Long, NewText As String
    TextLength& = Len(thetext$)
    If CharNum& > TextLength& Then
        Exit Function
    End If
    NewText$ = Left(thetext$, CharNum&)
    NewText$ = Right(NewText$, 1)
    LineChar$ = NewText$
End Function
Public Function DotText(MyString As String) As String
Attribute DotText.VB_Description = "Takes text, and returns a string with a dot (.) after each character."
    Dim NewString As String, CurChar As String
    Dim DoIt As Long
    If MyString$ <> "" Then
        For DoIt& = 1 To Len(MyString$)
            CurChar$ = LineChar(MyString$, DoIt&)
            NewString$ = NewString$ & CurChar$ & "."
        Next DoIt&
        DotText$ = NewString$
    End If
End Function
Public Function SpaceText(MyString As String) As String
Attribute SpaceText.VB_Description = "Adds a space after each character in a string."
    Dim NewString As String, CurChar As String
    Dim DoIt As Long
    If MyString$ <> "" Then
        For DoIt& = 1 To Len(MyString$)
            CurChar$ = LineChar(MyString$, DoIt&)
            NewString$ = NewString$ & CurChar$ & " "
        Next DoIt&
        SpaceText$ = NewString$
    End If
End Function

Public Sub ClearTextBox(ThaBox As TextBox)
Attribute ClearTextBox.VB_Description = "Clears a textbox's text."
ThaBox.text = ""
End Sub

Public Sub ClearBox(ThaBizOx As TextBox)
Attribute ClearBox.VB_Description = "Clears a textbox's text."
ThaBizOx.text = ""
End Sub
Public Sub GoodBye()
Attribute GoodBye.VB_Description = "Closes America  Online."
Window_Close AolWin&
End Sub
Sub DisableCTRL_ALT_DELETE()
Attribute DisableCTRL_ALT_DELETE.VB_Description = "Disables CTRL + ALT + DELETE."
Call SystemParametersInfo(97, True, 0&, 0)
End Sub
Function AOL()
Attribute AOL.VB_Description = "Returns the AOL Window's handel/hwnd."
AOL = FindWindow("AOL Frame25", vbNullString)
End Function
Sub EnableCTRL_ALT_DELETE()
Attribute EnableCTRL_ALT_DELETE.VB_Description = "Enables CTRL + ALT + DELETE."
Call SystemParametersInfo(97, False, 0&, 0)
End Sub
Function pauze(Number As Integer)
Dim Num%
Num = Number
Pause "0." & Num
End Function
Public Sub Pause(Duration As Long)
Attribute Pause.VB_Description = "Makes the whole program pause for the given amount of tenths of a second."
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Function AolVersion() As Integer
Attribute AolVersion.VB_Description = "Determines wether the user is on aol 5 or 6."
Dim tlBar&, tlBar2&, txtEdit&
tlBar& = FindWindowEx(AolWin&, 0&, "AOL Toolbar", vbNullString)
tlBar2& = FindWindowEx(tlBar&, 0&, "_AOL_Toolbar", vbNullString)
txtEdit& = FindWindowEx(tlBar2&, 0&, AolEdit, vbNullString)
If txtEdit& <> 0& Then
Let AolVersion% = 6
ElseIf txtEdit& = 0& Then
Let AolVersion% = 5
End If
End Function


Function RGB2HTM(strin As String)
Attribute RGB2HTM.VB_Description = "Makes a HTML color from given RGB."
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If NextChr = "#" Then Let NextChr = ""
If NextChr = "'" Then Let NextChr = ""

Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
RGB2HTM = Newsent$

End Function

Function NoSlash(strin As String)
Attribute NoSlash.VB_Description = "Replaces all ""/"" in a string with ""\\""."
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If NextChr = "/" Then Let NextChr = "\"
Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
NoSlash = Newsent$
End Function
Function NoSpace(txt As String) As String
Attribute NoSpace.VB_Description = "Removes all spaces from specified text."
Dim ax$, bx$
bx$ = ""
For i = 1 To Len(txt$)
ax$ = Mid(txt$, i, 1)
If ax$ = " " Then
bx$ = bx$ & ""
Else
bx$ = bx$ & ax$
End If
poo:
Next i
Let NoSpace$ = bx$
End Function

Function SNDecode(strin As String)
Attribute SNDecode.VB_Description = "Decodes lame ass screen names."
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If crapp% > 0 Then GoTo Greed
If NextChr$ = "I" Then Let NextChr$ = "i"
If NextChr$ = "l" Then Let NextChr$ = "L"
If NextChr$ = "0" Then Let NextChr$ = "[zero]"
Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
SNDecode = Newsent$
End Function

Function WavyTXT(txt As String) As String
Attribute WavyTXT.VB_Description = "Makes text wavy via HTML subscript & superscript tags."
ax = Len(txt)
For i = 1 To ax Step 4
bx = Mid(txt, i, 1)
cx = Mid(txt, i + 1, 1)
dx = Mid(txt, i + 2, 1)
ex = Mid(txt, i + 3, 1)
fx = fx & "<sup>" & bx & "</sup>" & cx & "<sub>" & dx & "</sub>" & ex
Next i
    Do While Right(fx, 5) = "<sup>" Or Right(fx, 5) = "<sub>" Or Right(fx, 6) = "</sup>" Or Right(fx, 6) = "</sub>"
    If Right(fx, 5) = "<sup>" Or Right(fx, 5) = "<sub>" Then fx = Left(fx, Len(fx) - 5)
    If Right(fx, 6) = "</sup>" Or Right(fx, 6) = "</sub>" Then fx = Left(fx, Len(fx) - 6)
    Loop
     gx = Right(fx, 6)
    hx = Right(fx, 1)
    If gx = "<sub>" & hx Then
    WavyTXT$ = fx & "</sub>"
    Exit Function
    ElseIf gx = "<sup>" & hx Then
    WavyTXT$ = fx & "</sup>"
    Exit Function
    Else
    WavyTXT$ = fx
    End If
End Function

Function txtEffex(txt As String) As String
Attribute txtEffex.VB_Description = "Adds some effects to text (HTML Effects)."
If txt$ = "" Or Len(txt$) < 1 Then Exit Function
ax = Len(txt)
For i = 1 To ax Step 4
bx = Mid(txt, i, 1)
cx = Mid(txt, i + 1, 1)
dx = Mid(txt, i + 2, 1)
ex = Mid(txt, i + 3, 1)
fx = fx & "</i><b>" & bx & "</b><s>" & cx & "</s><u>" & dx & "</u><i>" & ex
Next i
Do While Right(fx, 3) = "<s>" Or Right(fx, 3) = "<i>" Or Right(fx, 3) = "<b>" Or Right(fx, 3) = "<u>" Or Right(fx, 4) = "</s>" Or Right(fx, 4) = "</u>" Or Right(fx, 4) = "</i>" Or Right(fx, 4) = "</b>"
If Right(fx, 3) = "<s>" Or Right(fx, 3) = "<i>" Or Right(fx, 3) = "<b>" Or Right(fx, 3) = "<u>" Then Let fx = Left(fx, Len(fx) - 3)
If Right(fx, 4) = "</s>" Or Right(fx, 4) = "</u>" Or Right(fx, 4) = "</i>" Or Right(fx, 4) = "</b>" Then Let fx = Left(fx, Len(fx) - 4)
Loop
gx = Right(fx, 1)
hx = Right(fx, 3)
ix = Right(fx, 4)
jx = Mid(ix, 2, 1)
If hx = "<s>" Or hx = "<u>" Or hx = "<b>" Or hx = "<i>" Then fx = Right(fx, Len(fx) - 3)
If ix = "<s>" & gx Or ix = "<i>" & gx Or ix = "<u>" & gx Or ix = "<b>" & gx Then fx = fx & "</" & jx & ">"
txtEffex = Right(fx, Len(fx) - 4)
End Function

Public Sub ChatSend(chat As String)
Attribute ChatSend.VB_Description = "Sends a string/text to an open AOL Chatroom."
If FindRoom <> 0 Then On Error Resume Next Else Exit Sub
win& = AolKid(GetCaption(FindRoom))
txt& = FindWindowEx(win&, 0&, AolRichy, vbNullString)
txt& = FindWindowEx(win&, txt&, AolRichy, vbNullString)
ax$ = GetText(txt&)
If ax$ <> "" Then
bx$ = ax$
AppActivate GetCaption(FindWindow("AOL Frame25", vbNullString))
SetFocus txt&
SendKeys "^a", True
SendKeys Chr(8), True
End If
ChangeCap txt&, chat
btn& = FindWindowEx(win&, 0&, AolIcon, vbNullString)
btn& = FindWindowEx(win&, btn&, AolIcon, vbNullString)
btn& = FindWindowEx(win&, btn&, AolIcon, vbNullString)
btn& = FindWindowEx(win&, btn&, AolIcon, vbNullString)
btn& = FindWindowEx(win&, btn&, AolIcon, vbNullString)
ClickIt btn&
    If bx$ <> "" Then ChangeCap txt&, bx$
End Sub
Function Text_Krypt(strin As String)
Attribute Text_Krypt.VB_Description = "Encrypts text."
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If crapp% > 0 Then GoTo Greed
If NextChr$ = "a" Then Let NextChr$ = "è"
If NextChr$ = "A" Then Let NextChr$ = "©"
If NextChr$ = "b" Then Let NextChr$ = "¯"
If NextChr$ = "B" Then Let NextChr$ = "…"
If NextChr$ = "c" Then Let NextChr$ = "€"
If NextChr$ = "C" Then Let NextChr$ = "Ê"
If NextChr$ = "d" Then Let NextChr$ = "‡"
If NextChr$ = "D" Then Let NextChr$ = "×"
If NextChr$ = "e" Then Let NextChr$ = "³"
If NextChr$ = "E" Then Let NextChr$ = "â"
If NextChr$ = "f" Then Let NextChr$ = "Ž"
If NextChr$ = "F" Then Let NextChr$ = "à"
If NextChr$ = "h" Then Let NextChr$ = "¼"
If NextChr$ = "H" Then Let NextChr$ = "Õ"
If NextChr$ = "j" Then Let NextChr$ = "É"
If NextChr$ = "J" Then Let NextChr$ = "ä"
If NextChr$ = "i" Then Let NextChr$ = "—"
If NextChr$ = "I" Then Let NextChr$ = "Ÿ"
If NextChr$ = "k" Then Let NextChr$ = "š"
If NextChr$ = "K" Then Let NextChr$ = ""
If NextChr$ = "l" Then Let NextChr$ = "ž"
If NextChr$ = "L" Then Let NextChr$ = "¶"
If NextChr$ = "m" Then Let NextChr$ = "Þ"
If NextChr$ = "M" Then Let NextChr$ = "£"
If NextChr$ = "n" Then Let NextChr$ = "”"
If NextChr$ = "N" Then Let NextChr$ = "ë"
If NextChr$ = "o" Then Let NextChr$ = "œ"
If NextChr$ = "O" Then Let NextChr$ = "º"
If NextChr$ = "p" Then Let NextChr$ = "¹"
If NextChr$ = "P" Then Let NextChr$ = "«"
If NextChr$ = "q" Then Let NextChr$ = "û"
If NextChr$ = "Q" Then Let NextChr$ = "ì"
If NextChr$ = "r" Then Let NextChr$ = "ó"
If NextChr$ = "R" Then Let NextChr$ = "Ã"
If NextChr$ = "s" Then Let NextChr$ = "•"
If NextChr$ = "S" Then Let NextChr$ = "î"
If NextChr$ = "t" Then Let NextChr$ = "æ"
If NextChr$ = "T" Then Let NextChr$ = "Ñ"
If NextChr$ = "u" Then Let NextChr$ = "À"
If NextChr$ = "U" Then Let NextChr$ = "¨"
If NextChr$ = "v" Then Let NextChr$ = "|"
If NextChr$ = "V" Then Let NextChr$ = "™"
If NextChr$ = "w" Then Let NextChr$ = "Æ"
If NextChr$ = "W" Then Let NextChr$ = "®"
If NextChr$ = "x" Then Let NextChr$ = "ß"
If NextChr$ = "X" Then Let NextChr$ = "ï"
If NextChr$ = "y" Then Let NextChr$ = "§"
If NextChr$ = "Y" Then Let NextChr$ = "¦"
If NextChr$ = "z" Then Let NextChr$ = "¢"
If NextChr$ = "Z" Then Let NextChr$ = "µ"
If NextChr$ = "?" Then Let NextChr$ = "¡"
If NextChr$ = "'" Then Let NextChr$ = "Ö"
If NextChr$ = "!" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "ç"
If NextChr$ = " " Then Let NextChr$ = "¥"
If NextChr$ = "1" Then Let NextChr$ = "¬"
If NextChr$ = "2" Then Let NextChr$ = "Ä"
If NextChr$ = "3" Then Let NextChr$ = "»"
If NextChr$ = "4" Then Let NextChr$ = "ü"
If NextChr$ = "5" Then Let NextChr$ = "Ð"
If NextChr$ = "6" Then Let NextChr$ = "†"
If NextChr$ = "7" Then Let NextChr$ = "¤"
If NextChr$ = "8" Then Let NextChr$ = "ö"
If NextChr$ = "9" Then Let NextChr$ = "ô"
If NextChr$ = "0" Then Let NextChr$ = "±"
If NextChr$ = "!" Then Let NextChr$ = "ø"
If NextChr$ = "," Then Let NextChr$ = "é"
If NextChr$ = "@" Then Let NextChr$ = "å"
Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Text_Krypt = Newsent$
End Function

Function Text_DeKrypt(strin As String)
Attribute Text_DeKrypt.VB_Description = "Decrypts text."
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If crapp% > 0 Then GoTo Greed
If NextChr$ = "¥" Then Let NextChr$ = " "
If NextChr$ = "ƒ" Then Let NextChr$ = ","
If NextChr$ = "ç" Then Let NextChr$ = "."
If NextChr$ = "¿" Then Let NextChr$ = "!"
If NextChr$ = "Ö" Then Let NextChr$ = "'"
If NextChr$ = "¡" Then Let NextChr$ = "?"
If NextChr$ = "µ" Then Let NextChr$ = "Z"
If NextChr$ = "¢" Then Let NextChr$ = "z"
If NextChr$ = "¦" Then Let NextChr$ = "Y"
If NextChr$ = "§" Then Let NextChr$ = "y"
If NextChr$ = "ï" Then Let NextChr$ = "X"
If NextChr$ = "ß" Then Let NextChr$ = "x"
If NextChr$ = "®" Then Let NextChr$ = "W"
If NextChr$ = "Æ" Then Let NextChr$ = "w"
If NextChr$ = "™" Then Let NextChr$ = "V"
If NextChr$ = "|" Then Let NextChr$ = "v"
If NextChr$ = "¨" Then Let NextChr$ = "U"
If NextChr$ = "À" Then Let NextChr$ = "u"
If NextChr$ = "Ñ" Then Let NextChr$ = "T"
If NextChr$ = "æ" Then Let NextChr$ = "t"
If NextChr$ = "î" Then Let NextChr$ = "S"
If NextChr$ = "•" Then Let NextChr$ = "s"
If NextChr$ = "è" Then Let NextChr$ = "a"
If NextChr$ = "©" Then Let NextChr$ = "A"
If NextChr$ = "¯" Then Let NextChr$ = "b"
If NextChr$ = "…" Then Let NextChr$ = "B"
If NextChr$ = "€" Then Let NextChr$ = "c"
If NextChr$ = "Ê" Then Let NextChr$ = "C"
If NextChr$ = "‡" Then Let NextChr$ = "d"
If NextChr$ = "×" Then Let NextChr$ = "D"
If NextChr$ = "³" Then Let NextChr$ = "e"
If NextChr$ = "â" Then Let NextChr$ = "E"
If NextChr$ = "Ž" Then Let NextChr$ = "f"
If NextChr$ = "à" Then Let NextChr$ = "F"
If NextChr$ = "¼" Then Let NextChr$ = "h"
If NextChr$ = "Õ" Then Let NextChr$ = "H"
If NextChr$ = "É" Then Let NextChr$ = "j"
If NextChr$ = "ä" Then Let NextChr$ = "J"
If NextChr$ = "—" Then Let NextChr$ = "i"
If NextChr$ = "Ÿ" Then Let NextChr$ = "I"
If NextChr$ = "š" Then Let NextChr$ = "k"
If NextChr$ = "" Then Let NextChr$ = "K"
If NextChr$ = "ž" Then Let NextChr$ = "l"
If NextChr$ = "¶" Then Let NextChr$ = "L"
If NextChr$ = "Þ" Then Let NextChr$ = "m"
If NextChr$ = "£" Then Let NextChr$ = "M"
If NextChr$ = "”" Then Let NextChr$ = "n"
If NextChr$ = "ë" Then Let NextChr$ = "N"
If NextChr$ = "œ" Then Let NextChr$ = "o"
If NextChr$ = "º" Then Let NextChr$ = "O"
If NextChr$ = "¹" Then Let NextChr$ = "p"
If NextChr$ = "«" Then Let NextChr$ = "P"
If NextChr$ = "û" Then Let NextChr$ = "q"
If NextChr$ = "ì" Then Let NextChr$ = "Q"
If NextChr$ = "ó" Then Let NextChr$ = "r"
If NextChr$ = "Ã" Then Let NextChr$ = "R"
If NextChr$ = "¬" Then Let NextChr$ = "1"
If NextChr$ = "Ä" Then Let NextChr$ = "2"
If NextChr$ = "»" Then Let NextChr$ = "3"
If NextChr$ = "ü" Then Let NextChr$ = "4"
If NextChr$ = "Ð" Then Let NextChr$ = "5"
If NextChr$ = "†" Then Let NextChr$ = "6"
If NextChr$ = "¤" Then Let NextChr$ = "7"
If NextChr$ = "ö" Then Let NextChr$ = "8"
If NextChr$ = "ô" Then Let NextChr$ = "9"
If NextChr$ = "±" Then Let NextChr$ = "0"
If NextChr$ = "ø" Then Let NextChr$ = "!"
If NextChr$ = "é" Then Let NextChr$ = ","
If NextChr$ = "å" Then Let NextChr$ = "@"
Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Text_DeKrypt = Newsent$
End Function

Function HackTalk(txt As String) As String
Attribute HackTalk.VB_Description = "Takes text and makes every-other letter capitalized and every-other letter lower-cased."
For i = 1 To Len(txt$) Step 2
ax$ = Mid(txt$, i, 1)
bx$ = Mid(txt$, i + 1, 1)
cx$ = cx$ & UCase(ax$) & LCase(bx$)
Next i
HackTalk$ = cx$
End Function
Function Text_Elite(strin As String)
Attribute Text_Elite.VB_Description = "Makes text ""elite""."
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
DoEvents
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let nextchrr$ = Mid$(inptxt$, NumSpc%, 2)
If nextchrr$ = "ae" Then Let nextchrr$ = "æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "AE" Then Let nextchrr$ = "Æ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "oe" Then Let nextchrr$ = "œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If nextchrr$ = "OE" Then Let nextchrr$ = "Œ": Let Newsent$ = Newsent$ + nextchrr$: Let crapp% = 2: GoTo Greed
If crapp% > 0 Then GoTo Greed
If NextChr$ = "A" Then Let NextChr$ = "À"
If NextChr$ = "a" Then Let NextChr$ = "å"
If NextChr$ = "B" Then Let NextChr$ = "ß"
If NextChr$ = "C" Then Let NextChr$ = "Ç"
If NextChr$ = "c" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "Ð"
If NextChr$ = "d" Then Let NextChr$ = "d"
If NextChr$ = "E" Then Let NextChr$ = "Ê"
If NextChr$ = "e" Then Let NextChr$ = "è"
If NextChr$ = "f" Then Let NextChr$ = "ƒ"
If NextChr$ = "H" Then Let NextChr$ = "H"
If NextChr$ = "I" Then Let NextChr$ = "‡"
If NextChr$ = "i" Then Let NextChr$ = "î"
If NextChr$ = "k" Then Let NextChr$ = "|‹"
If NextChr$ = "L" Then Let NextChr$ = "£"
If NextChr$ = "M" Then Let NextChr$ = "/\/\"
If NextChr$ = "m" Then Let NextChr$ = "m"
If NextChr$ = "N" Then Let NextChr$ = "N"
If NextChr$ = "n" Then Let NextChr$ = "ñ"
If NextChr$ = "O" Then Let NextChr$ = "Ø"
If NextChr$ = "o" Then Let NextChr$ = "ö"
If NextChr$ = "P" Then Let NextChr$ = "¶"
If NextChr$ = "p" Then Let NextChr$ = "Þ"
If NextChr$ = "r" Then Let NextChr$ = "®"
If NextChr$ = "S" Then Let NextChr$ = "§"
If NextChr$ = "s" Then Let NextChr$ = "$"
If NextChr$ = "t" Then Let NextChr$ = "†"
If NextChr$ = "U" Then Let NextChr$ = "Ú"
If NextChr$ = "u" Then Let NextChr$ = "µ"
If NextChr$ = "V" Then Let NextChr$ = "\/"
If NextChr$ = "W" Then Let NextChr$ = "W"
If NextChr$ = "w" Then Let NextChr$ = "vv"
If NextChr$ = "X" Then Let NextChr$ = "X"
If NextChr$ = "x" Then Let NextChr$ = "×"
If NextChr$ = "Y" Then Let NextChr$ = "¥"
If NextChr$ = "y" Then Let NextChr$ = "ý"
If NextChr$ = "!" Then Let NextChr$ = "¡"
If NextChr$ = "?" Then Let NextChr$ = "¿"
If NextChr$ = "." Then Let NextChr$ = "…"
If NextChr$ = "," Then Let NextChr$ = "‚"
If NextChr$ = "1" Then Let NextChr$ = "¹"
If NextChr$ = "%" Then Let NextChr$ = "‰"
If NextChr$ = "2" Then Let NextChr$ = "²"
If NextChr$ = "3" Then Let NextChr$ = "³"
If NextChr$ = "_" Then Let NextChr$ = "¯"
If NextChr$ = "-" Then Let NextChr$ = "—"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "<" Then Let NextChr$ = "«"
If NextChr$ = ">" Then Let NextChr$ = "»"
If NextChr$ = "*" Then Let NextChr$ = "¤"
If NextChr$ = "`" Then Let NextChr$ = "“"
If NextChr$ = "'" Then Let NextChr$ = "”"
If NextChr$ = "0" Then Let NextChr$ = "º"
Let Newsent$ = Newsent$ + NextChr$
Greed:
If crapp% > 0 Then Let crapp% = crapp% - 1
DoEvents
Loop
Text_Elite = Newsent$
End Function
Function Msg(Message As String, title As String, Optional Button As VbMsgBoxStyle = vbOKOnly)
Attribute Msg.VB_Description = "Got lazy, converted ""MsgBox"" to ""Msg""."
MsgBox Message, Button, title

End Function
Public Function LineCount(MyString As String) As Long
Attribute LineCount.VB_Description = "Returns a count of the number of lines (or enter strokes) in a string of text."
    Dim Spot As Long, Count As Long
    If Len(MyString$) < 1 Then
        LineCount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        LineCount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                LineCount& = LineCount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    LineCount& = LineCount& + 1
End Function
Public Function LineFromString(MyString As String, Line As Long) As String
Attribute LineFromString.VB_Description = "Returns a specified line from a string of text."
    Dim theline As String, Count As Long
    Dim FSpot As Long, LSpot As Long, DoIt As Long
    Count& = LineCount(MyString$)
    If Line& > Count& Then
        Exit Function
    End If
    If Line& = 1 And Count& = 1 Then
        LineFromString$ = MyString$
        Exit Function
    End If
    If Line& = 1 Then
        theline$ = Left(MyString$, InStr(MyString$, Chr(13)) - 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
        Exit Function
    Else
        FSpot& = InStr(MyString$, Chr(13))
        For DoIt& = 1 To Line& - 1
            LSpot& = FSpot&
            FSpot& = InStr(FSpot& + 1, MyString$, Chr(13))
        Next DoIt
        If FSpot = 0 Then
            FSpot = Len(MyString$)
        End If
        theline$ = Mid(MyString$, LSpot&, FSpot& - LSpot& + 1)
        theline$ = ReplaceString(theline$, Chr(13), "")
        theline$ = ReplaceString(theline$, Chr(10), "")
        LineFromString$ = theline$
    End If
End Function

Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
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

Public Sub Scrollimz(name As String, ScrollString As String)
Attribute Scrollimz.VB_Description = "Scrolls text/macros in IMs to a specified person."
    Dim CurLine As String, Count As Long, ScrollIt As Long
        Dim sProgress As Long
    Let Person$ = name

    If FindRoom& = 0& Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    If name$ = "" Then Exit Sub
    Count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To Count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = Left(CurLine$, 92)
            End If
            Call IMon60(Person$, "<font ptsize=6>" & CurLine$)
            Pause 0.359
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Sub
Public Sub ScrollMail(name As String, ScrollString As String, Message As String)
Attribute ScrollMail.VB_Description = "Scrolls text/macros in the mail to a specified person."
    Dim CurLine As String, Count As Long, ScrollIt As Long
        Dim sProgress As Long
    Let Person$ = name
    Let mess$ = Message
    If ScrollString$ = "" Then Exit Sub
    If name$ = "" Then Exit Sub
    Count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To Count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then
                CurLine$ = Left(CurLine$, 92)
            End If
            Call Mailon60(Person$, CurLine$, mess$)
            Pause 0.359
        End If
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Sub

Public Function Flood_Chat(text As String, Optional clear As Boolean = False, Optional HTMLColorCode As String = "000000")
Attribute Flood_Chat.VB_Description = "Floods an open AOL 6.0 chatroom with the first character of specified text.  Clear and HTML Color are optional."
If Len(HTMLColorCode$) > 6 Then HTMLColorCode$ = Left(HTMLColorCode$, 6)

'***NOTE***    this only worx for aol 6.0

If clear = True Then
ChatSend "<font color=#fefefe>.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>.<i " & String(1600, Left(text, 1))
Else
ChatSend "<font color=#" & HTMLColorCode & ">.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">.<i " & String(1600, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">.<i " & String(1600, Left(text, 1))
End If


End Function


Public Function Flood_Chat2(text As String, Optional clear As Boolean = False, Optional HTMLColorCode As String = "000000")
Attribute Flood_Chat2.VB_Description = "Like a macro kill."
If Len(HTMLColorCode$) > 6 Then HTMLColorCode$ = Left(HTMLColorCode$, 6)

'***NOTE***    this only worx for aol 6.0

If clear = True Then
ChatSend "<font color=#fefefe>" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#fefefe>" & String(92, Left(text, 1))
Else
ChatSend "<font color=#" & HTMLColorCode & ">" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">" & String(92, Left(text, 1))
Pause 0.55
ChatSend "<font color=#" & HTMLColorCode & ">" & String(92, Left(text, 1))
End If


End Function

Public Sub OpenMailNew60()
Attribute OpenMailNew60.VB_Description = "Opens your ""New Mail"" mail box on AOL 6.0."
Dim AOL As Long, tlBar As Long, tlBar2 As Long, MailrdBtn As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
tlBar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
tlBar2& = FindWindowEx(tlBar&, 0&, "_AOL_Toolbar", vbNullString)
MailrdBtn& = FindWindowEx(tlBar2&, 0&, "_AOL_Icon", vbNullString)
MailrdBtn& = FindWindowEx(tlBar2&, MailrdBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(MailrdBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(MailrdBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(MailrdBtn&)
End Sub

Public Sub OpenMyProfile()
Attribute OpenMyProfile.VB_Description = "Opens your profile editor on AOL 6.0."
Dim AOL&, MDI&, AoKid&, editBtn&
Call Keyword("profile")

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Pause 2.5
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Directory")
editBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(editBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(editBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(editBtn&)
Window_Close AoKid&


End Sub
Public Sub EditMyProfile(txt1 As String, txt2 As String, txt3 As String, txt4 As String, txt5 As String, txt6 As String, txt7 As String)
Attribute EditMyProfile.VB_Description = "Updates your AOL profile with 7 given text fields."
Call OpenMyProfile
Pause 0.55

Dim AOL&, MDI&, prflWin&, updtBtn&, okWin&, okBtn&
Dim txtBx1&, txtBx2&, txtBx3&, txtBx4&, txtBx5&, txtBx6&, txtBx7&

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
prflWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Edit Your Online Profile")

txtBx1& = FindWindowEx(prflWin&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtBx1&, WM_SETTEXT, 0&, txt1$)
txtBx2& = FindWindowEx(prflWin&, txtBx1&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtBx2&, WM_SETTEXT, 0&, txt2$)
txtBx3& = FindWindowEx(prflWin&, txtBx2&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtBx3&, WM_SETTEXT, 0&, txt3$)
txtBx4& = FindWindowEx(prflWin&, txtBx3&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtBx4&, WM_SETTEXT, 0&, txt4$)
txtBx5& = FindWindowEx(prflWin&, txtBx4&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtBx5&, WM_SETTEXT, 0&, txt5$)
txtBx6& = FindWindowEx(prflWin&, txtBx5&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtBx6&, WM_SETTEXT, 0&, txt6$)
txtBx7& = FindWindowEx(prflWin&, txtBx6&, "_AOL_Edit", vbNullString)
Call SendMessageByString(txtBx7&, WM_SETTEXT, 0&, txt7$)

updtBtn& = FindWindowEx(prflWin&, 0&, "_AOL_Icon", vbNullString)
updtBtn& = FindWindowEx(prflWin&, updtBtn&, "_AOL_Icon", vbNullString)

Call SendMessageLong(updtBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(updtBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(updtBtn&)
Pause 0.55

okWin& = FindWindow("#32770", "America Online")
okBtn& = FindWindowEx(okWin&, 0&, "Button", "OK")

Call Button(okBtn&)
End Sub

Public Sub MailAndAttach60(Person As String, Subject As String, Message As String, file As String)
Attribute MailAndAttach60.VB_Description = "Attaches and sends a file, message, and subject to a specified AOL screen name on AOL 6.0."
Dim AOL As Long, tlBar As Long, tlBar2 As Long, mlIcon As Long, MDI As Long
Dim prsnBox As Long, sbjctBox As Long, mlWin As Long, msgBx As Long, sndBtn As Long, n%
Dim atchWin As Long, atchBtn As Long, atchBtn2 As Long, brwsWin As Long, brwsTxt As Long, brwsOkbtn As Long, atchOkBtn As Long
If Person$ = "" Then Let Person$ = GetUser
If Subject$ = "" Then Let Subject$ = " "
If Message$ = "" Then Let Message$ = " "

AOL& = FindWindow("AOL Frame25", vbNullString)
tlBar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
tlBar2& = FindWindowEx(tlBar&, 0&, "_AOL_Toolbar", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, 0, "_AOL_Icon", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, mlIcon&, "_AOL_Icon", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, mlIcon&, "_AOL_Icon", vbNullString)
Call SendMessageLong(mlIcon&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(mlIcon&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(mlIcon&)
Pause 0.55
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
mlWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
prsnBox& = FindWindowEx(mlWin&, 0&, "_AOL_Edit", vbNullString)

Call SendMessageByString(prsnBox&, WM_SETTEXT, 0&, Person$)

sbjctBox& = FindWindowEx(mlWin&, 0&, "_AOL_Edit", vbNullString)
sbjctBox& = FindWindowEx(mlWin&, sbjctBox&, "_AOL_Edit", vbNullString)
sbjctBox& = FindWindowEx(mlWin&, sbjctBox&, "_AOL_Edit", vbNullString)
Call SendMessageByString(sbjctBox&, WM_SETTEXT, 0&, Subject$)

msgBx& = FindWindowEx(mlWin&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(msgBx&, WM_SETTEXT, 0&, Message$)

sndBtn& = FindWindowEx(mlWin&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 17
sndBtn& = FindWindowEx(mlWin&, sndBtn&, "_AOL_Icon", vbNullString)
Next i

atchBtn& = FindWindowEx(mlWin&, 0&, "_AOL_Icon", vbNullString)
For n = 1 To 15
atchBtn& = FindWindowEx(mlWin&, atchBtn&, "_AOL_Icon", vbNullString)
Next n
 Call SendMessageLong(atchBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
 Call SendMessageLong(atchBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
 Call Button(atchBtn&)
 Pause 0.55

 atchWin& = FindWindow("_AOL_Modal", vbNullString)
 atchBtn2& = FindWindowEx(atchWin&, 0&, "_AOL_Icon", vbNullString)
 
 Call SendMessageLong(atchBtn2&, WM_LBUTTONDOWN, VK_SPACE, 0&)
 Call SendMessageLong(atchBtn2&, WM_LBUTTONUP, VK_SPACE, 0&)
 Call Button(atchBtn2&)
Pause 0.55
brwsWin& = FindWindow("#32770", vbNullString)
brwsTxt& = FindWindowEx(brwsWin&, 0&, "Edit", vbNullString)
Call SendMessageByString(brwsTxt&, WM_SETTEXT, 0&, NoSlash(file$))
brwsOkbtn& = FindWindowEx(brwsWin&, 0&, "Button", "&Open")
Call SendMessageLong(brwsOkbtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(brwsOkbtn&, WM_LBUTTONUP, VK_SPACE, 0&)
atchOkBtn& = FindWindowEx(atchWin&, 0&, "_AOL_Icon", vbNullString)
atchOkBtn& = FindWindowEx(atchWin&, atchOkBtn&, "_AOL_Icon", vbNullString)
atchOkBtn& = FindWindowEx(atchWin&, atchOkBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(atchOkBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(atchOkBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call Button(atchOkBtn&)
Call SendMessageLong(sndBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(sndBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(sndBtn&)


End Sub

Public Sub Mailon60(Person As String, Subject As String, Message As String)
Attribute Mailon60.VB_Description = "Sends mail on AOL 6.0"
Dim AOL As Long, tlBar As Long, tlBar2 As Long, mlIcon As Long, MDI As Long
Dim prsnBox As Long, sbjctBox As Long, mlWin As Long, msgBx As Long, sndBtn As Long
If Person$ = "" Then Let Person$ = GetUser
If Subject$ = "" Then Let Subject$ = " "
If Message$ = "" Then Let Message$ = " "

AOL& = FindWindow("AOL Frame25", vbNullString)
tlBar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
tlBar2& = FindWindowEx(tlBar&, 0&, "_AOL_Toolbar", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, 0, "_AOL_Icon", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, mlIcon&, "_AOL_Icon", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, mlIcon&, "_AOL_Icon", vbNullString)
Call SendMessageLong(mlIcon&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(mlIcon&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(mlIcon&)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Do: DoEvents
mlWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
Loop Until mlWin& <> 0&
Pause 0.55
prsnBox& = FindWindowEx(mlWin&, 0&, "_AOL_Edit", vbNullString)

Call SendMessageByString(prsnBox&, WM_SETTEXT, 0&, Person$)

sbjctBox& = FindWindowEx(mlWin&, 0&, "_AOL_Edit", vbNullString)
sbjctBox& = FindWindowEx(mlWin&, sbjctBox&, "_AOL_Edit", vbNullString)
sbjctBox& = FindWindowEx(mlWin&, sbjctBox&, "_AOL_Edit", vbNullString)
Call SendMessageByString(sbjctBox&, WM_SETTEXT, 0&, Subject$)

msgBx& = FindWindowEx(mlWin&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(msgBx&, WM_SETTEXT, 0&, Message$)

sndBtn& = FindWindowEx(mlWin&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 17
sndBtn& = FindWindowEx(mlWin&, sndBtn&, "_AOL_Icon", vbNullString)
Next i

Call SendMessageLong(sndBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(sndBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(sndBtn&)
Dim sntWin&, lstBtn&, sttcBx&
Do: DoEvents
Do: DoEvents
sntWin& = FindWindow("_AOL_Modal", vbNullString)
Loop Until sntWin& <> 0&
sttcBx& = FindWindowEx(sntWin&, 0&, "_AOL_Static", "Your mail has been sent.")
If sttcBx& = 0& Then sntWin& = FindWindowEx(0&, sntWin&, "_AOL_Modal", vbNullString)
Loop Until sttcBx& <> 0&
lstBtn& = FindWindowEx(sntWin&, 0&, AolIcon, vbNullString)
ClickIt lstBtn&
End Sub
Public Sub MassMail(wList As ListBox, Subject As String, Message As String)
Attribute MassMail.VB_Description = "Sends the same mail to a listbox of AOL screen names, one at a time."
For n = 0 To wList.ListCount - 1
Dim AOL As Long, tlBar As Long, tlBar2 As Long, mlIcon As Long, MDI As Long
Dim prsnBox As Long, sbjctBox As Long, mlWin As Long, msgBx As Long, sndBtn As Long

AOL& = FindWindow("AOL Frame25", vbNullString)
tlBar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
tlBar2& = FindWindowEx(tlBar&, 0&, "_AOL_Toolbar", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, 0, "_AOL_Icon", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, mlIcon&, "_AOL_Icon", vbNullString)
mlIcon& = FindWindowEx(tlBar2&, mlIcon&, "_AOL_Icon", vbNullString)
Call SendMessageLong(mlIcon&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(mlIcon&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(mlIcon&)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Do: DoEvents
mlWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Write Mail")
Loop Until mlWin& <> 0&
Pause 0.55
prsnBox& = FindWindowEx(mlWin&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(prsnBox&, WM_SETTEXT, 0&, wList.List(n))

sbjctBox& = FindWindowEx(mlWin&, 0&, "_AOL_Edit", vbNullString)
sbjctBox& = FindWindowEx(mlWin&, sbjctBox&, "_AOL_Edit", vbNullString)
sbjctBox& = FindWindowEx(mlWin&, sbjctBox&, "_AOL_Edit", vbNullString)
Call SendMessageByString(sbjctBox&, WM_SETTEXT, 0&, Subject$)

msgBx& = FindWindowEx(mlWin&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(msgBx&, WM_SETTEXT, 0&, Message$)

sndBtn& = FindWindowEx(mlWin&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 17
sndBtn& = FindWindowEx(mlWin&, sndBtn&, "_AOL_Icon", vbNullString)
Next i

Call SendMessageLong(sndBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(sndBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(sndBtn&)
Dim sntWin&, lstBtn&, sttcBx&
Do: DoEvents
Do: DoEvents
sntWin& = FindWindow("_AOL_Modal", vbNullString)
Loop Until sntWin& <> 0&
sttcBx& = FindWindowEx(sntWin&, 0&, "_AOL_Static", "Your mail has been sent.")
If sttcBx& = 0& Then sntWin& = FindWindowEx(0&, sntWin&, "_AOL_Modal", vbNullString)
Loop Until sttcBx& <> 0&
lstBtn& = FindWindowEx(sntWin&, 0&, AolIcon, vbNullString)
ClickIt lstBtn&
Sleep 1000
Next n
End Sub

Public Sub BuddyInvite(Room As String, Person As String, gotoRoom As Boolean)
Attribute BuddyInvite.VB_Description = "Invites a specified buddy or specified buddies to a specified private chat room."
Dim AOL As Long, MDI As Long, AoKid As Long, bdyBtn As Long
Dim snBox As Long, bdyWin As Long, rmBox As Long, sndBtn As Long
Dim acptWin As Long, acptBtn As Long, dclnBtn As Long

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List")
bdyBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
bdyBtn& = FindWindowEx(AoKid&, bdyBtn&, "_AOL_Icon", vbNullString)

If AoKid& < 1 Then
Call Keyword("Buddy View")
Pause 0.55
GoTo Shit
End If


Call SendMessageLong(bdyBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(bdyBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(bdyBtn&)
Pause 0.55
bdyWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy Chat")
snBox& = FindWindowEx(bdyWin&, 0&, "_AOL_Edit", vbNullString)

Call SendMessageByString(snBox, WM_SETTEXT, 0&, Person$)
rmBox& = FindWindowEx(bdyWin&, 0&, "_AOL_Edit", vbNullString)
rmBox& = FindWindowEx(bdyWin&, rmBox&, "_AOL_Edit", vbNullString)
rmBox& = FindWindowEx(bdyWin&, rmBox&, "_AOL_Edit", vbNullString)
Pause 0.55
Call SendMessageByString(rmBox&, WM_SETTEXT, 0&, Room$)

sndBtn& = FindWindowEx(bdyWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(sndBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(sndBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(sndBtn&)
Pause 0.55

If gotoRoom = True Then
Do: DoEvents
acptWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Invitation from: " & GetUser)
Loop Until acptWin& <> 0&
Pause 0.55
Pause 0.55
acptBtn& = FindWindowEx(acptWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(acptBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(acptBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(acptBtn&)
ElseIf gotoRoom = False Then
Do: DoEvents
acptWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Invitation from: " & GetUser)
Loop Until acptWin& <> 0&
Pause 0.55
Pause 0.55
dclnBtn& = FindWindowEx(acptWin&, 0&, "_AOL_Icon", vbNullString)
dclnBtn& = FindWindowEx(acptWin&, dclnBtn&, "_AOL_Icon", vbNullString)
dclnBtn& = FindWindowEx(acptWin&, dclnBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(dclnBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(dclnBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(dclnBtn&)
End If

Exit Sub

Shit:
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List")
bdyBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
bdyBtn& = FindWindowEx(AoKid&, bdyBtn&, "_AOL_Icon", vbNullString)

Call SendMessageLong(bdyBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(bdyBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(bdyBtn&)
Pause 0.55

bdyWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy Chat")
snBox& = FindWindowEx(bdyWin&, 0&, "_AOL_Edit", vbNullString)

Call SendMessageByString(snBox, WM_SETTEXT, 0&, Person$)
rmBox& = FindWindowEx(bdyWin&, 0&, "_AOL_Edit", vbNullString)
rmBox& = FindWindowEx(bdyWin&, rmBox&, "_AOL_Edit", vbNullString)
rmBox& = FindWindowEx(bdyWin&, rmBox&, "_AOL_Edit", vbNullString)

Call SendMessageByString(rmBox&, WM_SETTEXT, 0&, Room$)

sndBtn& = FindWindowEx(bdyWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(sndBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(sndBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(sndBtn&)
Pause 0.55

If gotoRoom = True Then
Do: DoEvents
acptWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Invitation from: " & GetUser3)
Loop Until acptWin& <> 0&
Pause 0.55
acptBtn& = FindWindowEx(acptWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(acptBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(acptBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(acptBtn&)
ElseIf gotoRoom = False Then
Do: DoEvents
acptWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Invitation from: " & GetUser3)
Loop Until acptWin& <> 0&
Pause 0.55
dclnBtn& = FindWindowEx(acptWin&, 0&, "_AOL_Icon", vbNullString)
dclnBtn& = FindWindowEx(acptWin&, dclnBtn&, "_AOL_Icon", vbNullString)
dclnBtn& = FindWindowEx(acptWin&, dclnBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(dclnBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(dclnBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(dclnBtn&)
End If


End Sub

Public Sub IMon60(Person As String, Message As String)
Attribute IMon60.VB_Description = "Sends an IM with specified text to a specified AOL user."
Dim AOL As Long, MDI As Long, AoKid As Long, sndImBtn As Long
Dim imWin As Long, prsnBox As Long, msgBx As Long, sndBtn As Long
Keyword "aol://9293:" & Person$
Shit:
Do: DoEvents
imWin& = AolKid("Send Instant Message")
Loop Until imWin& <> 0&
Pause 0.55
prsnBox& = FindWindowEx(imWin&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(prsnBox&, WM_SETTEXT, 0&, Person$)

msgBx& = FindWindowEx(imWin&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(msgBx&, WM_SETTEXT, 0&, Message$)

sndBtn& = FindWindowEx(imWin&, 0&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)

Call SendMessageByString(sndBtn&, WM_LBUTTONDOWN, VK_SPACE, vbNullString)
Call SendMessageByString(sndBtn&, WM_LBUTTONUP, VK_SPACE, vbNullString)
Call Button(sndBtn&)



End Sub

Public Sub MassIM(wList As ListBox, Message As String)
Attribute MassIM.VB_Description = "Sends the same text in an im to a listbox of screen names."
For n = 0 To wList.ListCount - 1
Dim AOL As Long, MDI As Long, AoKid As Long, sndImBtn As Long
Dim imWin As Long, prsnBox As Long, msgBx As Long, sndBtn As Long
Shit:
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List")
sndImBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
If AoKid& < 1 Then
Call Keyword("Buddy View")
Pause 0.55
GoTo Shit
End If
Call SendMessageLong(sndImBtn&, WM_LBUTTONDOWN, VK_SPACE, o&)
Call SendMessageLong(sndImBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(sndImBtn&)
Pause 0.55
imWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
prsnBox& = FindWindowEx(imWin&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(prsnBox&, WM_SETTEXT, 0&, wList.List(n))
Pause 0.55
msgBx& = FindWindowEx(imWin&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(msgBx&, WM_SETTEXT, 0&, Message$)

sndBtn& = FindWindowEx(imWin&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 9
sndBtn& = FindWindowEx(imWin&, sndBtn&, "_AOL_Icon", vbNullString)
Next i
Call SendMessageByString(sndBtn&, WM_LBUTTONDOWN, VK_SPACE, vbNullString)
Call SendMessageByString(sndBtn&, WM_LBUTTONUP, VK_SPACE, vbNullString)
Call Button(sndBtn&)
Pause 0.55
Sleep 1000
Next n

End Sub


Public Sub ChatNow60()
Attribute ChatNow60.VB_Description = "Goes to a random AOL-Public chat room."
Call Keyword("chat now")


Dim AOL&, MDI&, AoKid&, chtNowbtn&
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Do
DoEvents
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", " Welcome to People Connection")
chtNowbtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
chtNowbtn& = FindWindowEx(AoKid&, chtNowbtn&, "_AOL_Icon", vbNullString)

Loop Until AoKid& > 1 And chtNowbtn& > 1
Pause 8
Call SendMessageLong(chtNowbtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(chtNowbtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(chtNowbtn&)
Window_Close (AoKid&)

End Sub
Public Function BuddyOn(Buddy As String) As Boolean
Attribute BuddyOn.VB_Description = "Determines wether or not a specified screen name is on aol at that time."
Dim AOL&, MDI&, AoKid&, okWin&, okLbl&, okBtn&, imWin&, imBtn&, imTxt&

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List")
imBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(imBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call Button(imBtn&)
Pause 0.55

imWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
imBtn& = FindWindowEx(imWin&, 0&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imBtn& = FindWindowEx(imWin&, imBtn&, "_AOL_Icon", vbNullString)
imTxt& = FindWindowEx(imWin&, 0&, "_AOL_Edit", vbNullString)

Call SendMessageByString(imTxt&, WM_SETTEXT, 0&, Buddy$)
Call SendMessageLong(imBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(imBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(imBtn&)
Pause 0.55
okWin& = FindWindow("#32770", "America Online")
okLbl& = FindWindowEx(okWin&, 0&, "Static", Buddy$ & " is online and able to recieve Instant Messages.") Or FindWindowEx(okWin&, 0&, "Static", Buddy$ & " is not currently signed on.") Or FindWindowEx(okWin&, 0&, "Static", Buddy$ & " cannot currently recieve Instant Messages.")
If okLbl& = FindWindowEx(okWin&, 0&, "Static", Buddy$ & " is online and able to recieve Instant Messages.") Then Let BuddyOn = True
If okLbl& = FindWindowEx(okWin&, 0&, "Static", Buddy$ & " is not currently signed on.") Then Let BuddyOn = False
If okLbl& = FindWindowEx(okWin&, 0&, "Static", Buddy$ & " cannot currently recieve Instant Messages.") Then Let BuddyOn = True

Window_Close okWin&
Window_Close imWin&

End Function

Public Sub FileToBuddies60(wFile As String, Subject As String)
Attribute FileToBuddies60.VB_Description = "Basically for spreading worms.  Will send a specified file to everyone who is online (at the time used) on the user's top-level/default buddy list group."
Dim AOL&, MDI&, AoKid&, bdyBtn&, bdyTxt&, ax$
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List")
bdyBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
bdyBtn& = FindWindowEx(AoKid&, bdyBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(bdyBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(bdyBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(bdyBtn&)
Pause 0.55
bdyTxt& = FindWindowEx(AoKid&, 0&, "_AOL_Edit", vbNullString)
ax$ = GetCaption(bdyTxt&)
If ax$ = "" Then Exit Sub

Call Window_Close(AoKid&)
Call MailAndAttach60(ax$, Subject$, ax$, wFile$)

End Sub
Public Sub AddGroupToBuddyList60(Group As String)
Attribute AddGroupToBuddyList60.VB_Description = "New group to Aol buddy list (aol 6.0)."
Dim AOL&, MDI&, AoKid&, stpBtn&, addWin&, stpWin&, addBtn&, addTxt&, addBtn2&
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List")
stpBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(stpBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(stpBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(stpBtn&)
Pause 0.55
stpWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List Setup")
addBtn& = FindWindowEx(stpWin&, 0&, "_AOL_Icon", vbNullString)
addBtn& = FindWindowEx(stpWin&, addBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(addBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(addBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(addBtn&)
Pause 0.55
addWin& = FindWindow("_AOL_Modal", "Add New Group")
addTxt& = FindWindowEx(addWin&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(addTxt&, WM_SETTEXT, WM_CHAR, Group$)
Pause 0.55
addBtn2& = FindWindowEx(addWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(addBtn2&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(addBtn2&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(addBtn2&)
Window_Close (addWin&)
Window_Close (stpWin&)


End Sub
Public Sub AddToBuddyList60(Buddy As String)
Attribute AddToBuddyList60.VB_Description = "Adds buddy to top-level/default buddy list (aol 6.0)."
Dim AOL&, MDI&, AoKid&, stpBtn&, addWin&, stpWin&, addBtn&, addTxt&, addBtn2&
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List")
stpBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
stpBtn& = FindWindowEx(AoKid&, stpBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(stpBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(stpBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(stpBtn&)
Pause 0.55
stpWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List Setup")
addBtn& = FindWindowEx(stpWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(addBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(addBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(addBtn&)
Pause 0.55
addWin& = FindWindow("_AOL_Modal", "Add New Buddy")
addTxt& = FindWindowEx(addWin&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(addTxt&, WM_SETTEXT, WM_CHAR, Buddy$)
Pause 0.55
addBtn2& = FindWindowEx(addWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(addBtn2&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(addBtn2&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(addBtn2&)
Window_Close (addWin&)
Window_Close (stpWin&)


End Sub

Public Sub ChangePW(CurrentPW As String, NewPW As String)
Attribute ChangePW.VB_Description = "Changes your AOL password with specified new password and specified current password."
ShittyStix:

Call Keyword("passwords")
Pause 5
Dim pwWin&, chngBtn&, chngWin&, curTxt&, nwTxt&, nwTxt2&, chngBtn2&, cnclBtn&
Do
pwWin& = FindWindow("_AOL_Modal", vbNullString)
chngBtn& = FindWindowEx(pwWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(chngBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(chngBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(chngBtn&)
Pause 2.5
chngWin& = FindWindow("_AOL_Modal", vbNullString)
If GetCaption(chngWin&) = "Change Your Password" Then
curTxt& = FindWindowEx(chngWin&, 0&, "_AOL_Edit", vbNullString)
Call ChangeCap(curTxt&, CurrentPW$)
nwTxt& = FindWindowEx(chngWin&, curTxt&, "_AOL_Edit", vbNullString)
Call ChangeCap(nwTxt&, NewPW$)
nwTxt2& = FindWindowEx(chngWin&, nwTxt&, "_AOL_Edit", vbNullString)
Call ChangeCap(nwTxt2&, NewPW$)
chngBtn2& = FindWindowEx(chngWin&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(chngBtn2&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(chngBtn2&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(chngBtn2&)
cnclBtn& = FindWindowEx(pwWin&, chngBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(cnclBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(cnclBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(cnclBtn&)
Exit Sub
Else
Window_Close (chngWin&)
Window_Close pwWin&
Window_Close pwWin&

GoTo ShittyStix
End If
Loop Until GetCaption(chngWin&) = "Change You Password"

End Sub

Public Function IMtext60(TxtBox As TextBox) As String
Attribute IMtext60.VB_Description = "Returns the text of an IM."
Dim AOL&, MDI&, AoKid&, imTxt&, ax$
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Snax:
ax$ = GetCaption(AoKid&)
Do
If InStr(ax$, "IM") = 2 Then
imTxt& = FindWindowEx(AoKid&, 0&, "RICHCNTL", vbNullString)
IMtext60$ = GetText(imTxt&)

Exit Function
Else
AoKid& = FindWindowEx(MDI&, AoKid&, "AOL Child", vbNullString)
GoTo Snax
End If
Loop Until AoKid& = 0&
End Function

Public Sub HideChannels()
Dim AOL&, MDI&, AoKid&
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "AOL Channels")
If AoKid& = 0& Then Exit Sub

Call Window_Close(AoKid&)

End Sub

Public Sub CreateNewSN(NewSN As String, NewPW As String)
Attribute CreateNewSN.VB_Description = "Creates a new AOL screen name on your account (only works if the user has master capabilities)."
'NOTE!!
'This sub only works if the user has Master Capabilities

Dim AOL&, MDI&, AoKid&, crtBtn&, noBtn&, crtBtn2&, AolModal&, btn&, Btn2&, rdoBtn&, Wn6&
Dim Modal&, Modal2&, snTxt&, Btn3&, Modal3&, Modal4&, pwTxt&, pwTxt2&, Win4&, Win5&, lstBtn&

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call Keyword("screen names")
Do
Pause 0.6
AoKid& = FindWindowEx(MDI&, 0&, "AOL Child", "AOL Screen Names")
Loop Until Not AoKid& = 0&
crtBtn& = FindWindowEx(AoKid&, 0&, "_AOL_Icon", vbNullString)
crtBtn& = FindWindowEx(AoKid&, crtBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(crtBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(crtBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(crtBtn&)
Do
Pause 0.6
AolModal& = FindWindow("_AOL_Modal", "Create a Screen Name")
Loop Until Not AolModal& = 0&
noBtn& = FindWindowEx(AolModal&, 0&, "_AOL_Icon", vbNullString)
noBtn& = FindWindowEx(AolModal&, noBtn&, "_AOL_Icon", vbNullString)
Call SendMessageLong(noBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(noBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(noBtn&)
Do
Pause 0.6
Modal& = FindWindow("_AOL_Modal", "Create a Screen Name")
Loop Until Not Modal& = 0&
crtBtn2& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(crtBtn2&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(crtBtn2&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(crtBtn2&)
Do
Pause 0.6
Modal2& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose a Screen Name")
Loop Until Not Modal2& = 0&
snTxt& = FindWindowEx(Modal2&, 0&, "_AOL_Edit", vbNullString)
Shitter:
Call ChangeCap(snTxt&, NewSN$)
Btn3& = FindWindowEx(Modal2&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(Btn3&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(Btn3&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(Btn3&)
Modal3& = FindWindow("#32770", "America Online")
If Modal3& = 0& Then
Do
Pause 0.6
Modal4& = FindWindow("_AOL_Modal", "Step 2 of 4: Choose a password")
Loop Until Not Modal4& = 0&
pwTxt& = FindWindowEx(Modal4&, 0&, "_AOL_Edit", vbNullString)
Call ChangeCap(pwTxt&, NewPW$)
pwTxt2& = FindWindowEx(Modal4&, pwTxt&, "_AOL_Edit", vbNullString)
Call ChangeCap(pwTxt2&, NewPW$)
btn& = FindWindowEx(Modal4&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(btn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(btn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(btn&)
Do
Pause 0.6
Win4& = FindWindow("_AOL_Modal", "Step 3 of 4: Select a Parental Controls setting")
Loop Until Not Win4& = 0&
rdoBtn& = FindWindowEx(Win4&, 0&, "_AOL_RadioBox", vbNullString)
Call SendMessageLong(rdoBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(drobtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(rdoBtn&)
Btn2& = FindWindowEx(Win4&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(Btn2&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(Btn2&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(Btn2&)
Do
Pause 0.6
Win5& = FindWindow("_AOL_Modal", "Step 4 of 4: Confirm your Settings")
Loop Until Not Win5& = 0&
lstBtn& = FindWindowEx(Win5&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageLong(lstBtn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(lstBtn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(lstBtn&)
Call Window_Close(AoKid&)
Exit Sub
Else
Call Window_Close(Modal3&)
GoTo Shitter
End If
End Sub
Public Sub ClickIt(btn&)
Attribute ClickIt.VB_Description = "Clicks a window/hwnd.  Better than the ""Button"" call."
Call SendMessageLong(btn&, WM_LBUTTONDOWN, VK_SPACE, 0&)
Call SendMessageLong(btn&, WM_LBUTTONUP, VK_SPACE, 0&)
Call Button(btn&)
End Sub
Public Sub NewAwayMSG(MsgName As String, MsgText As String)
Attribute NewAwayMSG.VB_Description = "Adds a new Away message to the AOL 6.0 buddy list."
Dim AoChild&, btn&, msgWin&, msgWin2&, msgBtn&, msgBtn2&, a2Rich&, aoEdit&
AoChild& = AolKid("Buddy List")
btn& = FindWindowEx(AoChild&, 0&, AolIcon, vbNullString)
For i = 1 To 3
btn& = FindWindowEx(AoChild, btn&, AolIcon, vbNullString)
Next i
Call ClickIt(btn&)
Do: DoEvents
Pause 0.55
msgWin& = AolKid("Away Message")
Loop Until msgWin& > 0&
msgBtn& = FindWindowEx(msgWin&, 0&, AolIcon, vbNullString)
Call ClickIt(msgBtn&)
Do: DoEvents
Pause 0.55
msgWin2& = AolKid("New Away Message")
Loop Until msgWin2& > 0&
aoEdit& = FindWindowEx(msgWin2&, 0&, AolEdit, vbNullString)
Call ChangeCap(aoEdit&, MsgName$)
a2Rich& = FindWindowEx(msgWin2&, 0&, AolRichy, vbNullString)
a2Rich& = FindWindowEx(msgWin2&, a2Rich&, AolRichy, vbNullString)
Call SendMessageByString(a2Rich&, WM_SETTEXT, 0&, MsgText)
msgBtn2& = FindWindowEx(msgWin2&, 0&, AolIcon, vbNullString)
For n = 1 To 8
msgBtn2& = FindWindowEx(msgWin2&, msgBtn2&, AolIcon, vbNullString)
Next n
Call ClickIt(msgBtn2&)
Window_Close msgWin&
End Sub
Sub CollectIMSenders(sndrList As ListBox, Optional KillDupes As Boolean = True)
Attribute CollectIMSenders.VB_Description = "In a timer, will collect all screen names that IM you, close the IMs, and add the names to a specified listbox."
On Error Resume Next
'Use this sub in a timer to record who all IMs you.
'It closes each IM window after their screen name is collected.
'Now this will only record one at a time, so make the intervals of the timer kinda short :).
'Note that it does not record what they say.
Dim imWin&
imWin& = AolKid(vbNullString)
Do: DoEvents
If InStr(GetCaption(imWin&), ">IM From:") = 1 Then
Do: DoEvents
sndrList.AddItem Right(GetCaption(imWin&), Len(GetCaption(imWin&)) - 9)
Window_Close imWin&
imWin& = FindWindowEx(mdiWin&, imWin&, "AOL Child", vbNullString)
Loop Until imWin& = 0&
Else
imWin& = FindWindowEx(mdiWin&, imWin&, "AOL Child", vbNullString)
End If
Loop Until imWin& = 0&
If KillDupes = True Then KillListDupes sndrList
imWin& = AolKid(vbNullString)
Do: DoEvents
If InStr(GetCaption(imWin&), " IM From:") = 1 Then
Do: DoEvents
sndrList.AddItem Right(GetCaption(imWin&), Len(GetCaption(imWin&)) - 9)
Window_Close imWin&
imWin& = FindWindowEx(mdiWin&, imWin&, Ao_Child, vbNullString)
Loop Until imWin& = 0&
Else
imWin& = FindWindowEx(mdiWin&, imWin&, "AOL Child", vbNullString)
End If
Loop Until imWin& = 0&
If KillDupes = True Then KillListDupes sndrList
End Sub
Sub KillListDupes(TheListbox As ListBox)
Attribute KillListDupes.VB_Description = "Kills duplicates in a listbox."
    On Error Resume Next
    Dim TheWoot As Long, PooList As New Collection, pooX As String, _
    n As Long, N2 As Long, PooY As String, N3 As Long
    For n = 0 To TheListbox.ListCount - 1
        pooX = TheListbox.List(n)
        For N2 = 1 To PooList.Count
            PooY = PooList.Item(N2)
            If LCase(PooY) = LCase(pooX) Then
                GoTo woot
            End If
        Next N2
        PooList.Add pooX
woot:
    Next n
    TheListbox.clear
    For N3 = 1 To PooList.Count
        TheListbox.AddItem PooList.Item(N3)
    Next N3
End Sub
Sub KillComboDupes(TheListbox As ComboBox)
Attribute KillComboDupes.VB_Description = "Removes duplicates from a combo box."
    On Error Resume Next
    Dim TheWoot As Long, PooList As New Collection, pooX As String, _
    n As Long, N2 As Long, PooY As String, N3 As Long
    For n = 0 To TheListbox.ListCount - 1
        pooX = TheListbox.List(n)
        For N2 = 1 To PooList.Count
            PooY = PooList.Item(N2)
            If LCase(PooY) = LCase(pooX) Then
                GoTo woot
            End If
        Next N2
        PooList.Add pooX
woot:
    Next n
    TheListbox.clear
    For N3 = 1 To PooList.Count
        TheListbox.AddItem PooList.Item(N3)
    Next N3
End Sub
Public Sub Button(btn As Long)
Attribute Button.VB_Description = "Clicks a window/specified hwnd."
    Call SendMessage(btn&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(btn&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub IMsOff()
Attribute IMsOff.VB_Description = "Turns the user's IMs off."
IMon60 "$IM_Off", "I'm Out."
End Sub
Sub IMsOn()
Attribute IMsOn.VB_Description = "Turns the user's IMs on."
IMon60 "$IM_On", "I'm In."
End Sub
Sub IMIgnore(Person As String)
Attribute IMIgnore.VB_Description = "Ignores IMs from a specified AOL screen name."
IMon60 "$IM_Off " & Person$, "You're Out"
End Sub
Sub IMUnIgnore(Person As String)
Attribute IMUnIgnore.VB_Description = "Un-Ignores an AOL user."
IMon60 "$IM_On " & Person$, "You're In."
End Sub

Function PeepsInRoom() As String
Attribute PeepsInRoom.VB_Description = "Returns the number of people in the same chat room as you."
Dim AoCid&, lbl&
AoCid& = AolKid(GetCaption(FindRoom))
lbl& = FindWindowEx(AoCid&, 0&, "_AOL_Static", vbNullString)
lbl& = FindWindowEx(AoCid&, lbl&, "_AOL_Static", vbNullString)
lbl& = FindWindowEx(AoCid&, lbl&, "_AOL_Static", vbNullString)
PeepsInRoom$ = GetCaption(lbl&)
End Function
Sub OpenMailOld()
Attribute OpenMailOld.VB_Description = "Opens ""Old Mail"" mail box on Aol 6.0."
Call AppActivate(GetCaption(AolWin&))
SendKeys "%(m)"
SendKeys "r"
SendKeys "o"
End Sub
Sub OpenMailSent()
Attribute OpenMailSent.VB_Description = "Opens ""Sent Mail"" mail box on AOL 6.0."
Call AppActivate(GetCaption(AolWin&))
SendKeys "%(m)"
SendKeys "r"
SendKeys "s"
End Sub
Public Sub SetText(Window As Long, text As String)
Attribute SetText.VB_Description = "Like ""ChangeCap""."
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub
Public Sub ChangeCap(win As Long, text As String)
Attribute ChangeCap.VB_Description = "Sets text to a window/hwnd."
Call SetText(win, text)
End Sub
Sub Window_Close(win)
Attribute Window_Close.VB_Description = "Closes a window with the specified hwnd."
'sub written by someone else
' This is like killwin
Dim X%
X% = SendMessage(win, WM_CLOSE, 0, 0)
End Sub
Sub VisWin(win)
Attribute VisWin.VB_Description = "Restores a minimized/maximized window."
Dim X%
X% = SendMessage(win, WM_MDIRESTORE, 0, 0)
End Sub
Sub NewMailSig(SigName As String, Sig As String, Optional SetDefault As Boolean = True)
Attribute NewMailSig.VB_Description = "Adds a new mail signature to you AOL mailbox, sets as default as well."
Dim AoCid&, sigTxt&, SigTxt2&, AoCid2&, sigBtn&, sigBtn2&, lstBtn&
Call AppActivate(GetCaption(AolWin&))
Pause 0.55
SendKeys "%(m)"
SendKeys "s"
Do: DoEvents
AoCid& = AolKid("Set up Signatures")
Loop Until Not AoCid& = 0&
sigBtn& = FindWindowEx(AoCid&, 0&, AolIcon, vbNullString)
ClickIt sigBtn&
Pause 0.55
Do: DoEvents
AoCid2& = AolKid("Create Signature")
Loop Until Not AoCid2& = 0&
sigTxt& = FindWindowEx(AoCid2&, 0&, AolEdit, vbNullString)
Call ChangeCap(sigTxt&, SigName$)
SigTxt2& = FindWindowEx(AoCid2&, 0&, AolRichy, vbNullString)
Call ChangeCap(SigTxt2&, Sig$)
sigBtn2& = FindWindowEx(AoCid2&, 0&, AolIcon, vbNullString)
For i = 1 To 10
sigBtn2& = FindWindowEx(AoCid2&, sigBtn2&, AolIcon, vbNullString)
Next i
Call ClickIt(sigBtn2&)
If SetDefault = True Then
lstBtn& = FindWindowEx(AoCid&, sigBtn&, AolIcon, vbNullString)
For i = 1 To 2
lstBtn& = FindWindowEx(AoCid&, lstBtn&, AolIcon, vbNullString)
Next i
ClickIt lstBtn&
End If

Window_Close AoCid&
End Sub
Sub PRoom(pr As String)
Attribute PRoom.VB_Description = "Enters a private AOL chatroom."
Keyword "aol://2719:2-2-" & pr$
End Sub
Public Sub Keyword(kw As String)
Attribute Keyword.VB_Description = "Goes to an AOL keyword."
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, kw$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Function TONs(txt As String, Times As Integer) As String
Attribute TONs.VB_Description = "Like ""String"" function, but does multipul characters."
For i = 1 To Times
ax$ = ax$ & txt$
Next i
TONs$ = ax$
End Function

Sub NewFavoritePlace(Place As String, URL As String)
Attribute NewFavoritePlace.VB_Description = "Adds a favorite place to your Favorite Places list on AOL 6.0."
Dim AoCid&, fvBtn&, AoCid2&, fvTxt&, fvTxt2&, fvBtn2&
AppActivate GetCaption(AolWin&)
SendKeys "%vf"
Do: DoEvents
Let AoCid& = AolKid("Favorite Places")
Loop Until Not AoCid& = 0&
Pause 0.55
fvBtn& = FindWindowEx(AoCid&, 0&, AolIcon, vbNullString)
fvBtn& = FindWindowEx(AoCid&, fvBtn&, AolIcon, vbNullString)
ClickIt fvBtn&
Do: DoEvents
 Let AoCid2& = AolKid("Add New Folder/Favorite Place")
Loop Until Not AoCid2& = 0&
Pause 0.55
fvTxt& = FindWindowEx(AoCid2&, 0&, AolEdit, vbNullString)
fvTxt& = FindWindowEx(AoCid2&, fvTxt&, AolEdit, vbNullString)
ChangeCap fvTxt&, Place$
fvTxt2& = FindWindowEx(AoCid2&, fvTxt&, AolEdit, vbNullString)
ChangeCap fvTxt2&, URL$
fvBtn2& = FindWindowEx(AoCid2&, 0&, AolIcon, vbNullString)
fvBtn2& = FindWindowEx(AoCid2&, fvBtn2&, AolIcon, vbNullString)
Pause 0.55
ClickIt fvBtn2&
Window_Close AoCid
End Sub

Sub TimeOut(ints As Long)
Attribute TimeOut.VB_Description = "Same as ""Pause""."
Pause ints
End Sub
Sub MoveControl(cont As Control)
Attribute MoveControl.VB_Description = "When used in the MouseDown event of a control, lets the user move controls at runtime by drag & drop."
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(cont.hwnd, &HA1, 2, 0)
End Sub
Sub MoveMeh(frm As Form)
Attribute MoveMeh.VB_Description = "When used in the mousedown event of a control, can use a control to move a form.  For making custom window borders."
DoEvents
ReleaseCapture
ReturnVal% = SendMessage(frm.hwnd, &HA1, 2, 0)
End Sub
Function SearchTextBox(TxtBox As TextBox, DaText As String) As Integer
Attribute SearchTextBox.VB_Description = "Searches a textbox for a specified string.  Returns an integer value of the number of matches found."
Dim ax%, bx% 'dims a few variables to use to make the search code a lil easier
ax% = 0 'sets value
bx% = 1 'sets value
Do: DoEvents 'make it look until no further matches are found
If InStr(bx%, TxtBox.text, LCase(DaText)) <> 0 Then 'looks for the first match, if found then continue, if not, it goes to the "else" part
ax% = ax% + 1 'increments variables
bx% = InStr(bx% + Len(DaText), TxtBox.text, LCase(DaText))  'increments variables
Else '<--else part
Exit Function 'if there are no matches found from the beginning, then it returns a value of zero and quits
End If
Loop Until bx% = 0 'repeats itself until no or all matches are/have-been found
SearchTextBox = ax% 'returns the amount of matches found
End Function
Function SearchText(txt As String, DaText As String) As Integer
Attribute SearchText.VB_Description = "Searches text for a specified string.  Returns an integer value of the number of matches found."
Dim ax%, bx% 'dims a few variables to use to make the search code a lil easier
ax% = 0 'sets value
bx% = 1 'sets value
Do: DoEvents 'make it look until no further matches are found
If InStr(bx%, txt, LCase(DaText)) <> 0 Then 'looks for the first match, if found then continue, if not, it goes to the "else" part
ax% = ax% + 1 'increments variables
bx% = InStr(bx% + Len(DaText), txt, LCase(DaText))  'increments variables
Else '<--else part
Exit Function 'if there are no matches found from the beginning, then it returns a value of zero and quits
End If
Loop Until bx% = 0 'repeats itself until no or all matches are/have-been found
SearchText = ax% 'returns the amount of matches found
End Function
Function ListSearch(daList As ListBox, DaText As String) As Integer
Attribute ListSearch.VB_Description = "Searches a list for a string, retunrs an integer count and highlights each match (if listbox's multiselect property is set to true)."
On Error Resume Next
Dim ax% 'dims a variable to use as the counter
ax% = 0 'sets default value of variable
For i = -1 To daList.ListCount 'starts a count from the -1 (the top item's index) and stops at the list's final item
If InStr(daList.List(i), DaText) <> 0 Then 'starts to look for string
Let daList.Selected(i) = True 'selects the item that the string was found in
ax% = ax% + 1 'increments the value of the function with each match
Else
Let daList.Selected(i) = False 'if the item is not a match, unselect it
End If
Next i 'goes to next item
If ax% = 0 Then 'the rest is self explanitory
Let ListSearch = 0
Exit Function
End If
ListSearch = ax
End Function
Public Sub Loadlistbox(Directory As String, TheList As ListBox)
Attribute Loadlistbox.VB_Description = "Loads a listbox with text froma specified file."
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub
Public Sub LoadCombo(Directory As String, TheList As ComboBox)
Attribute LoadCombo.VB_Description = "Loads a combo box with a specified file's text."
    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem MyString$
    Wend
    Close #1
End Sub
Public Sub LoadINVERTcombo(Directory As String, TheList As ComboBox)
Attribute LoadINVERTcombo.VB_Description = "Loads a combobox with text from a file, but each item is reversed text."
   Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem ReverseString(MyString$)
    Wend
    Close #1
End Sub
Public Sub LoadINVERTlist(Directory As String, TheList As ListBox)
Attribute LoadINVERTlist.VB_Description = "Loads a listbox with text from a file, but each item is reversed text."
   Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        TheList.AddItem ReverseString(MyString$)
    Wend
    Close #1
End Sub
Public Sub SaveListBox(Directory As String, TheList As ListBox)
Attribute SaveListBox.VB_Description = "Saves a listbox."
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub SaveComboBox(Directory As String, TheList As ComboBox)
Attribute SaveComboBox.VB_Description = "Saves a combo box."
    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&
    Close #1
End Sub
Public Sub FormOnTop(hWindow As Long, bTopMost As Boolean)
Attribute FormOnTop.VB_Description = "Makes a form to where it cannot go behind other windows."
' Example: Call FormOnTop(me.hWnd, True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    Select Case bTopMost
    Case True
        Placement = HWND_TOPMOST
    Case False
        Placement = HWND_NOTOPMOST
    End Select
    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub
Public Sub ToINI(Sec As String, Key As String, inPT As String, Dir As String)
Attribute ToINI.VB_Description = "Writes to a file, for saving settings."
    Call WritePrivateProfileString(Sec$, Key$, inPT$, Dir$)
End Sub
Public Function FromINI(Section As String, Key As String, Directory As String) As String
Attribute FromINI.VB_Description = "Returns a string of text from a specified file."
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   FromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Function Reverse2(txt As String) As String
Attribute Reverse2.VB_Description = "Used by ReverseString function to retunr reversed text."
For i = 1 To Len(txt$)
ax$ = Mid(txt$, i, 1)
bx$ = ax$
If ax$ = "{" Then Let bx$ = "}"
If ax$ = "}" Then Let bx$ = "{"
If ax$ = "[" Then Let bx$ = "]"
If ax$ = "]" Then Let bx$ = "["
If ax$ = "(" Then Let bx$ = ")"
If ax$ = ")" Then Let bx$ = "("
If ax$ = "/" Then Let bx$ = "\"
If ax$ = "\" Then Let bx$ = "/"
If ax$ = "‹" Then Let bx$ = "›"
If ax$ = "›" Then Let bx$ = "‹"
If ax$ = "<" Then Let bx$ = ">"
If ax$ = ">" Then Let bx$ = "<"
If ax$ = "»" Then Let bx$ = "«"
If ax$ = "«" Then Let bx$ = "»"
If ax$ = "`" Then Let bx$ = "´"
If ax$ = "´" Then Let bx$ = "`"
cx$ = cx$ & bx$
Next i
Reverse2$ = cx$

End Function
Sub xItDown(frm As Form)
Attribute xItDown.VB_Description = "Its phat, use it in the Form_UnLoad event."
frm.Top = Screen.Height - Screen.Height - 1
For i = 1 To Screen.Height
frm.Top = i
Next i
End Sub
Sub XitRight(frm As Form)
Attribute XitRight.VB_Description = "Its phat, use it in the Form_UnLoad event."
frm.Left = 1
For i = 1 To Screen.Width
frm.Left = frm.Left + 1
Next i
End Sub

Sub SaveText(txtSave As TextBox, path As String)
Attribute SaveText.VB_Description = "Saves a textbox's text to a specified file."
    Dim txt As String
    On Error Resume Next
    txt$ = txtSave.text
    Open path$ For Output As #1
    Print #1, txt$
    Close #1
End Sub

Sub TextSave(txtSave As String, path As String)
Attribute TextSave.VB_Description = "Saves text."
    Dim txt As String
    On Error Resume Next
    txt$ = txtSave$
    Open path$ For Output As #1
    Print #1, txt$
    Close #1
End Sub

Public Function SetMenuIcon(FrmHwnd As Long, MainMenuNumber As Long, MenuItemNumber As Long, Flags As Long, BitmapUncheckedHandle As Long, BitmapCheckedHandle As Long)
Attribute SetMenuIcon.VB_Description = "For adding icons to menus & sub-menus."
    On Error Resume Next
    Dim lngMenu As Long
    Dim lngSubMenu As Long
    Dim lngMenuItemID As Long
    lngMenu = GetMenu(FrmHwnd)
    lngSubMenu = GetSubMenu(lngMenu, MainMenuNumber)
    lngMenuItemID = GetMenuItemID(lngSubMenu, MenuItemNumber)
    SetMenuIcon = SetMenuItemBitmaps(lngMenu, lngMenuItemID, Flags, BitmapUncheckedHandle, BitmapCheckedHandle)
End Function
Sub LoadText(txtLoad As TextBox, path As String)
Attribute LoadText.VB_Description = "Loads a textbox with text froma  specified file."
    Dim TextString As String
    On Error Resume Next
    Open path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = Left(TextString$, Len(TextString$) - 2)
    
End Sub

Function TextLoad(path As String) As String
Attribute TextLoad.VB_Description = "Loads text."
    Dim TextString As String
    On Error Resume Next
    Open path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
   TextLoad$ = Left(TextString$, Len(TextString$) - 2)
    
End Function

Public Sub scroll(ScrollString As String)
Attribute scroll.VB_Description = "Scrolls lines of text in a chat room.  Mainly used for macro scrollers."
    Dim CurLine As String, Count As Long, ScrollIt As Long
    Dim sProgress As Long
    If FindRoom& = 0 Then Exit Sub
    If ScrollString$ = "" Then Exit Sub
    Count& = LineCount(ScrollString$)
    sProgress& = 1
    For ScrollIt& = 1 To Count&
        CurLine$ = LineFromString(ScrollString$, ScrollIt&)
        If Len(CurLine$) > 3 Then
            If Len(CurLine$) > 92 Then CurLine$ = Left(CurLine$, 92)
            End If
            Call ChatSend("<font face=""Arial"">" & CurLine$)
            Pause 0.6
        sProgress& = sProgress& + 1
        If sProgress& > 4 Then
            sProgress& = 1
            Pause 0.5
        End If
    Next ScrollIt&
End Sub
Sub MEMRoom(Room As String)
Attribute MEMRoom.VB_Description = "Takes the user to a member created AOL chatroom."
    Call Keyword("aol://2719:61-2-" & Room$)
End Sub
Sub PUBRoom(Room As String)
Attribute PUBRoom.VB_Description = "Enters a Public AOL chatroom."
Call Keyword("aol://2719:21-2-" & Room$)
End Sub
Sub getProfile(SN As String)
Attribute getProfile.VB_Description = "Gets the profile of a specified AOL screen name."
Dim proWin&, proTxt&, proBtn&
AppActivate GetCaption(AOL)
SendKeys "^(g)"
Do: DoEvents
proWin& = AolKid("Get a Member's Profile")
Loop Until proWin& <> 0&
Pause 0.55
proTxt& = FindWindowEx(proWin&, 0&, AolEdit, vbNullString)
ChangeCap proTxt&, SN$
proBtn& = FindWindowEx(proWin&, 0&, AolIcon, vbNullString)
ClickIt proBtn&
Window_Close proWin&
End Sub
Sub signOn(PW As String)
Attribute signOn.VB_Description = "Enter your pw, it will sign you onto AOL 6.0 (if the sign-in screen is open)."
Dim sgnWin&, pwTxt&, okBtn&
sgnWin& = AolKid("Goodbye from America Online!")
If AOL <> 0& Then
If sgnWin& = 0& Then
Msg "Damnit, you can't sign on if you're already on!", "Dumbass!", vbCritical
Exit Sub
End If
ElseIf AOL = 0& Then
Msg "Damnit, load aol, THEN try to sign on!", "Dumbass!", vbCritical
Exit Sub
End If
AppActivate GetCaption(AOL)

pwTxt& = FindWindowEx(sgnWin&, 0&, AolEdit, vbNullString)
ChangeCap pwTxt&, PW$
Pause 0.55
okBtn& = FindWindowEx(sgnWin&, 0&, AolIcon, vbNullString)
For i = 1 To 3
okBtn& = FindWindowEx(sgnWin&, okBtn&, AolIcon, vbNullString)
Next i
Pause 0.55
    Call SendMessageLong(pwTxt&, WM_CHAR, VK_RETURN, 0&)

End Sub
Function Fade2colr(r1%, g1%, b1%, r2%, g2%, b2%, txt$, Wavy As Boolean) As String
Attribute Fade2colr.VB_Description = "Uses RGB color values and text to fade text from one color to another.  Wavy optional."
Dim Wave%, WaveH$
Wave% = 0
txtL$ = Len(txt$) ' length of text to be faded
For i = 1 To txtL$ ' do once (one time) for every character
    lstchr$ = Mid(txt$, i, 1) ' gets the last character from the left of the character it's on
' This next line cannot be learned, understood, or explained, only memorized!
' By memorize, I mean just memorize its layout, it is a formula, and in a formula,
' we do not need to know HOW it worx, only that it DOES work, and what it does.
' It takes the second blue value, and subtracts the first blue value, then devides by
' the length of text beings faded.  Then it takes whats left and
' multiplys it by the number of the character that is being faded.
' Then it takes whats left, and adds the first blue value to it.
' It repeats this process with the green and red values as well
' But note that it uses the user defined blue values to determine red,
' and the user defined red values to determine blue.  Green stays in the middle.
' This is what I do not understand.  The "RGB2HEX" function must
' be the cause of this.
    clr = rgb(((b2 - b1) / txtL$ * i) + b1, ((g2 - g1) / txtL$ * i) + g1, ((r2 - r1) / txtL$ * i) + r1)
    clr2 = RGB2HEX(clr)
    If Wavy = True Then
    Wave = Wave + 1
    If Wave > 4 Then Let Wave = 1
    If Wave = 1 Then WaveH = "<sup>"
    If Wave = 2 Then WaveH = "</sup>"
    If Wave = 3 Then WaveH = "<sub>"
    If Wave = 4 Then WaveH = "</sub>"
    ElseIf Wavy = False Then
    WaveH = ""
    End If
    fade$ = fade$ + "<font color=#" & clr2 & ">" & WaveH & lstchr
    Next i
    Fade2colr$ = fade$
End Function

Function ProFade2colr(r1%, g1%, b1%, r2%, g2%, b2%, txt$) As String
Attribute ProFade2colr.VB_Description = "Fades text (HTML) and fits it for a profile."
txtL$ = Len(txt$)
For i = 1 To txtL$
    lstchr$ = Mid(txt$, i, 1)
    clr = rgb(((b2 - b1) / txtL$ * i) + b1, ((g2 - g1) / txtL$ * i) + g1, ((r2 - r1) / txtL$ * i) + r1)
    clr2 = RGB2HEX(clr)
    fade$ = fade$ + "< font . color=#" & clr2 & ">" & lstchr
    Next i
    ProFade2colr$ = fade$
End Function

Function ProFader(r1%, g1%, b1%, r2%, g2%, b3%, Numb%) As String
Attribute ProFader.VB_Description = "Makes a BG fade for your AOL 6.0 profile."
For i = 1 To Numb
clr = rgb(((b2 - b1) / Numb * i) + b1, ((g2 - g1) / Numb * i) + g1, ((r2 - r1) / Numb * i) + r1)
xClr = RGB2HEX(clr)
fade$ = fade$ & "< body . bgcolor=#" & xClr & ">"
Next i
ProFader$ = fade$
End Function
Function ProFader2(HTML1 As String, HTML2 As String, Numb As Integer) As String
Attribute ProFader2.VB_Description = "Makes a BG fade for your AOL 6.0 profile."
ProFader2$ = ProFader(Mid(Hex2Dec(HTML1$), 1, 2), Mid(Hex2Dec(HTML1$), 3, 2), Mid(Hex2Dec(HTML1$), 5, 2), Mid(Hex2Dec(HTML2$), 1, 2), Mid(Hex2Dec(HTML2$), 3, 2), Mid(Hex2Dec(HTML2$), 5, 2), Numb%)
End Function
Function RGB2HEX(rgb)
Attribute RGB2HEX.VB_Description = "Makes a HTML color from given RGB."
'heh, I didnt make this one...  <--monk-e-god didn't make this one neither :)
'function sampled from monk-e-fade3.bas
    a$ = Hex(rgb)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGB2HEX = a$
End Function

Function RGB2HTML(rgb)
Attribute RGB2HTML.VB_Description = "Makes a HTML color from given RGB."
RGB2HTML = VerseHTM(RGB2HEX(rgb))
End Function

Function Fade3Colr(r1%, g1%, b1%, r2%, g2%, b2%, r3%, g3%, b3%, txt$, Wavy As Boolean) As String
Attribute Fade3Colr.VB_Description = "Uses RGB color values and text to fade text from one color to another to another.  Wavy optional."
Dim Wave%, WaveH$
Wave% = 0
txtL% = Len(txt$) ' length of text ot be faded
txtL2% = Int(txtL% / 2) ' to fade text, you must do it one color at a time.  to do this, we have 3 colors,
' you start by fading the first half of the text from the first color to the second, then we have to
' fade the second half of the text from the second color to the third.
frst$ = Left(txt$, txtL2%) '  Read the Fade2Colr and the rest is easy :)
lst$ = Right(txt$, txtL% - txtL2%)
TLen% = Len(frst$)
For i = 1 To TLen%
lstchr$ = Mid(frst$, i, 1)
clr1 = rgb(((b2 - b1) / TLen% * i) + b1, ((g2 - g1) / TLen% * i) + g1, ((r2 - r1) / TLen% * i) + r1)
xClr = RGB2HEX(clr1)
If Wavy = True Then
Wave = Wave + 1
If Wave > 4 Then Let Wave = 1
If Wave = 1 Then Let WaveH$ = "<sup>"
If Wave = 2 Then Let WaveH$ = "</sup>"
If Wave = 3 Then Let WaveH$ = "<sub>"
If Wave = 4 Then Let WaveH$ = "</sub>"
Else
WaveH$ = ""
End If
Fade1$ = Fade1$ & "<font color=#" & xClr & ">" & WaveH$ & lstchr$
Next i
TLen% = Len(lst$)
For i = 1 To TLen%
If Wavy = True Then
Wave = Wave + 1
If Wave > 4 Then Let Wave = 1
If Wave = 1 Then Let WaveH$ = "<sup>"
If Wave = 2 Then Let WaveH$ = "</sup>"
If Wave = 3 Then Let WaveH$ = "<sub>"
If Wave = 4 Then Let WaveH$ = "</sub>"
Else
WaveH$ = ""
End If
lstchr$ = Mid(lst, i, 1)
clr2 = rgb(((b3 - b2) / TLen% * i) + b2, ((g3 - g2) / TLen% * i) + g2, ((r3 - r2) / TLen% * i) + r2)
xClr2 = RGB2HEX(clr2)
Fade2$ = Fade2$ & "<font color=#" & xClr2 & ">" & WaveH$ & lstchr$
Next i
Fade3Colr$ = Fade1$ & Fade2$
End Function

Function ProFade3Colr(r1%, g1%, b1%, r2%, g2%, b2%, r3%, g3%, b3%, txt$) As String
Attribute ProFade3Colr.VB_Description = "Fades text (HTML) and fits it for a profile."

txtL% = Len(txt$) ' length of text ot be faded
txtL2% = Int(txtL% / 2) ' to fade text, you must do it one color at a time.  to do this, we have 3 colors,
' you start by fading the first half of the text from the first color to the second, then we have to
' fade the second half of the text from the second color to the third.
frst$ = Left(txt$, txtL2%) '  Read the Fade2Colr and the rest is easy :)
lst$ = Right(txt$, txtL% - txtL2%)
TLen% = Len(frst$)
For i = 1 To TLen%
lstchr$ = Mid(frst$, i, 1)
clr1 = rgb(((b2 - b1) / TLen% * i) + b1, ((g2 - g1) / TLen% * i) + g1, ((r2 - r1) / TLen% * i) + r1)
xClr = RGB2HEX(clr1)
Fade1$ = Fade1$ & "< font . color=#" & xClr & ">" & lstchr$
Next i
TLen% = Len(lst$)
For i = 1 To TLen%
lstchr$ = Mid(lst, i, 1)
clr2 = rgb(((b3 - b2) / TLen% * i) + b2, ((g3 - g2) / TLen% * i) + g2, ((b3 - b2) / TLen% * i) + r2)
xClr2 = RGB2HEX(clr2)
Fade2$ = Fade2$ & "< font . color=#" & xClr2 & ">" & lstchr$
Next i
ProFade3Colr$ = Fade1$ & Fade2$
End Function



Function ALLFX(txt As String) As String
Attribute ALLFX.VB_Description = "Adds some text effects to a string via HTML tags."
' Does not fade.
Dim ax$, bx$, Wave%, WaveH$
For n = 1 To Len(txt$)
ax$ = Mid(txt$, n, 1)
Wave = Wave + 1
If Wave% > 4 Then Let Wave% = 1
If Wave% = 1 Then Let WaveH$ = "<b><sup>"
If Wave% = 2 Then Let WaveH$ = "</b><u></sup>"
If Wave% = 3 Then Let WaveH$ = "</u><s><sub>"
If Wave% = 4 Then Let WaveH$ = "</s><i></sub>"
bx$ = bx$ & WaveH$ & ax$ & "</i></u></s></b>"
Next n
ALLFX$ = bx$
End Function
Sub FadePreview(PicB As PictureBox, ByVal FadedText As String)
Attribute FadePreview.VB_Description = "Previews HTML-Faded in picturebox."
'sub sampled from monkefade3.bas
'by aDRaMoLEk
FadedText$ = Replacer(FadedText$, Chr(13), "+chr13+")
OSM = PicB.ScaleMode
PicB.ScaleMode = 3
TextOffX = 0: TextOffY = 0
StartX = 2: StartY = 0
PicB.Font = "Arial": PicB.FontSize = 10
PicB.FontBold = False: PicB.FontItalic = False: PicB.FontUnderline = False: PicB.FontStrikethru = False
PicB.AutoRedraw = True: PicB.ForeColor = 0&: PicB.Cls
For X = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, X, 1)
  If c$ = "<" Then
    TagStart = X + 1
    TagEnd = InStr(X + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    X = TagEnd + 1
    Select Case T$
      Case "u"
        PicB.FontUnderline = True
      Case "/u"
        PicB.FontUnderline = False
      Case "s"
        PicB.FontStrikethru = True
      Case "/s"
        PicB.FontStrikethru = False
      Case "b"    'start bold
        PicB.FontBold = True
      Case "/b"   'stop bold
        PicB.FontBold = False
      Case "i"    'start italic
        PicB.FontItalic = True
      Case "/i"   'stop italic
        PicB.FontItalic = False
      Case "sup"  'start superscript
        TextOffY = -1
      Case "/sup" 'end superscript
        TextOffY = 0
      Case "sub"  'start subscript
        TextOffY = 1
      Case "/sub" 'End Subscript
        TextOffY = 0
      Case Else
        If Left$(T$, 10) = "font color" Then 'change font color
          ColorStart = InStr(T$, "#")
          ColorString$ = Mid$(T$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = rgb(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If c$ = "+" And Mid(FadedText$, X, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        X = X + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print c$
        TextOffX = TextOffX + PicB.TextWidth(c$)
    End If
  End If
Next X
PicB.ScaleMode = OSM
End Sub

Function Replacer(TheStr As String, This As String, WithThis As String)
'function sampled from monkefade3.bas
'by monk-e-god
Dim STRwo13s As String
STRwo13s = TheStr
Do While InStr(1, STRwo13s, This)
DoEvents
thepos% = InStr(1, STRwo13s, This)
STRwo13s = Left(STRwo13s, (thepos% - 1)) + WithThis + Right(STRwo13s, Len(STRwo13s) - (thepos% + Len(This) - 1))
Loop

Replacer = STRwo13s
End Function

Function Hex2Dec!(ByVal strHex$)
Attribute Hex2Dec.VB_Description = "Returns an RGB value of a color from specified HTML."
'function sampled from monkefade3.bas
'by aDRaMoLEk
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function

Function GETVAL%(ByVal strLetter$)
'function sampled from monkefade3.bas
'by aDRaMoLEk
  Select Case strLetter$
    Case "0"
      GETVAL = 0
    Case "1"
      GETVAL = 1
    Case "2"
      GETVAL = 2
    Case "3"
      GETVAL = 3
    Case "4"
      GETVAL = 4
    Case "5"
      GETVAL = 5
    Case "6"
      GETVAL = 6
    Case "7"
      GETVAL = 7
    Case "8"
      GETVAL = 8
    Case "9"
      GETVAL = 9
    Case "A"
      GETVAL = 10
    Case "B"
      GETVAL = 11
    Case "C"
      GETVAL = 12
    Case "D"
      GETVAL = 13
    Case "E"
      GETVAL = 14
    Case "F"
      GETVAL = 15
  End Select
End Function
Function Lots(wChar As String, wNumber As Integer) As String
Attribute Lots.VB_Description = "Like ""String"" function only does multipul characters."
' Does the same as the "String(Number, Char)" function.
' Except this one will only allow 1 character, the First one.
If Len(wChar$) > 1 Then Let wChar$ = Left(wChar$, 1)
For i = 1 To wNumber
ax$ = ax$ & wChar$
Next i
Lots$ = ax$
End Function

Function PickLTR(txt As String, wLetter As String) As String
Attribute PickLTR.VB_Description = "Lets you specify a letter/character/number to place after each letter in a specified string."
If txt$ = "" Or Len(txt$) < 1 Then Exit Function
If wLetter$ = "" Then Let wLetter$ = "-"
If Len(wLetter$) > 1 Then wLetter$ = Left(wLetter$, 1)
If txt$ = "" Then Exit Function
For i = 1 To Len(txt$)
wLtr$ = Mid(txt$, i, 1)
ax$ = ax$ & wLtr$ & wLetter$
Next i
PickLTR$ = ax$

End Function

Function ReverseString(txt As String) As String
Attribute ReverseString.VB_Description = "Takes text, and reverses it."
' In Dos32.bas, this function was used in alot of other
' function, but dos made this function with a whole
' bunch of Do While loops & shit, that was not needed.
' As you can see, I simplified it and perfected it by adding
' my Reverse2 function which replaces a "\" with a "/",
' a "(" with a ")", and so on, and so on.  Dos just didn't
' want people claiming his code as their own.       :)
bx$ = txt$
For i = 1 To Len(txt$)
ax$ = Right(bx$, 1)
cx$ = cx$ & ax$
bx$ = Left(bx$, Len(bx$) - 1)
Next i
ReverseString$ = Reverse2(cx$)
End Function

Function GetChatText() As String
Attribute GetChatText.VB_Description = "Returns a string value of the text in an open AOL 6.0 chatroom."
roomy& = FindRoom
If roomy& = 0& Then Exit Function

cntl& = FindWindowEx(roomy&, 0&, "RICHCNTL", vbNullString)
dood$ = GetText(cntl&)
For i = 1 To Len(dood$)
ax$ = Mid(dood$, i, 1)
If ax$ = Chr(13) Then
Let bx$ = bx$ & vbNewLine
Else
Let bx$ = bx$ & ax$
End If

Next i
GetChatText = bx$

End Function

Function GetChr(txt As String) As Integer
Attribute GetChr.VB_Description = "Returns an integer value of the character code of a single specified letter."
Dim ax$, bx%
If Len(txt$) > 1 Then Let ax$ = Left(txt$, 1) Else ax$ = txt$
For i = 32 To 255
If ax$ = Chr(i) Then bx% = i
Next i
GetChr = bx
End Function

Function CUTit(text As String, wStyle As txtStyles) As String
Attribute CUTit.VB_Description = "Will return a string value or either numbers, or letters from a string.  Removes the unspecified."
For i = 1 To Len(text)
ax$ = Mid(text, i, 1)
If wStyle = wChars Then
    For n = 48 To 57
    If ax$ = Chr(n) Then bx$ = bx$ & ax$ Else bx$ = bx$ & ""
    Next n
ElseIf wStyle = wNumbers Then
    For n = 97 To 122
    If LCase(ax$) = Chr(n) Then bx$ = bx$ & ax$ Else bx$ = bx$ & ""
    Next n
End If
Next i
CUTit = bx$
End Function

Sub ScrollOut(frm As Form)
Attribute ScrollOut.VB_Description = "Makes a form scroll out from the top-left hand corner of the screen.  Used best in the Form_Load event of a form."
Dim ax%, bx%
ax% = frm.Width
bx% = frm.Height
frm.Height = 120
frm.Width = 120
frm.Top = 10
frm.Left = 10

If ax% > bx% Then
    For i = 120 To ax%
        frm.Width = i
        If frm.Height < bx% + 1 Then
            frm.Height = i
            For n = 1 To 10
            Pause 0.5
            Next n
        End If
    Next i
ElseIf bx% > ax% Then
    For i = 120 To bx%
        frm.Height = i
        If frm.Width < ax% + 1 Then
            frm.Width = i
            For n = 1 To 10
            Pause 0.5
            Next n
        End If
    Next i
End If
End Sub
Function Next3Chr(poo As String) As String
Attribute Next3Chr.VB_Description = "Takes a given 3 chr and finds its next 3 chr in abc order."
Dim X1$, X2$, X3$, X1C%, X2C%, X3C%
If Len(poo$) > 3 Then
Let ax$ = Right(poo$, 3)
Else
ax$ = poo$
End If
X1$ = Mid(ax$, 1, 1)
X2$ = Mid(ax$, 2, 1)
X3$ = Mid(ax$, 3, 1)
    X1C% = GetChr(X1$)
    X2C% = GetChr(X2$)
    X3C% = GetChr(X3$)
        X3C% = X3C% + 1
        If X3C% < 48 Then Let X3C% = 48
        If X3C% > 57 And X3C% < 97 Then Let X3C% = 97
        If X3C% > 122 Then
            Let X3C% = 48
            Let X2C% = X2C% + 1
        End If
        
        If X2C% < 48 Then Let X2C% = 48
        If X2C > 57 And X2C% < 97 Then Let X2C% = 97
        If X2C% > 122 Then
            Let X2C% = 48
            Let X1C% = X1C% + 1
        End If
        
        If X1C% < 97 Then Let X1C% = 97
        If X1C% > 122 Then Let X1C% = 97
 Next3Chr$ = Chr(X1C%) & Chr(X2C%) & Chr(X3C%)
            End Function
Sub Find3Chr(SN As String, frm As Form, Optional TillFound As Boolean = True)
Attribute Find3Chr.VB_Description = "Will try to make a 3 character AOL screen name until it finds one."
Keyword "parental controls"

Do: DoEvents
win& = AolKid(" AOL Parental Controls")
Loop Until win& <> 0&
Pause 0.7
btn& = FindWindowEx(win&, 0, AolIcon, vbNullString)
ClickIt btn&

Do: DoEvents
win2& = AolKid("AOL Screen Names")
Loop Until win2& <> 0&
Pause 0.7
Btn2& = FindWindowEx(win2&, 0&, AolIcon, vbNullString)
Btn2& = FindWindowEx(win2&, Btn2&, AolIcon, vbNullString)
ClickIt Btn2&

Do: DoEvents
win3& = FindWindow("_AOL_Modal", "Create a Screen Name")
Loop Until win3& <> 0&
Pause 0.55
Btn3& = FindWindowEx(win3&, 0&, AolIcon, vbNullString)
Btn3& = FindWindowEx(win3&, Btn3&, AolIcon, vbNullString)
ClickIt Btn3&

Do: DoEvents
Win4& = FindWindow("_AOL_Modal", "Create a Screen Name")
Loop Until Win4& <> 0&
Pause 0.55
btn4& = FindWindowEx(Win4&, 0&, AolIcon, vbNullString)
ClickIt btn4&

Do: DoEvents
Win5& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose a Screen Name")
Loop Until Win5& <> 0&
Pause 0.55
txt& = FindWindowEx(Win5&, 0&, "_AOL_Edit", vbNullString)
ChangeCap txt&, GetUser
btn5& = FindWindowEx(Win5&, 0&, AolIcon, vbNullString)
ClickIt btn5&

Do: DoEvents
msgWin& = FindWindow("#32770", "America Online")
Loop Until msgWin& <> 0&
Pause 0.55
msgBtn& = FindWindowEx(msgWin&, 0&, "Button", "OK")
ClickIt msgBtn&
ChangeCap txt&, GetUser
ClickIt btn5&

Do: DoEvents
win6& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose Another Screen Name")
Loop Until win6& <> 0&
Pause 0.55
btn6& = FindWindowEx(win6&, 0&, AolIcon, vbNullString)
For i = 1 To 2
btn6& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
Next i
ClickIt btn6&
Pause 0.55
txt2& = FindWindowEx(win6&, 0&, "_AOL_Edit", vbNullString)
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
ChangeCap txt2&, SN$
btnd& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
ClickIt btnd&
ax$ = SN$
If TillFound = True Then
Do: DoEvents
ax$ = Next3Chr(ax$)
Do: DoEvents
msgWin2& = FindWindow("#32770", "America Online")
pwWin& = FindWindow("_AOL_Modal", "Step 2 of 4: Choose a password")
Loop Until msgWin2& <> 0& Or pwWin& <> 0&
If msgWin2& <> 0& Then
Pause 0.55
Window_Close msgWin2&
End If
If frm.Tag = "found" Then Exit Sub
pwWin& = FindWindow("_AOL_Modal", "Step 2 of 4: Choose a password")
ChangeCap txt2&, ax$
btnd& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
ClickIt btnd&
Loop Until pwWin& <> 0& Or frm.Tag = "found"
End If

End Sub
Sub FakeInternal(SN As String, PW As String)
Attribute FakeInternal.VB_Description = "Makes a fake internal screen name with a given suffix.  Only works if the user has master capabilities."
Keyword "parental controls"

Do: DoEvents
win& = AolKid(" AOL Parental Controls")
Loop Until win& <> 0&
Pause 0.7
btn& = FindWindowEx(win&, 0, AolIcon, vbNullString)
ClickIt btn&

Do: DoEvents
win2& = AolKid("AOL Screen Names")
Loop Until win2& <> 0&
Pause 0.7
Btn2& = FindWindowEx(win2&, 0&, AolIcon, vbNullString)
Btn2& = FindWindowEx(win2&, Btn2&, AolIcon, vbNullString)
ClickIt Btn2&

Do: DoEvents
win3& = FindWindow("_AOL_Modal", "Create a Screen Name")
Loop Until win3& <> 0&
Pause 0.55
Btn3& = FindWindowEx(win3&, 0&, AolIcon, vbNullString)
Btn3& = FindWindowEx(win3&, Btn3&, AolIcon, vbNullString)
ClickIt Btn3&

Do: DoEvents
Win4& = FindWindow("_AOL_Modal", "Create a Screen Name")
Loop Until Win4& <> 0&
Pause 0.55
btn4& = FindWindowEx(Win4&, 0&, AolIcon, vbNullString)
ClickIt btn4&

Do: DoEvents
Win5& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose a Screen Name")
Loop Until Win5& <> 0&
Pause 0.55
txt& = FindWindowEx(Win5&, 0&, "_AOL_Edit", vbNullString)
ChangeCap txt&, "a12"
btn5& = FindWindowEx(Win5&, 0&, AolIcon, vbNullString)
ClickIt btn5&
Do: DoEvents
msgWin& = FindWindow("#32770", "America Online")
Loop Until msgWin& <> 0&
Pause 0.55
Window_Close msgWin&
ClickIt btn5&
Do: DoEvents
win6& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose Another Screen Name")
Loop Until win6& <> 0&
Pause 0.55
btn6& = FindWindowEx(win6&, 0&, AolIcon, vbNullString)
btn6& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
ClickIt btn6&
Pause 0.55
txt2& = FindWindowEx(win6&, 0&, "_AOL_Edit", vbNullString)
ChangeCap txt2&, String(10, "1")
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
ChangeCap txt2&, SN$ & String(10 - Len(SN$) + 1, " ") & "INT" ''  last textbox
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
btnd& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
btnd& = FindWindowEx(win6&, btnd&, AolIcon, vbNullString)
ClickIt btnd&
Do: DoEvents
sugWin& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose Another Screen Name")
Loop Until sugWin& <> 0&
sugbtn& = FindWindowEx(sugWin&, 0&, AolIcon, vbNullString)
For tk = 1 To 2
sugbtn& = FindWindowEx(sugWin&, sugbtn&, AolIcon, vbNullString)
Next tk
ClickIt sugbtn&
Pause 0.55
Do: DoEvents
win6& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose Another Screen Name")
dlist& = FindWindowEx(win6&, 0&, "_AOL_Listbox", vbNullString)
If dlist& = 0& Then win6& = 0&: Pause 0.55
Loop Until win6& <> 0&
Pause 0.55
btn6& = FindWindowEx(win6&, 0&, AolIcon, vbNullString)
btn6& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
btn6& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
ClickIt btn6&
Do: DoEvents
Do: DoEvents
win6& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose Another Screen Name")
Loop Until win6& <> 0&
dlist& = FindWindowEx(win6&, 0&, "_AOL_Listbox", vbNullString)
If dlist& <> 0& Then win6& = 0&: Pause 0.55
Loop Until win6& <> 0&

Pause 0.55
btn6& = FindWindowEx(win6&, 0&, AolIcon, vbNullString)
btn6& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
btn6& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
ClickIt btn6&
txt2& = FindWindowEx(win6&, 0&, "_AOL_Edit", vbNullString)
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
txt2& = FindWindowEx(win6&, txt2&, "_AOL_Edit", vbNullString)
btnd& = FindWindowEx(win6&, btn6&, AolIcon, vbNullString)
ChangeCap txt2&, "INT " & SN$
ClickIt btnd&
Do: DoEvents
pwWin& = FindWindow("_AOL_Modal", "Step 2 of 4: Choose a password")
Loop Until pwWin& <> 0&
Pause 0.55
lbtn& = FindWindowEx(pwWin&, 0&, AolIcon, vbNullString)
txt9& = FindWindowEx(pwWin&, 0&, "_AOL_Edit", vbNullString)
ChangeCap txt9&, PW$
txt9& = FindWindowEx(pwWin&, txt9&, "_AOL_Edit", vbNullString)
ChangeCap txt9&, PW$
ClickIt lbtn&

End Sub
Function VerseHTM$(HTM$)
Attribute VerseHTM.VB_Description = "Don't worry about this one and don't delete it."
If Len(HTM$) > 6 Then Let HTM$ = Left(HTM$, 6)
ax$ = Mid(HTM$, 1, 2)
bx$ = Mid(HTM$, 3, 2)
cx$ = Mid(HTM$, 5, 2)
VerseHTM$ = cx & bx & ax
End Function
Function HTML2RGB!(HTM$)
Attribute HTML2RGB.VB_Description = "Returns a RGB color from an HTML color code."
HTML2RGB! = Hex2Dec!(VerseHTM$(HTM$))
End Function
Function snglBold(txt As String) As String
Attribute snglBold.VB_Description = "Makes the first letter of each word in a string bolded."
For i = 1 To Len(txt$)
ax$ = Mid(txt$, i, 1)
bx$ = ax$
If ax$ = " " Then
Let bx$ = ax$ & "<b>" & Mid(txt$, i + 1, 1) & "</b>"
i = i + 1
End If
cx$ = cx$ & bx$
dx$ = "<b>" & Left(cx$, 1) & "</b>" & Right(cx$, Len(cx$) - 1)
Next i
snglBold = dx$
End Function

Function fntFade(txt As String) As String
Attribute fntFade.VB_Description = "Makes HTML text where each letter is a different font."
For i = 1 To Len(txt$)
ax$ = Mid(txt$, i, 1)
p = i
If i > Screen.FontCount Then Let p = i - Screen.FontCount
bx$ = bx$ & "<font face=""" & Screen.Fonts(p) & """>" & ax$
Next i
fntFade$ = bx$
End Function
Function TextWack(txt As String) As String
Attribute TextWack.VB_Description = "Makes text wacked out like lamers do on AOL."
For i = 1 To Len(txt$)
ax$ = Mid(txt$, i, 1)
Select Case LCase(ax$)
Case "w"
ax$ = "vv"
Case "e"
ax$ = "3"
Case "a"
ax$ = "4"
Case "i"
ax$ = "1"
Case "l"
ax$ = "1"
Case "r"
ax$ = "I2"
Case "s"
ax$ = "z"
Case "x"
ax$ = "×"
Case "m"
ax$ = "lVI"
End Select
bx$ = bx$ & ax$
Next i
TextWack$ = bx$
End Function
