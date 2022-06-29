Attribute VB_Name = "kEwL32"
'Prerelease version
'I plan on making details on every line of this .bas
'kEwL32.bas was kreated on Sunday, November 15, 4:49 pm Eastern Standard
'kEwL32.bas was kreated by Pharoah
'kEwL32.bas was kreated for purpose of learning API.  If you have any questons or if you experience difficulty, or want to make suggestions, mail me at Pharoah_mh@hotmail.com
'Send some subs and functions also that you think need to be added
'On every sub and function, there is a usage that you use with it in your button or whatever
'* = Can only be used in this bas file not in buttons or forms
'To change it so it can be used elsewhere, change Private to Public
'------------------------------Declarations------------------------------
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'------------------------------Constants------------------------------
'Sound Constants
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
'Window Message Constants
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_COMMAND = &H111
'Message Box Constants
Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_APPLMODAL = &H0&
Public Const MB_COMPOSITE = &H2
Public Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Public Const MB_DEFBUTTON1 = &H0&
Public Const MB_DEFBUTTON2 = &H100&
Public Const MB_DEFBUTTON3 = &H200&
Public Const MB_DEFMASK = &HF00&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONSTOP = MB_ICONHAND
Public Const MB_MISCMASK = &HC000&
Public Const MB_MODEMASK = &H3000&
Public Const MB_NOFOCUS = &H8000&
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_PRECOMPOSED = &H1
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_SETFOREGROUND = &H10000
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&
Public Const MB_TYPEMASK = &HF&
Public Const MB_USEGLYPHCHARS = &H4
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&
'Message Box Identification Constants
Public Const IDABORT = 3
Public Const IDIGNORE = 5
Public Const IDNO = 7
Public Const IDOK = 1
Public Const IDRETRY = 4
Public Const IDYES = 6
Public Const IDCANCEL = 2
'Handle of Window constants
Public Const HWND_TOPMOST = -1
'Show Window Position constants
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
'Show Window Constants
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_SHOW = 5
'Get Window Constants
Public Const GW_HWNDNEXT = 2
Public Const GW_CHILD = 5
Public Sub StayOnTop(Frm As Form) 'StayOnTop Me
'This keeps you form on top of all other windows
X = SetWindowPos(Frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
Public Sub PlayWav(FileName As String)
X& = sndPlaySound(FileName$, SND_ASYNC Or SND_NODEFAULT)
End Sub
Public Sub Pause(Duration) 'Pause .05
'This pauses what is going on in the code for as long as the duration in seconds.
Current = Timer
Do While Timer - Current < Duration
X = DoEvents()
Loop
End Sub
Public Sub WriteINI(AppName As String, KeyName As String, entry As String, FileName As String, AppPath As Boolean) 'WriteINI "Stuff", "Letters", "ABC", "FileName.stu", False)
'This will create an INI file with whatever you want in it.
'The example above will make this entry in C:\Windows\FileName.stu
'[Stuff]
'Letters=ABC
If AppPath = True Then
    FileName$ = App.Path & "\" & FileName$
End If
X& = WritePrivateProfileString(AppName, KeyName$, entry$, FileName$)
End Sub
Public Function ReadINI(AppName As String, KeyName As String, FileName As String, AppPath As Boolean) As String 'Letters$ = ReadINI("Stuff", "Letters", "FileName.stu", False)
'This reads the contents of an ini file
If AppPath = True Then
    FileName$ = App.Path & "\" & FileName$
End If
Nulls$ = String(255, 0)
ReadINI = Left(Nulls$, GetPrivateProfileString(AppName$, ByVal KeyName$, "a", Nulls$, Len(Nulls$), FileName$))
End Function
Public Sub MoveForm(Frm As Form) 'MoveForm Me
'This will move a form when a user activates a certain object
ReleaseCapture
Moving = SendMessage(Frm.hWnd, WM_NCLBUTTONDOWN, 2, 0)
End Sub
Public Sub MenuPopup(Frm As Form, MenuName As Menu, Xpos As Long, Ypos As Long) 'MenuPopup Me, frmMenus.mnuMenu, Label1.Left, Label1.Top + Label1.Height
'This creates a popup menu as easy as using Popupmenu.
'The example creates a popup menu on the left side of the label and on the bottom.
Frm.PopupMenu MenuName, 0, Xpos, Ypos
End Sub
Private Function GetAOLWindow() As Long '*
'This gets the AOL window handle by the class name since the title changes
GetAOLWindow = FindWindow("AOL Frame25", vbNullString)
End Function
Private Function FindClassName(hWnd As Long) As String '*
'Finds the class name of a window
Nulls$ = String$(200, 0)
Length& = GetClassName(hWnd, Nulls$, 199)
FindClassName = Left$(Nulls$, Length&)
End Function
Private Function FindWindowText(hWnd As Long) As String '*
'Finds the window text of a window
Nulls$ = String$(200, 0)
Length& = GetWindowText(hWnd, Nulls$, 199)
FindWindowText = Left$(Nulls$, Length&)
End Function
Private Function FindChildByClass(parenthwnd As Long, ClassToFind As String) As Long '*
'Finds the child by it's class name
ChildHandle& = GetWindow(parenthwnd&, GW_CHILD)
Do
ChildClass$ = FindClassName(ChildHandle&)
If ChildClass$ = ClassToFind$ Then
FindChildByClass = ChildHandle&
Else
ChildHandle& = GetWindow(ChildHandle&, GW_HWNDNEXT)
End If
Loop Until ChildClass$ = ClassToFind$ Or ChildHandle& = 0
End Function
Private Function FindChildByTitle(parenthwnd As Long, TitleToFind As String) As Long '*
'Finds a child by it's title
ChildHandle& = GetWindow(parenthwnd&, GW_CHILD)
Do
ChildText$ = FindWindowText(ChildHandle&)
If ChildText$ = TitleToFind$ Then
FindChildByTitle = ChildHandle&
Else
ChildHandle& = GetWindow(ChildHandle&, GW_HWNDNEXT)
End If
Loop Until ChildClass$ = ClassToFind$ Or ChildHandle& = 0
End Function
Private Function GetMDIClient() As Long '*
'Gets the MDIClient
AOL& = GetAOLWindow()
GetMDIClient = FindChildByClass(AOL&, "MDIClient")
End Function
Private Function GetWelcome() As Long '*
'Gets the welcome screen
MDI& = GetMDIClient()
ChildHandle& = GetWindow(MDI&, GW_CHILD)
Do
ChildText$ = FindWindowText(ChildHandle&)
If Left$(ChildText$, 9) = "Welcome, " Then
GetWelcome = ChildHandle&
Else
ChildHandle& = GetWindow(ChildHandle&, GW_HWNDNEXT)
End If
Loop Until Left$(ChildText$, 9) = "Welcome, " Or ChildHandle& = 0
End Function
Public Function GetVersion() As String 'Ver$ = GetVersion()
'Gets the version of AOL via the welcome screen
Welcome& = GetWelcome()
Child& = GetWindow(Welcome&, GW_CHILD)
Rich% = 0
Icon% = 0
Do
ChildClass$ = FindClassName(Child&)
Select Case ChildClass$
Case "_AOL_Static"
GetVersion = "2.5"
Exit Function
Case "_AOL_Icon"
Icon% = Icon% + 1
Case "RICHCNTL"
Rich% = Rich% + 1
End Select
Child& = GetWindow(Child&, GW_HWNDNEXT)
Loop Until Child& = 0
If Rich% = 5 And Icon% = 11 Then
GetVersion = "3.0"
End If
If Rich% = 8 And Icon% = 21 Then
GetVersion = "4.0"
End If
End Function
Public Function GetScreenName() As String 'SN$ = GetScreenName()
'Gets the screen name of the user
ChildText$ = FindWindowText(GetWelcome())
Spce = InStr(1, ChildText$, " ")
Exclm = InStr(Spce + 1, ChildText$, "!")
WordLen = (Exclm - Spce) - 1
If WordLen = -1 Then WordLen = 1
GetScreenName = Mid(ChildText$, Spce + 1, WordLen)
End Function
Private Function FindNext(Last1 As Long) As Long '*
'This will find the next class if the one you are looking for is not the one gotten
Next01& = GetWindow(Last1&, GW_HWNDNEXT)
Do
LastName$ = FindClassName(Last1&)
NextName$ = FindClassName(Next01&)
If NextName$ = LastName$ Then
FindNext = Next01&
Else
Next01& = GetWindow(Next01&, GW_HWNDNEXT)
End If
Loop Until NextName$ = LastName$ Or Next01& = 0
End Function
Private Function GetChatWindow() As Long '*
'Gets the chat room
'Note:  This is set to find the chatroom for all versions of AOL
MDI& = GetMDIClient()
Chat& = FindChildByClass(MDI&, "AOL Child")
Ver$ = GetVersion
Select Case Ver$
Case "4.0"
Do
List& = FindChildByClass(Chat&, "_AOL_Listbox")
Imag& = FindChildByClass(Chat&, "_AOL_Image")
Glph& = FindChildByClass(Chat&, "_AOL_Glyph")
Cmbo& = FindChildByClass(Chat&, "_AOL_Combobox")
Icon& = FindChildByClass(Chat&, "_AOL_Icon")
Stat& = FindChildByClass(Chat&, "_AOL_Static")
Rich& = FindChildByClass(Chat&, "RICHCNTL")
Rich2& = FindNext(Rich&)
If List& <> 0 Then
If Imag& <> 0 Then
If Glph& <> 0 Then
If Cmbo& <> 0 Then
If Icon& <> 0 Then
If Stat& <> 0 Then
If Rich& <> 0 Then
If Rich2& <> 0 Then
GetChatWindow = Chat&
S$ = "Found 4.0 Chat"
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Loop Until S$ = "Found 4.0 Chat" Or Chat& = 0
Case "3.0"
Do
View& = FindChildByClass(Chat&, "_AOL_View")
Edit& = FindChildByClass(Chat&, "_AOL_Edit")
Icon& = FindChildByClass(Chat&, "_AOL_Icon")
Imag& = FindChildByClass(Chat&, "_AOL_Image")
Stat& = FindChildByClass(Chat&, "_AOL_Static")
Glph& = FindChildByClass(Chat&, "_AOL_Glyph")
List& = FindChildByClass(Chat&, "_AOL_Listbox")
If View& <> 0 Then
If Edit& <> 0 Then
If Icon& <> 0 Then
If Imag& <> 0 Then
If Stat& <> 0 Then
If Glph& <> 0 Then
If List& <> 0 Then
GetChatWindow = Chat&
S$ = "Found 3.0 Chat"
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Loop Until S$ = "Found 3.0 Chat" Or Chat& = 0
Case "2.5"
Do
View& = FindChildByClass(Chat&, "_AOL_View")
Edit& = FindChildByClass(Chat&, "_AOL_Edit")
Icon& = FindChildByClass(Chat&, "_AOL_Icon")
Imag& = FindChildByClass(Chat&, "_AOL_Image")
Stat& = FindChildByClass(Chat&, "_AOL_Static")
Glph& = FindChildByClass(Chat&, "_AOL_Glyph")
List& = FindChildByClass(Chat&, "_AOL_Listbox")
If View& <> 0 Then
If Edit& <> 0 Then
If Icon& <> 0 Then
If Imag& <> 0 Then
If Stat& <> 0 Then
If Glph& <> 0 Then
If List& <> 0 Then
GetChatWindow = Chat&
S$ = "Found 2.5 Chat"
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Else
Chat& = GetWindow(Chat&, GW_HWNDNEXT)
End If
Loop Until S$ = "Found 2.5 Chat" Or Chat& = 0
End Select
End Function
Public Function GetChatView() As Long 'View& = GetChatView()
'Gets the View of the chat window
Chat& = GetChatWindow()
Ver$ = GetVersion()
Select Case Ver$
Case "4.0"
GetChatView = FindChildByClass(Chat&, "RICHCNTL")
Case "3.0"
GetChatView = FindChildByClass(Chat&, "_AOL_View")
Case "2.5"
GetChatView = FindChildByClass(Chat&, "_AOL_View")
End Select
End Function
Private Function GetChatEdit() As Long '*
'Gets the chat edit box for sending text
Chat& = GetChatWindow()
Ver$ = GetVersion
Select Case Ver$
Case "4.0"
List& = FindChildByClass(Chat&, "RICHCNTL")
GetChatEdit = FindNext(List&)
Case "3.0"
GetChatEdit = FindChildByClass(Chat&, "_AOL_Edit")
Case "2.5"
GetChatEdit = FindChildByClass(Chat&, "_AOL_Edit")
End Select
End Function
Private Sub SetText(hWnd As Long, Txt As String) '*
'Sets the text to an object
X& = SendMessageByString(hWnd, WM_SETTEXT, 0&, Txt)
End Sub
Private Sub SetChar(hWnd As Long, char As Long) '*
'Sets the character to an object.  This is usually used to hit the enter key
X& = SendMessageByNum(hWnd, WM_CHAR, char, 0&)
End Sub
Public Sub SendChat(Txt As String) 'SendChat "SuPZz Y'aLL"
'This sends chat to the chat room no matter what version of aol you have 2.5 - 4.0
Edit& = GetChatEdit()
DoEvents
SetChar Edit&, 13
DoEvents
Pause 0.3
SetText Edit&, Txt
DoEvents
SetChar Edit&, 13
DoEvents
End Sub
Public Sub RunMenu(Horizontal As Long, Vertical As Long) 'RunMenu 0, 0
'This runs a menu that is to the left and down.  The example above clicks on File and New
a = GetAOLWindow()
M = GetMenu(a)
sm = GetSubMenu(M, horz)
gi = GetMenuItemID(sm, vert)
F = SendMessageByNum(a, WM_COMMAND, gi, 0)
End Sub
Public Sub ButtonClick(hWnd As Long) 'ButtonClick IMSendButton&
'Clicks on a button
Down = SendMessageByNum(hWnd, WM_LBUTTONDOWN, 0&, 0&)
Pause 0.1
Up = SendMessageByNum(hWnd, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function AHREF(Address As String, Link As String, Unlink As String) As String 'AHref "http://www.aol.com", "Click Here", "to goto the aol webpage"
'Makes a link in a chat room with send text
AHREF = "< A HREF=" & Address$ & ">" & Link$ & "</A>" & Unlink$
End Function
Public Sub SendTextToKeywordEdit(Txt As String) 'SendTextToKeywordEdit("http://www.aol.com")
'This is for aol4.0 only because 2.5 and 3.0 don't have one on the toolbar
Tool& = GetAOLToolbar()
Tool_Bar& = FindChildByClass(Tool&, "_AOL_ToolBar")
Combo& = FindChildByClass(Tool_Bar&, "_AOL_Combobox")
Edit& = FindChildByClass(Combo&, "_AOL_Edit")
SetText Edit&, Txt$
Pause 0.05
SetChar Edit&, 13
End Sub
Private Function GetAOLToolbar() As Long '*
'Gets the AOL toolbar
MDI& = GetAOLMDIClient()
GetAOLToolbar = FindChildByClass(MDI&, "AOL Toolbar")
End Function
Public Function ChatLagg(Txt As String) As String 'SendChat ChatLagg("THIS IS A LAGG!!!!!!")
'Laggs the chat in aol4.0
a = Len(Txt)
For B = 1 To a
C = Left$(Txt, B)
D = Right$(C, 1)
E = "<HTML></HTML><B>" & D
F = F & E
Next B
ChatLagg = F
End Function
Public Function ClearChat() As String 'SendChat ClearChat
'Clears the chat room not eat it.
a$ = ".<pre=                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  "
a$ = a$ & "                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  "
ClearChat = a$
End Function
Public Function Wavy(Text1) 'SendChat Wavy("Do you like wavy text?")
'makes the chat wavy
py = 0
a = Len(Text1)
For B = 1 To a
C = Left(Text1, B)
D = Right(C, 1)
py = py + 1 '
If py = 1 Then
Msg = Msg & "<Sub>" & D
End If
If py = 2 Then
Msg = Msg & "<Sup>" & D
End If
If py = 3 Then
Msg = Msg & "</Sub>" & D
End If
If py = 4 Then
Msg = Msg & "</Sup>" & D
py = 0
End If
Next B
Wavy = Msg
End Function
Function CFade1(Back1 As Long, Back2 As Long, Txt As String, ByVal Wavy As Boolean) As String
a = Len(Txt)
For B = 1 To a
C = Left(Txt, B)
D = Right(C, 1)
E = ((Back1 - Back2) / 3) / a
F = E * B

Msg = Msg & "<Font Color=#" & F & ">" & D
Next B
CFade1 = Msg
End Function


