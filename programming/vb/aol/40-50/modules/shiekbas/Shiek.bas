Attribute VB_Name = "Shiek"
'NEED ANY MORE CODE????????????????????

' This Bas is for AOL 5.0,  Shiek V. 1.1


' E-mail me at Thatsquality@aol.com if you need me
' To write any functions at all or if you have ?'s
' or complaints. I ran out of ideas of stuff to add
' So e-mail me with ideas and ill put them on
' This was made with a little Dos code and i owe it all
' To dos because i read his tutorial to learn this crap
' So if you need any more code or anythin g e-mail me at

'               ThatsQuality@aol.com


Option Explicit
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessID Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long


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

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

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
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Function FindAOL() As Long
' just a sub for my subs, don't worry bout it
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
FindAOL& = AOL&
End Function
Function FindMDI() As Long
' just a sub for my subs, don't worry bout it
Dim MDI As Long, AOL As Long
AOL& = FindAOL
MDI& = FindWindowEx(AOL&, 0, "MDICLIENT", vbNullString)
FindMDI& = MDI&
End Function
Function FindRoom() As Long
' just a sub for my subs, don't worry bout it
Dim MDI As Long, Window As Long, Room As Long
Dim Rich As Long, AOLList As Long
Dim AOLIcon As Long, AOLStatic As Long
MDI& = FindMDI
Window& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
Rich = FindWindowEx(Window&, 0, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(Window&, 0, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(Window&, 0, "_AOL_icon", vbNullString)
AOLStatic& = FindWindowEx(Window&, 0, "_AOL_Static", vbNullString)
If Rich& <> 0 And AOLList& <> 0 And AOLIcon& <> 0 And AOLStatic& <> 0 Then
Room& = Window
Else
End If
Do While Window& <> 0
Window& = FindWindowEx(MDI&, Window&, "AOL Child", vbNullString)
Rich = FindWindowEx(Window&, 0, "RICHCNTL", vbNullString)
AOLList& = FindWindowEx(Window&, 0, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(Window&, 0, "_AOL_icon", vbNullString)
AOLStatic& = FindWindowEx(Window&, 0, "_AOL_Static", vbNullString)
If Rich& <> 0 And AOLList& <> 0 And AOLIcon& <> 0 And AOLStatic& <> 0 Then
Room& = Window
Else
End If
Loop
FindRoom& = Room&
End Function
Function GetUser() As String
' This Returns The SN of the user if they are online
Dim MDI As Long, Child As Long
Dim Caption As String, CaptionLength As Long
Dim WelcomeCaption As String, CommaPos As Long
Dim user As String, ExclamPos As Long, Length As Long
MDI& = FindMDI
Child& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
CaptionLength& = GetWindowTextLength(Child&)
Caption$ = String(CaptionLength, 0)
Call GetWindowText(Child&, Caption$, CaptionLength + 1)
If InStr(LCase(Caption$), LCase("Welcome")) = 1 Then
WelcomeCaption$ = Caption$
Else
End If
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
CaptionLength& = GetWindowTextLength(Child&)
Caption$ = String(CaptionLength, 0)
Call GetWindowText(Child&, Caption$, CaptionLength + 1)
If InStr(LCase(Caption$), LCase("welcome")) = 1 Then
WelcomeCaption$ = Caption$
Exit Do
Else
End If
Loop Until Child& = 0
If WelcomeCaption = "" Then
GetUser = ""
Exit Function
End If
CommaPos& = InStr(WelcomeCaption$, ",")
ExclamPos& = InStr(WelcomeCaption$, "!")
Length& = ExclamPos - CommaPos - 2
user$ = Mid(WelcomeCaption$, CommaPos + 2, Length)
GetUser$ = user$
End Function
Function FindWindowByCaption(Caption As String) As Long
' just a sub for my subs, don't worry bout it
Dim MDI As Long, Child As Long
MDI& = FindMDI
Child = FindWindowEx(MDI&, 0, "AOL Child", Caption$)
If Child <> 0 Then
FindWindowByCaption = Child&
End If
Child = FindWindow("AOL Modal", Caption$)
If Child& <> 0 Then
FindWindowByCaption = Child&
End If
End Function
Public Sub OpenNewMail()
If GetUser = "" Then
Exit Sub
End If
'This Opens the NewMail
Dim MDI As Long, Child As Long
Dim Caption As String, CaptionLength As Long
Dim welcome As Long, MailButton As Long
MDI& = FindMDI
Child& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
CaptionLength& = GetWindowTextLength(Child&)
Caption$ = String(CaptionLength, 0)
Call GetWindowText(Child&, Caption$, CaptionLength + 1)
If InStr(LCase(Caption$), LCase("Welcome")) = 1 Then
welcome& = Child
Else
End If
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
CaptionLength& = GetWindowTextLength(Child&)
Caption$ = String(CaptionLength, 0)
Call GetWindowText(Child&, Caption$, CaptionLength + 1)
If InStr(LCase(Caption$), LCase("welcome")) = 1 Then
welcome& = Child&
Exit Do
Else
End If
Loop Until Child& = 0
MailButton& = FindWindowEx(welcome&, 0, "_AOL_Icon", vbNullString)
MailButton& = FindWindowEx(welcome&, MailButton&, "_AOL_Icon", vbNullString)
ClickIcon MailButton
Dim Mail
Do
Mail = FindMailNew
TimeOut 0.1
Loop Until Mail <> 0
End Sub
Function FindMailCenter() As Long
' just a sub for my subs, don't worry bout it
Dim AOL As Long, Toolbar As Long, ToolbarChild As Long
Dim Icon As Long
AOL& = FindAOL
Toolbar& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
ToolbarChild& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
Icon& = FindWindowEx(ToolbarChild&, 0, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(ToolbarChild&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(ToolbarChild&, Icon&, "_AOL_Icon", vbNullString)
FindMailCenter& = Icon&
End Function
Public Sub OpenFlashMail()
'This will open the FlashMail Window
Dim MailCenterMenu As Long, WinVis As Long
Dim MailCenter As Long, Cursor As POINTAPI
MailCenter& = FindMailCenter
Call GetCursorPos(Cursor)
Call SetCursorPos(Screen.Height, Screen.Width)
Call PostMessage(MailCenter&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(MailCenter&, WM_LBUTTONUP, 0&, 0&)
Do
MailCenterMenu& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(MailCenterMenu)
Loop Until WinVis& = 1
Call PostMessage(MailCenterMenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(MailCenterMenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(MailCenterMenu&, WM_KEYDOWN, VK_UP, 0&)
    Call PostMessage(MailCenterMenu&, WM_KEYUP, VK_UP, 0&)
    Call PostMessage(MailCenterMenu&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call PostMessage(MailCenterMenu&, WM_KEYUP, VK_RIGHT, 0&)
    Call PostMessage(MailCenterMenu&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(MailCenterMenu&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(Cursor.X, Cursor.Y)
Dim Flashmail As Long
Do
Flashmail = FindMailFlash
TimeOut 0.1
Loop Until Flashmail <> 0
End Sub
Public Sub OpenMailFlash(Mail As Long)
If GetUser = "" Then
Exit Sub
End If
' This sub opens a mail in flashmail
' If you have mail = 0 then this will open the first mail
' if you have mail = 1 then this will open the second mail
' etc-- as long as the flashmail window is already open
Dim Flashmail As Long, MDI As Long, Tree As Long
MDI = FindMDI
Flashmail& = FindWindowEx(MDI&, 0, "AOL Child", "Incoming/Saved Mail")
Tree& = FindWindowEx(Flashmail&, 0, "_AOL_Tree", vbNullString)
Call SendMessage(Tree&, LB_SETCURSEL, Mail&, 0)
Call PostMessage(Tree&, WM_KEYDOWN, VK_RETURN, 0)
Call PostMessage(Tree&, WM_KEYUP, VK_RETURN, 0)
End Sub
Public Sub OpenMailNew(Mail As Long)
If GetUser = "" Then
Exit Sub
End If
' This sub opens a mail in Newmail
' If you have mail = 0 then this will open the first mail
' if you have mail = 1 then this will open the second mail
' etc-- as long as the Newmail window is already open
Dim Newmail As Long, Tree As Long
Dim TabControl As Long
Dim TabPage As Long
Newmail& = FindMailNew
TabControl = FindWindowEx(Newmail&, 0, "_AOL_TabControl", vbNullString)
TabPage = FindWindowEx(TabControl, 0, "_AOL_TabPage", vbNullString)
Tree& = FindWindowEx(TabPage&, 0, "_AOL_Tree", vbNullString)
Call SendMessage(Tree&, LB_SETCURSEL, Mail&, 0)
Call PostMessage(Tree&, WM_KEYDOWN, VK_RETURN, 0)
Call PostMessage(Tree&, WM_KEYUP, VK_RETURN, 0)
End Sub
Function FindMailNew() As Long
' just a sub for my subs, don't worry bout it
Dim MDI As Long, MailCaption As String, Newmail As Long
MDI = FindMDI
MailCaption$ = GetUser & "'s Online Mailbox"
Newmail& = FindWindowEx(MDI&, 0, "AOL Child", MailCaption)
FindMailNew& = Newmail&
End Function
Function FindMailWindow() As Long
' just a sub for my subs, don't worry bout it
'This gets the handle of an open mail window
' Which is the window with forward, reply etc... on it
Dim MDI As Long, Child As Long, Icon As Long
Dim TheMail As Long, Rich As Long, AOLStatic As Long
Dim ListBox As Long, Checkbox As Long, RichText As String
Dim Caption As String
MDI = FindMDI
Child& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
Checkbox = FindWindowEx(Child&, 0, "_AOL_Checkbox", vbNullString)
Caption = GetCaption(Child&)
Icon& = FindWindowEx(Child&, 0, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Rich = FindWindowEx(Child&, 0, "RICHCNTL", vbNullString)
RichText = GetText(Rich)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, AOLStatic, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
ListBox& = FindWindowEx(Child&, 0, "_AOL_Listbox", vbNullString)
If InStr(Caption, "Message") < 2 And InStr(Caption, GetUser) < 2 And InStr(RichText, "Date") > 5 And Icon& <> 0 And Rich <> 0 And AOLStatic <> 0 And ListBox = 0 And Checkbox& = 0 Then
TheMail& = Child
End If
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
Checkbox = FindWindowEx(Child&, 0, "_AOL_Checkbox", vbNullString)
Caption = GetCaption(Child)
Icon& = FindWindowEx(Child&, 0, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Child&, Icon&, "_AOL_Icon", vbNullString)
Rich = FindWindowEx(Child&, 0, "RICHCNTL", vbNullString)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, AOLStatic, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
AOLStatic = FindWindowEx(Child&, 0, "_AOL_Static", vbNullString)
ListBox& = FindWindowEx(Child&, 0, "_AOL_Listbox", vbNullString)
If InStr(Caption, "Message") < 2 And InStr(Caption, GetUser) < 2 And InStr(RichText, "Date") > 5 And Icon& <> 0 And Rich <> 0 And AOLStatic <> 0 And ListBox = 0 And Checkbox& = 0 Then
TheMail& = Child
End If
Loop Until Child = 0
FindMailWindow& = TheMail&
End Function
Sub TimeOut(Seconds)
'makes your program wait for the amount of time you wnat
Dim Start As Long
Start = Timer
Do While Timer - Start < Seconds
DoEvents
Loop
End Sub
Public Sub CloseWindow(Window As Long)
' just a sub for my subs, don't worry bout it
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub MailToListNew(TheList As ListBox)
If GetUser = "" Then
Exit Sub
End If
    'This sends all your new mail to a list you want
    ' Make sure your newmail is already open
    ' Dos's code
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailNew&
    If MailBox& = 0& Then Exit Sub
    TabControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
    TabPage& = FindWindowEx(TabControl&, 0&, "_AOL_TabPage", vbNullString)
    mTree& = FindWindowEx(TabPage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem MyString$
    Next AddMails&
End Sub
Public Sub MailList(List As ListBox, Person As String, Subject As String)
If GetUser = "" Then
Exit Sub
End If
' This will mail a list in server list format
Dim Message As String, ListText As String, Listcount As Long
Dim Counter As Long
Listcount = List.Listcount
ListText = List.List(0)
Message = "0)            " & ListText & Chr(13)
Counter = 1
Do
ListText = List.List(Counter)
If ListText <> "" Then
Message = Message & Counter & ")            " & ListText & Chr(13)
End If
Counter = Counter + 1
Loop Until ListText = ""
SendMail Person, Subject, Message
End Sub
Function FindMailFlash() As Long
' just a sub for my subs, don't worry bout it
Dim MDI As Long, Flashmail As Long
MDI& = FindMDI
Flashmail& = FindWindowEx(MDI&, 0, "AOL Child", "Incoming/Saved Mail")
FindMailFlash& = Flashmail&
End Function
Public Sub MailToListFlash(TheList As ListBox)
    'This sends all your flash  mail to a list you want
    ' Make sure your flash mail is already open
    ' Dos's code
    Dim MailBox As Long, TabControl As Long, TabPage As Long
    Dim mTree As Long, AddMails As Long, sLength As Long
    Dim Spot As Long, MyString As String, Count As Long
    MailBox& = FindMailFlash&
        If MailBox& = 0& Then Exit Sub
    mTree& = FindWindowEx(MailBox&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(mTree&, LB_GETCOUNT, 0&, 0&)
    If Count& = 0 Then Exit Sub
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(mTree&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(mTree&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        TheList.AddItem MyString$
    Next AddMails&
End Sub
Public Sub SetMailPref()
'This is important... it sets the mail preferences
'That this bas needs in order for tyhe server codes
'To work correctly
'Just call this sub every time the serve is started and
'everything will work fine
Dim AOL As Long, Toolbar As Long, IconSet As Long
Dim Icon As Long, MyAOL As Long, WinVis As Long
Dim Cursor As POINTAPI
AOL = FindAOL
Toolbar = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
If GetUser = "" Then
IconSet = 5
Else
IconSet = 6
End If
Icon = FindWindowEx(Toolbar, 0, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Toolbar, Icon&, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Toolbar, Icon&, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Toolbar, Icon&, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Toolbar, Icon&, "_AOL_Icon", vbNullString)
If IconSet = 6 Then
Icon = FindWindowEx(Toolbar, Icon&, "_AOL_Icon", vbNullString)
Else
End If
Call GetCursorPos(Cursor)
Call SetCursorPos(Screen.Height, Screen.Width)
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(Icon&, WM_LBUTTONUP, 0, 0)
Do
MyAOL& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(MyAOL&)
Loop Until WinVis& = 1
Call PostMessage(MyAOL&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(MyAOL&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(MyAOL&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(MyAOL&, WM_KEYUP, VK_DOWN, 0&)
    Call PostMessage(MyAOL&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(MyAOL&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(Cursor.X, Cursor.Y)
Dim Preferences As Long, MDI As Long
MDI& = FindMDI
Do
Preferences = FindWindowEx(MDI&, 0, "AOL Child", "Preferences")
TimeOut 0.3
Loop Until Preferences <> 0
Icon = FindWindowEx(Preferences, 0, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Preferences, Icon&, "_AOL_Icon", vbNullString)
Icon = FindWindowEx(Preferences, Icon&, "_AOL_Icon", vbNullString)
Call SendMessage(Icon&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(Icon&, WM_LBUTTONUP, 0, 0)
Dim Modal As Long, Check As Long
Do
Modal = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until Modal <> 0
Do
Check = FindWindowEx(Modal&, 0, "_AOL_Checkbox", vbNullString)
Loop Until Check <> 0
Call SendMessage(Check, BM_SETCHECK, False, vbNullString)
Check& = FindWindowEx(Modal&, Check&, "_AOL_Checkbox", vbNullString)
Call SendMessage(Check, BM_SETCHECK, True, vbNullString)
Icon& = FindWindowEx(Modal&, 0, "_AOL_Icon", vbNullString)
Do
Call SendMessage(Icon&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(Icon&, WM_LBUTTONUP, 0, 0)
WinVis = IsWindowVisible(Modal&)
Loop Until WinVis& = 0
Do
CloseWindow (Preferences&)
WinVis = IsWindowVisible(Preferences)
Loop Until WinVis = 0
End Sub
Public Sub Keyword(TheKeyWord As String)
If GetUser = "" Then
Exit Sub
End If
' Goto a keyword
Dim AOL As Long, Toolbar As Long, Toolbar2 As Long
Dim Icon As Long, KeyWordMenu As Long, WinVis As Long
Dim KeyWin As Long, Text As Long, Go As Long
Dim Counter As Long, Menu As Long, Cursor As POINTAPI
Dim MDI As Long
KillGlyph
If GetUser = "" Then
Exit Sub
End If
AOL& = FindAOL
Toolbar& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
Toolbar2& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
Icon& = FindWindowEx(Toolbar2&, 0, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar2&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar2&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar2&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar2&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar2&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar2&, Icon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(Cursor)
Call SetCursorPos(Screen.Height, Screen.Width)
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
Do
Menu& = FindWindow("#32768", vbNullString)
WinVis = IsWindowVisible(Menu&)
Loop Until WinVis = 1
Call PostMessage(Menu&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Menu&, WM_KEYUP, VK_DOWN, 0&)
Call PostMessage(Menu&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Menu&, WM_KEYUP, VK_DOWN, 0&)
Call PostMessage(Menu&, WM_KEYDOWN, VK_DOWN, 0&)
Call PostMessage(Menu&, WM_KEYUP, VK_DOWN, 0&)
Call PostMessage(Menu&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(Menu&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(Cursor.X, Cursor.Y)
MDI& = FindMDI
Counter = 0
Do
KeyWin& = FindWindowEx(MDI, 0, "Aol Child", "Keyword")
WinVis = IsWindowVisible(KeyWin)
TimeOut 0.1
Loop Until WinVis = 1
Do
Text& = FindWindowEx(KeyWin&, 0, "_AOL_Edit", vbNullString)
Loop Until Text& <> 0
Do
Go& = FindWindowEx(KeyWin&, 0, "_AOL_Icon", vbNullString)
Loop Until Go& <> 0
TextToText TheKeyWord, Text&
ClickIcon Go
End Sub
Public Sub OpenAFlashMailAndForward(TheMail As Long, Person As String, KillFwd As Boolean)
If GetUser = "" Then
Exit Sub
End If
' this will open the mail you want it to, if flashmail
' is allready open. Then it will forward it to the person
' you want it to. Killfwd is if you want the fwd: to be
' on the subject line or not
OpenMailFlash (TheMail&)
Dim Flash As Long
Do
Flash& = FindMailWindow
TimeOut 0.1
Loop Until Flash <> 0
If KillFwd = True Then
Call ForwardMail(Person$, True)
Else
Call ForwardMail(Person$, False)
End If
End Sub
Public Sub OpenANewMailAndForward(TheMail As Long, Person As String, KillFwd As Boolean)
If GetUser = "" Then
Exit Sub
End If
' this will open the mail you want it to, if the NEWmail
' is already open. Then it will forward it to the person
' you want it to. Killfwd is if you want the fwd: to be
' on the subject line or not
OpenMailNew (TheMail&)
If KillFwd = True Then
Call ForwardMail(Person$, True)
Else
Call ForwardMail(Person$, False)
End If
End Sub
Function LastChatLine() As String
If GetUser = "" Then
Exit Function
End If
'Gets the lastchatline
Dim Room As Long, RoomText As String, Length As Long
Dim Rich As Long, Enter As Long, Enter2 As Long
Dim FinalEnter As Long, Text As String
Room = FindRoom
Rich = FindWindowEx(Room&, 0, "RICHCNTL", vbNullString)
Length& = SendMessage(Rich&, WM_GETTEXTLENGTH, 0, 0)
RoomText$ = String(Length, 0)
Call SendMessageByString(Rich&, WM_GETTEXT, Length + 1, RoomText$)
Enter = InStr(RoomText, Chr(13))
Do
Enter = InStr(Enter + 1, RoomText$, Chr(13))
Enter2 = InStr(Enter + 1, RoomText$, Chr(13))
If Enter2 = 0 Then
FinalEnter = Enter
End If
Loop Until Enter2 = 0
Text$ = Mid(RoomText$, FinalEnter + 1, Length - FinalEnter + 1)
LastChatLine$ = Text$
End Function
Function LastChatLineSN() As String
If GetUser = "" Then
Exit Function
End If
' Gets the last chat line screen name
Dim SN As String, Length As Long, Chat As String
Chat$ = LastChatLine
Length& = InStr(Chat$, ":")
SN$ = Left(Chat$, Length - 1)
LastChatLineSN$ = SN$
End Function
Function LastChatLineText()
If GetUser = "" Then
Exit Function
End If
'gets the last chat line text
Dim Chat As String, Space As Long, Length As String
Dim Length2 As String, FinalLength As Long, Text As String
Dim Place As Long
Chat$ = LastChatLine
Space = InStr(Chat$, " ")
Length = Mid(Chat$, 1, 1)
Place = 2
Do
Length = Mid(Chat$, Place, 1)
Place = Place + 1
Length2 = Mid(Chat$, Place, 1)
If Length2 = "" Then
FinalLength = Place
End If
Loop Until Length2 = ""
Text$ = Mid(Chat$, Space + 2, FinalLength - Space + 1)
LastChatLineText = Text$
End Function
Function RoomName() As String
If GetUser = "" Then
Exit Function
End If
'Returns the name of the room you r in
' Use it like this: roomname = text1.text
Dim Room As Long, Caption As String, Length As Long
Room = FindRoom
Length& = GetWindowTextLength(Room&)
Caption$ = String(Length&, 0)
Call GetWindowText(Room&, Caption$, Length& + 1)
RoomName$ = Caption$
End Function
Public Sub SendIM(Person As String, Message As String)
If GetUser = "" Then
Exit Sub
End If
' Sends an instant message
Dim MDI As Long, IM As Long, Text1 As Long
Dim Text2 As Long, Send As Long, WinVis As Long
Call Keyword("aol://9293:")
MDI = FindMDI
Do
IM& = FindWindowEx(MDI, 0, "AOL Child", "Send Instant Message")
WinVis& = IsWindowVisible(IM&)
Loop Until WinVis = 1
Do
Text1& = FindWindowEx(IM&, 0, "_AOL_Edit", vbNullString)
WinVis& = IsWindowVisible(Text1&)
Loop Until WinVis = 1
Do
Text2& = FindWindowEx(IM&, 0, "RICHCNTL", vbNullString)
WinVis& = IsWindowVisible(Text2&)
Loop Until WinVis = 1
Call SendMessageByString(Text1&, WM_SETTEXT, 0&, Person$)
Call SendMessageByString(Text2&, WM_SETTEXT, 0&, Message$)
Send& = FindWindowEx(IM&, 0, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Do
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
WinVis = IsWindowVisible(Send&)
Loop Until WinVis = 1
Call SendMessage(Send&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(Send&, WM_LBUTTONUP, 0, 0)
CloseWindow IM
End Sub
Public Sub ChatSend(Message As String, Wavy As Boolean)
If GetUser = "" Then
Exit Sub
End If
' Sends lines to chat and makes them wavy if you want to
Dim Room As Long, Text As Long, Send As Long
Room& = FindRoom
Text& = FindWindowEx(Room&, 0, "RICHCNTL", vbNullString)
Text& = FindWindowEx(Room&, Text&, "RICHCNTL", vbNullString)
Send& = FindWindowEx(Room&, 0, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(Room&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(Room&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(Room&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(Room&, Send&, "_AOL_Icon", vbNullString)
Dim NewMessage As String, Letter As String, Counter As Long
Dim Letter2 As String
If Len(Message) > 1 And Wavy = True Then
Counter = 1
Do
Letter = Mid(Message, Counter, 1)
Letter2 = Mid(Message, Counter + 1, 1)
Counter = Counter + 2
NewMessage = NewMessage + Letter & "<sup>" & Letter2$ & "</sup>"
Loop Until Len(Message) < Counter - 1
Call SendMessageByString(Text&, WM_SETTEXT, 0, NewMessage$)
Else
Call SendMessageByString(Text&, WM_SETTEXT, 0, Message$)
End If
Call SendMessage(Send&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(Send&, WM_LBUTTONUP, 0, 0)
End Sub
Public Sub AddListToBuddyList(List As ListBox)
If GetUser = "" Then
Exit Sub
End If
' This obviuously adds a list to your buddylist
Dim TextHandle As Long, AddHandle As Long
Dim ListText As String, TextText As String
Dim BuddySetupEdit As Long, Buddysetup As Long
Dim ListLength As Long, ListBox As Long
Dim Length As Long, MDI As Long, Child As Long
MDI& = FindMDI
If GetUser = "" Then
Exit Sub
Else
End If
PressBuddySetupButton
PressBuddySetupEditButton
Buddysetup = FindWindowEx(MDI, 0, "AOL Child", GetUser & "'s Buddy List")
If Buddysetup = 0 Then
Buddysetup = FindWindowEx(MDI, 0, "AOL Child", GetUser & "'s Buddy Lists")
End If
Child = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
If InStr(GetCaption(Child), "Edit List") <> 0 Then
BuddySetupEdit = Child
End If
Do
Child = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
If InStr(GetCaption(Child), "Edit List") <> 0 Then
BuddySetupEdit = Child
End If
Loop Until Child& = 0
ListBox& = FindWindowEx(BuddySetupEdit, 0, "_AOL_Listbox", vbNullString)
TextHandle& = FindBuddySetupEditTextBox
AddHandle& = FindBuddySetupEditAdd
ListText$ = List.List(0)
Dim Length2 As Long
Do While ListText$ <> ""
If List.List(0) = "" Then
Exit Do
End If
Length& = SendMessage(ListBox, LB_GETCOUNT, 0, 0)
ListText$ = List.List(0)
TextToText ListText, TextHandle
ClickIcon AddHandle
List.RemoveItem (0)
Do
Length2& = SendMessage(ListBox, LB_GETCOUNT, 0, 0)
Loop Until Length2 > Length
Loop
Do
TimeOut 0.4
Loop Until GetText(TextHandle) = ""
AddHandle = FindWindowEx(BuddySetupEdit, AddHandle, "_AOL_Icon", vbNullString)
AddHandle = FindWindowEx(BuddySetupEdit, AddHandle, "_AOL_Icon", vbNullString)
Do
ClickIcon AddHandle
BuddySetupEdit = FindBuddySetupEdit
Loop Until BuddySetupEdit = 0
WaitForOK
CloseWindow Buddysetup
End Sub
Private Sub GetBuddySetupEditTextBoxText(TextHandle As Long, Text As String)
' just a sub for my subs, don't worry bout it
Dim TextLength As Long, Buffer As String
TextLength& = SendMessage(TextHandle&, WM_GETTEXTLENGTH, 0, 0)
Buffer$ = String(TextLength&, 0)
Call SendMessageByString(TextHandle&, WM_GETTEXT, TextLength& + 1, Buffer$)
Text$ = Buffer$
End Sub
Public Sub AddTextToAOLEdit(EditHandle As Long, Text As String)
'this is just a sub for my subs, don't worry bout it
Call SendMessageByString(EditHandle&, WM_SETTEXT, 0&, Text$)
End Sub
Function FindBuddySetupEditAdd() As Long
' just a sub for my subs, don't worry bout it
Dim EditHandle As Long, AddHandle As Long
EditHandle& = FindBuddySetupEdit
AddHandle& = FindWindowEx(EditHandle&, 0, "_AOL_Icon", vbNullString)
FindBuddySetupEditAdd& = AddHandle&

End Function
Function FindBuddySetupEditTextBox() As Long
' just a sub for my subs, don't worry bout it
Dim EditHandle As Long, TextHandle As Long
EditHandle& = FindBuddySetupEdit
TextHandle& = FindWindowEx(EditHandle&, 0, "_AOL_Edit", vbNullString)
TextHandle& = FindWindowEx(EditHandle&, TextHandle&, "_AOL_Edit", vbNullString)
FindBuddySetupEditTextBox& = TextHandle&
End Function
Public Sub PressBuddySetupEditButton()
If GetUser = "" Then
Exit Sub
End If
' just a sub for my subs, don't worry bout it
Dim EditHandle As Long, EditWindowHandle As Long
EditHandle& = FindBuddySetupEditButton
Call SendMessage(EditHandle&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(EditHandle&, WM_LBUTTONUP, 0, 0)
EditWindowHandle& = FindBuddySetupEdit
Dim List As Long
List& = FindWindowEx(EditWindowHandle&, 0, "_AOL_Listbox", vbNullString)
Do
TimeOut 0.5
EditWindowHandle& = FindBuddySetupEdit
List& = FindWindowEx(EditWindowHandle&, 0, "_AOL_Listbox", vbNullString)
Loop Until EditWindowHandle& <> 0 And List <> 0
TimeOut 0.5
End Sub
Function FindBuddySetupEdit() As Long

' just a sub for my subs, don't worry bout it
Dim MDI As Long, Child As Long, Caption As String
Dim EditHandle As Long
MDI& = FindMDI
Child& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
Caption$ = GetCaption(Child&)
If InStr(Caption$, "Edit List ") = 1 Then
EditHandle& = Child&
Else
End If
Child& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
Caption$ = GetCaption(Child&)
If InStr(Caption$, "Edit List ") = 1 Then
EditHandle& = Child&
Else
End If
FindBuddySetupEdit& = EditHandle&
End Function
Function FindBuddySetupEditButton() As Long
' just a sub for my subs, don't worry bout it
Dim SetupHandle As Long, EditHandle As Long
SetupHandle& = FindBuddySetup
EditHandle& = FindWindowEx(SetupHandle&, 0, "_AOL_Icon", vbNullString)
EditHandle& = FindWindowEx(SetupHandle&, EditHandle&, "_AOL_Icon", vbNullString)
FindBuddySetupEditButton& = EditHandle&
End Function
Function FindBuddyList() As Long
' just a sub for my subs, don't worry bout it
Dim MDI As Long, Handle As Long
MDI& = FindMDI
Handle& = FindWindowEx(MDI&, 0&, "AOL Child", "Buddy List Window")
FindBuddyList& = Handle&

End Function
Function FindBuddySetUpButton() As Long
' just a sub for my subs, don't worry bout it
Dim Buddy As Long, Icon As Long
Buddy& = FindBuddyList
Icon& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Buddy&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Buddy&, Icon&, "_AOL_Icon", vbNullString)
FindBuddySetUpButton& = Icon&
End Function
Private Sub PressBuddySetupButton()
If GetUser = "" Then
Exit Sub
End If
' just a sub for my subs, don't worry bout it
Dim Handle As Long
Handle& = FindBuddySetUpButton
Call SendMessage(Handle&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(Handle&, WM_LBUTTONUP, 0, 0)
Dim SetupHandle As Long
SetupHandle& = FindBuddySetup
Do
TimeOut 0.2
SetupHandle& = FindBuddySetup
Loop Until SetupHandle& <> 0
End Sub
Public Sub EnterPR(Room As String)
If GetUser = "" Then
Exit Sub
End If
'This will send you to a private room
Keyword "aol://2719:2-2-" + Room$
End Sub
Function FindBuddySetup() As Long
' just a sub for my subs, don't worry bout it
Dim MDI As Long, SetupHandle As Long
Dim Caption As String
MDI& = FindMDI
Caption$ = GetUser & "'s Buddy Lists"
SetupHandle& = FindWindowEx(MDI&, 0, "AOL Child", Caption$)
If SetupHandle = 0 Then
Caption$ = GetUser & "'s Buddy List"
SetupHandle& = FindWindowEx(MDI&, 0, "AOL Child", Caption$)
End If
FindBuddySetup& = SetupHandle&
End Function
Function GetCaption(Handle As Long) As String
' Returns the caption of a window who's handle you have
Dim Caption As String, TextLength As Long
TextLength& = SendMessage(Handle&, WM_GETTEXTLENGTH, 0, 0)
Caption = String(TextLength, 0)
Call SendMessageByString(Handle&, WM_GETTEXT, TextLength + 1, Caption$)
GetCaption$ = Caption
End Function
Public Sub ClickIcon(Icon As Long)
'Just a sub for my subs, don't worry about it
Call SendMessage(Icon&, WM_LBUTTONDOWN, 0, 0)
Call SendMessage(Icon&, WM_LBUTTONUP, 0, 0)
End Sub
Public Sub AntiIdle()
If GetUser = "" Then
Exit Sub
End If
'Run this on a loop to kill the idle
Dim Modal As Long, Icon As Long
Modal& = FindWindow("_AOL_Modal", vbNullString)
Icon& = FindWindowEx(Modal&, 0, "_AOL_Icon", vbNullString)
Call ClickIcon(Icon&)
End Sub
Public Sub CenterForm(TheForm As Form)
TheForm.Top = (Screen.Height * 0.85) / 2 - TheForm.Height / 2
TheForm.Left = Screen.Width / 2 - TheForm.Width / 2
End Sub
Public Sub RespondIM(Message As String)
If GetUser = "" Then
Exit Sub
End If
'This will respond to AN im with the text you want it to
' The only problem is if there is more than one im window
' Open it will send the text to a random window
' It will automatically press respond if needed
Dim IM As Long, Text As Long, Send As Long, Respond As Long
IM& = FindIM
If IM& = 0 Then
Exit Sub
End If
Send& = FindWindowEx(IM&, 0, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Send& = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
Text& = FindWindowEx(IM&, 0, "RICHCNTL", vbNullString)
Text& = FindWindowEx(IM&, Text&, "RICHCNTL", vbNullString)
If Text = 0 Then
Respond = FindWindowEx(IM&, Send&, "_AOL_Icon", vbNullString)
TimeOut 0.2
Text& = FindWindowEx(IM&, 0, "RICHCNTL", vbNullString)
Text& = FindWindowEx(IM&, Text&, "RICHCNTL", vbNullString)
End If
Call TextToText(Message$, Text&)
Call ClickIcon(Send&)
End Sub
Function FindIM() As Long
'this is just a sub for my subs, don't worry bout it
Dim IM As Long
Dim Child As Long, MDI As Long, Caption As String
MDI& = FindMDI
Child& = FindWindowEx(MDI&, 0, "AOL Child", vbNullString)
Caption$ = GetCaption(Child&)
If Mid(Caption$, 1, 8) = ">Instant" Then
IM& = Child&
End If
If Mid(Caption$, 3, 8) = "Instant " Then
IM& = Child&
End If
Do
Child& = FindWindowEx(MDI&, Child&, "AOL Child", vbNullString)
Caption$ = GetCaption(Child&)
If Mid(Caption$, 1, 8) = ">Instant" Then
IM& = Child&
End If
If Mid(Caption$, 3, 8) = "Instant " Then
IM& = Child&
End If
Loop Until Child& = 0
FindIM& = IM&
End Function
Public Sub TextToText(Text As String, Handle As Long)
' A sub for me
Call SendMessageByString(Handle&, WM_SETTEXT, 0, Text)
End Sub
Function SNFromIM() As String
'Returns the SN of the person your talking to
If GetUser = "" Then
Exit Function
End If
Dim IM As Long, Caption As String, Start As Long
Dim SN As String, Length As Long
IM& = FindIM
Caption = GetCaption(IM)
Start& = InStr(Caption$, ":")
Start& = Start + 2
Length& = GetStringLength(Caption$)
SN$ = Mid(Caption$, Start&, Length - Start + 1)
SNFromIM = SN
End Function
Function GetStringLength(TheString As String) As Long
'This gets a string's legnth
GetStringLength = Len(TheString)
End Function
Public Sub CloseWindowByName(name As String)
'Close a window by the windows caption(name)
Dim Window As Long, MDI As Long
MDI& = FindMDI
Window& = FindWindowEx(MDI&, 0, "AOL Child", name$)
CloseWindow Window
End Sub
Public Sub KillGlyph()
'Kill's that little AOL blue spinny thing in upper right
Dim AOL As Long, Toolbar As Long, Glyph As Long
AOL& = FindAOL
Toolbar& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
Glyph& = FindWindowEx(Toolbar&, 0, "_AOL_Glyph", vbNullString)
CloseWindow Glyph
End Sub
Function MessageFromIM() As String
If GetUser = "" Then
Exit Function
End If
'This gets you the last message from the im
Dim IM As Long, Length As Long, Colon As Long, Rich As Long
Dim Text As String, Colon2 As Long, Final As Long
Dim Message As String
IM& = FindIM
If IM& = 0 Then
Exit Function
End If
Rich& = FindWindowEx(IM&, 0, "RICHCNTL", vbNullString)
Text$ = GetText(Rich&)
Length& = GetStringLength(Text$)
Colon& = InStr(Text$, ":")
Colon2& = Colon&
Do
Colon& = InStr(Colon& + 1, Text$, ":")
If Colon& = 0 Then
Final& = Colon2
End If
Colon2& = Colon&
Loop Until Colon& = 0
Final = Final + 3
Message$ = Mid(Text$, Final&, Length - Final)
MessageFromIM$ = Message$
End Function
Public Sub WaitForOK()
'This waits for an OK box to come up then hits ok on it
Dim OKBox As Long, Icon As Long, Counter As Long
Dim WinVis As Long, Counter2 As Long
Counter& = 0
Counter2& = 0
Do
OKBox& = FindWindow("#32770", "America Online")
TimeOut 0.1
Loop Until OKBox <> 0
Icon& = FindWindowEx(OKBox&, 0, "Button", "OK")
Call SendMessage(Icon&, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(Icon&, WM_KEYUP, VK_SPACE, 0)
Call SendMessage(Icon&, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(Icon&, WM_KEYUP, VK_SPACE, 0)
End Sub

Function GetText(Handle As Long) As String
Dim Length As Long, Text As String
Length& = SendMessage(Handle&, WM_GETTEXTLENGTH, 0, 0)
Text$ = String(Length, 0)
Call SendMessageByString(Handle&, WM_GETTEXT, Length& + 1, Text$)
GetText$ = Text$
End Function
Public Sub PlaySound(Wav)
'Duh
Call sndPlaySound(Wav, SND_ASYNC Or SND_NODEFAULT)
End Sub
Public Sub AOLCaption(NewCaption As String)
' The classic sub that changes the aolcaption to NewCaption
Dim AOL As Long
AOL& = FindAOL
TextToText NewCaption, AOL&
End Sub
Public Sub Upchat()
If GetUser = "" Then
Exit Sub
End If
'Makes it so you can do stuff on AOL and upload too
Dim Modal As Long, Gauge As Long, AOL As Long
Dim Caption
AOL& = FindAOL
Modal& = FindWindow("_AOL_Modal", vbNullString)
Caption = GetCaption(Modal)
If InStr(Caption, "Transfer") > 2 Then
Call EnableWindow(AOL&, 1)
Call EnableWindow(Modal&, 0)
End If
End Sub
Public Sub UnUpchat()
If GetUser = "" Then
Exit Sub
End If
' Opposite of upchat
Dim Modal As Long, Gauge As Long, AOL As Long
Dim Caption As String
AOL& = FindAOL
Modal& = FindWindow("_AOL_Modal", vbNullString)
Caption = GetCaption(Modal)
If InStr(Caption, "Transfer") > 2 Then
Call EnableWindow(AOL&, 0)
Call EnableWindow(Modal&, 1)
End If
End Sub
Function EliteText(word$)
'Makes your words elite
'Jolt's code
Dim Made As String, q, Letter As String, leet As String, X
Made$ = ""
For q = 1 To Len(word$)
    Letter$ = ""
    Letter$ = Mid$(word$, q, 1)
    leet$ = ""
    X = Int(Rnd * 3 + 1)
    If Letter$ = "a" Then
    If X = 1 Then leet$ = "â"
    If X = 2 Then leet$ = "å"
    If X = 3 Then leet$ = "ä"
    End If
    If Letter$ = "b" Then leet$ = "b"
    If Letter$ = "c" Then leet$ = "ç"
    If Letter$ = "e" Then
    If X = 1 Then leet$ = "ë"
    If X = 2 Then leet$ = "ê"
    If X = 3 Then leet$ = "é"
    End If
    If Letter$ = "i" Then
    If X = 1 Then leet$ = "ì"
    If X = 2 Then leet$ = "ï"
    If X = 3 Then leet$ = "î"
    End If
    If Letter$ = "j" Then leet$ = ",j"
    If Letter$ = "n" Then leet$ = "ñ"
    If Letter$ = "o" Then
    If X = 1 Then leet$ = "ô"
    If X = 2 Then leet$ = "ð"
    If X = 3 Then leet$ = "õ"
    End If
    If Letter$ = "s" Then leet$ = "š"
    If Letter$ = "t" Then leet$ = "†"
    If Letter$ = "u" Then
    If X = 1 Then leet$ = "ù"
    If X = 2 Then leet$ = "û"
    If X = 3 Then leet$ = "ü"
    End If
    If Letter$ = "w" Then leet$ = "vv"
    If Letter$ = "y" Then leet$ = "ÿ"
    If Letter$ = "0" Then leet$ = "Ø"
    If Letter$ = "A" Then
    If X = 1 Then leet$ = "Å"
    If X = 2 Then leet$ = "Ä"
    If X = 3 Then leet$ = "Ã"
    End If
    If Letter$ = "B" Then leet$ = "ß"
    If Letter$ = "C" Then leet$ = "Ç"
    If Letter$ = "D" Then leet$ = "Ð"
    If Letter$ = "E" Then leet$ = "Ë"
    If Letter$ = "I" Then
    If X = 1 Then leet$ = "Ï"
    If X = 2 Then leet$ = "Î"
    If X = 3 Then leet$ = "Í"
    End If
    If Letter$ = "N" Then leet$ = "Ñ"
    If Letter$ = "O" Then leet$ = "Õ"
    If Letter$ = "S" Then leet$ = "Š"
    If Letter$ = "U" Then leet$ = "Û"
    If Letter$ = "W" Then leet$ = "VV"
    If Letter$ = "Y" Then leet$ = "Ý"
    If Len(leet$) = 0 Then leet$ = Letter$
    Made$ = Made$ & leet$
Next q

EliteText = Made$

End Function
Function CheckIfMaster() As Boolean
'This checks if the SN online is a master SN
If GetUser = "" Then
Exit Function
End If
Dim MDI As Long, Child As Long, Modal As Long, Icon As Long
Dim AOL As Long
AOL& = FindAOL
Keyword "aol://4344:1580.prntcon.12263709.564517913"
MDI& = FindMDI
Do
Child& = FindWindowEx(MDI&, 0, "AOL Child", " Parental Controls")
TimeOut 0.1
Icon& = FindWindowEx(Child&, 0, "_AOL_Icon", vbNullString)
Loop Until Child& <> 0 And Icon <> 0
Do
ClickIcon (Icon&)
Modal& = FindWindow("_AOL_Modal", vbNullString)
TimeOut 0.1
Loop Until Modal <> 0
If GetCaption(Modal) = "Parental Controls" Then
CheckIfMaster = True
Icon& = FindWindowEx(Modal&, 0, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Modal&, Icon&, "_AOL_Icon", vbNullString)
ClickIcon Icon&
ClickIcon Icon&
ClickIcon Icon&
ClickIcon Icon&
Else
Icon& = FindWindowEx(Modal&, 0, "_AOL_Icon", vbNullString)
ClickIcon Icon&
ClickIcon Icon&
ClickIcon Icon&
ClickIcon Icon&
CheckIfMaster = False
End If
CloseWindow (Child)
CloseWindow (Modal)
End Function
Public Function FileGetAttributes(TheFile As String) As Integer
    'Dos's Code
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        FileGetAttributes% = GetAttr(TheFile$)
    End If
End Function

Public Sub FileSetNormal(TheFile As String)
   'Dos's Code
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbNormal
    End If
End Sub

Public Sub FileSetReadOnly(TheFile As String)
    'Dos's Code
   Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub

Public Sub FileSetHidden(TheFile As String)
    'Dos's Code
   Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   'Dos's Code
   'Example x = getfromINI("hello","hey","C:\test.ini")
   ' x=15
   'Reads from C:\test.ini   Inside ini
   '[hello]
   ' hey = 15
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    'Dos's Code, it writes to the ini
    'Directory = "C:\windows\win.ini" - include filename
    ' In the INI:
    ' [Section]
    ' Key = KeyValue
   Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Sub StayOnTop(TheForm As Form)
'Sets a window so it will stay on top
Call SetWindowPos(TheForm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub
Public Sub StayNotOnTop(TheForm As Form)
'Sets a window to not be on top
Call SetWindowPos(TheForm.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub
Function CountFlashMail() As Long
'This will count your flashmail
Dim Flash As Long, Tree As Long, Count As Long
OpenFlashMail
Flash = FindMailFlash
Tree& = FindWindowEx(Flash&, 0, "_AOL_Tree", vbNullString)
Count& = SendMessage(Tree&, LB_GETCOUNT, 0, 0)
CountFlashMail& = Count&
CloseWindow Flash&
End Function
Function CountNewMail() As Long
'This will count your newmail
If GetUser = "" Then
Exit Function
End If
OpenNewMail
Dim Newmail As Long, Tree As Long
Dim TabControl As Long
Dim TabPage As Long, CountNewMail2 As Long
Do
Newmail& = FindMailNew
TimeOut 0.3
Loop Until Newmail <> 0
TabControl& = FindWindowEx(Newmail&, 0&, "_aol_tabcontrol", vbNullString)
TabPage& = FindWindowEx(TabControl&, 0&, "_aol_tabpage", vbNullString)
Tree& = FindWindowEx(TabPage&, 0&, "_aol_tree", vbNullString)
Do
CountNewMail& = SendMessage(Tree&, LB_GETCOUNT, 0, 0)
TimeOut 0.3
CountNewMail2& = SendMessage(Tree&, LB_GETCOUNT, 0, 0)
Loop Until CountNewMail = CountNewMail2 And CountNewMail <> 0
CloseWindow Newmail
End Function
Function GetProfile(ScreenName As String) As String
'This returns a members profile
' It will return Noprofile if there isn't a profile
If GetUser = "" Then
Exit Function
End If
Dim AOL As Long, MDI As Long, Toolbar As Long, Icon As Long
Dim Cursor As POINTAPI, WinVis As Long, PullDown As Long
Dim GetProfile2 As Long, Text As Long, Profile As Long
Dim NoProfile As Long, ProfileText As String, Profile2 As Long
Dim Button As Long
AOL& = FindAOL
MDI& = FindMDI
Toolbar& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
Icon& = FindWindowEx(Toolbar&, 0, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(Cursor)
Call SetCursorPos(Screen.Height, Screen.Width)
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
Do
PullDown& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(PullDown&)
Loop Until WinVis = 1
Call PostMessage(PullDown, WM_KEYDOWN, VK_UP, 0&)
Call PostMessage(PullDown&, WM_KEYUP, VK_UP, 0&)
Call PostMessage(PullDown&, WM_KEYDOWN, VK_UP, 0&)
Call PostMessage(PullDown&, WM_KEYUP, VK_UP, 0&)
Call PostMessage(PullDown&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(PullDown&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(Cursor.X, Cursor.Y)
Do
GetProfile2& = FindWindowEx(MDI&, 0, "AOL Child", "Get a Member's Profile")
TimeOut 0.1
Loop Until GetProfile2& <> 0
Text& = FindWindowEx(GetProfile2&, 0, "_AOL_Edit", vbNullString)
Icon& = FindWindowEx(GetProfile2&, 0, "_AOL_Icon", vbNullString)
Call TextToText(ScreenName, Text&)
Call ClickIcon(Icon&)
Do
Call ClickIcon(Icon&)
Profile& = FindWindowEx(MDI&, 0&, "AOL Child", "Member Profile")
Profile2 = FindWindowEx(Profile&, 0&, "RICHCNTL", vbNullString)
ProfileText$ = GetText(Profile2)
NoProfile& = FindWindow("#32770", "America Online")
Loop Until ProfileText$ <> "" Or NoProfile& <> 0
If NoProfile& <> 0 Then
GetProfile = "No Profile"
Button& = FindWindowEx(NoProfile, 0, "Button", "OK")
Call SendMessage(Button&, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(Button&, WM_KEYUP, VK_SPACE, 0)
Call CloseWindow(NoProfile)
Else
TimeOut 0.2
ProfileText$ = GetText(Profile2)
GetProfile$ = ProfileText
Call CloseWindow(Profile)
End If
Call CloseWindow(GetProfile2)
End Function
Public Function FileExists(sFileName As String) As Boolean
'Dos's code
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
Public Sub HideAWindow(Window As Long)
Call ShowWindow(Window&, SW_HIDE)
End Sub
Public Sub ShowAWindow(Window As Long)
' A sub for the bas unless you know API
Call ShowWindow(Window&, SW_SHOW)
End Sub
Public Sub AddBuddyListToListBox(TheList As ListBox)
' This adds all your people on your buddylist to a list box
If GetUser = "" Then
Exit Sub
End If
'mostly Dos's code
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    PressBuddySetupButton
    PressBuddySetupEditButton
    Room& = FindBuddySetupEdit
    If Room& = 0& Then
    Exit Sub
    Else
    End If
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessID(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            TheList.AddItem ScreenName$
            Next Index&
        Call CloseWindow(mThread)
    End If
    CloseWindow (FindBuddySetup)
    CloseWindow (FindBuddySetupEdit)
End Sub
Public Sub SendMail(Person As String, Subject As String, Message As String)
If GetUser = "" Then
Exit Sub
End If
'Simply sends a mail
Dim MDI As Long, Mail As Long, Text1 As Long, Text2 As Long
Dim RichCntl As Long, AOL As Long, Toolbar As Long
Dim Icon As Long
MDI& = FindMDI
AOL& = FindAOL
Toolbar& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(Toolbar&, 0, "_AOL_Toolbar", vbNullString)
Icon& = FindWindowEx(Toolbar&, 0, "_AOL_Icon", vbNullString)
Icon& = FindWindowEx(Toolbar&, Icon&, "_AOL_Icon", vbNullString)
Call PostMessage(Icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Icon&, WM_LBUTTONUP, 0&, 0&)
Do
Mail& = FindWindowEx(MDI&, 0, "AOL Child", "Write Mail")
TimeOut 0.1
Loop Until Mail& <> 0
Do
Text1& = FindWindowEx(Mail&, 0, "_AOL_Edit", vbNullString)
Text2& = FindWindowEx(Mail&, Text1, "_AOL_Edit", vbNullString)
Text2& = FindWindowEx(Mail&, Text2, "_AOL_Edit", vbNullString)
RichCntl& = FindWindowEx(Mail&, 0, "RICHCNTL", vbNullString)
Loop Until Text1 <> 0 And Text2 <> 0 And RichCntl <> 0
TextToText Person, Text1
TextToText Subject, Text2
TextToText Message, RichCntl
Dim RichText As String
Do
RichText = GetText(RichCntl)
TimeOut 0.15
Loop Until RichText <> "" Or Message = ""
Dim Counter As Long
Icon& = FindWindowEx(Mail&, 0, "_AOL_Icon", vbNullString)
Do
Icon& = FindWindowEx(Mail&, Icon&, "_AOL_Icon", vbNullString)
Counter = Counter + 1
Loop Until Counter = 15
ClickIcon Icon&
TimeOut 0.4
HideAWindow Mail&
End Sub
Public Sub AddRoomToListBox(TheList As ListBox, Addself As Boolean)
'Read title
If GetUser = "" Then
Exit Sub
End If
'mostly Dos's code
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom
    If Room& = 0& Then
    Exit Sub
    Else
    End If
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessID(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            TheList.AddItem ScreenName$
            Next Index&
        Call CloseWindow(mThread)
    End If
End Sub
Public Sub AddRoomToComboBox(TheList As ComboBox, Addself As Boolean)
' Read title
If GetUser = "" Then
Exit Sub
End If
'mostly Dos's code
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room& = FindRoom
    If Room& = 0& Then
    Exit Sub
    Else
    End If
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessID(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            TheList.AddItem ScreenName$
            Next Index&
        Call CloseWindow(mThread)
    End If
End Sub
Public Sub AddTreeToList(AOLTree As Long, List As ListBox, Addself As Boolean)
'this is just a sub for my subs, don't worry bout it
Dim AOLChild As Long, Counter As Long, Count As Long
Count = SendMessage(AOLTree&, LB_GETCOUNT, 0&, 0&)
Dim Text As String, lIndex As Long, Length As Long
Counter = 0
Do
Length = SendMessage(AOLTree, LB_GETTEXTLEN, 0, 0)
lIndex& = Counter
Text$ = String(Length, 0)
Call SendMessageByString(AOLTree&, LB_GETTEXT, lIndex&, Text$)
If Text$ <> GetUser Or Addself = True Then
List.AddItem Text$
End If
Counter = Counter + 1
Loop Until Counter + 1 > Count
End Sub
Public Sub ForwardMail(Person As String, KillFwd As Boolean)
If GetUser = "" Then
Exit Sub
End If
' If you have an mail open( like you just hit read and it
' Opened, then this will hit the forward button and send
' It to the person you want it to, killing the FWD: in
' The beggining if you want it to,
Dim MDIClient As Long, Caption As String, Sendwindow2 As Long
Dim AOLChild As Long, SendWin As Long, NewSubject As String
Dim AOLIcon As Long, AOLEdit As Long, Subject As String
MDIClient = FindMDI
Do
AOLChild& = FindMailWindow
TimeOut 0.1
Loop Until AOLChild& <> 0
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_aol_icon", vbNullString)
Caption = GetCaption(AOLChild&)
Do
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
SendWin& = FindWindowEx(MDIClient, 0, "AOL Child", "Fwd: " & Caption)
TimeOut 0.4
Loop Until SendWin <> 0
AOLEdit& = FindWindowEx(SendWin, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Person)
AOLEdit& = FindWindowEx(SendWin, AOLEdit&, "_aol_edit", vbNullString)
AOLEdit& = FindWindowEx(SendWin, AOLEdit&, "_aol_edit", vbNullString)
If KillFwd = True Then
Subject = GetText(AOLEdit)
NewSubject = Mid(Subject, 6, Len(Subject) - 4)
TextToText NewSubject, AOLEdit
End If
AOLIcon& = FindWindowEx(SendWin, 0&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
AOLIcon& = FindWindowEx(SendWin, AOLIcon&, "_aol_icon", vbNullString)
Call SendMessageLong(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
HideAWindow (SendWin)
CloseWindow (AOLChild)
Sendwindow2 = FindWindowEx(MDIClient, 0, "AOL Child", "Fwd: " & Caption)
AOLEdit = FindWindowEx(Sendwindow2, 0, "_AOL_Edit", vbNullString)
If GetText(AOLEdit) = "" Then
CloseWindow (Sendwindow2)
End If
Do
Sendwindow2 = FindWindowEx(MDIClient, Sendwindow2, "AOL Child", "Fwd: " & Caption)
AOLEdit = FindWindowEx(Sendwindow2, 0, "_AOL_Edit", vbNullString)
If GetText(AOLEdit) = "" Then
CloseWindow (Sendwindow2)
End If
Loop Until Sendwindow2 = 0

End Sub

Public Sub CloseFlashMail()
'This will close your flashmail if it is open
Dim Flashmail As Long
Flashmail = FindMailFlash
CloseWindow (Flashmail)
End Sub
Public Sub CloseNewail()
'This will close your Newmail if it is open
Dim Newmail As Long
Newmail = FindMailNew
CloseWindow (Newmail)
End Sub
Function FindSignon() As Long
Dim MDI As Long, Signon As Long
MDI& = FindMDI
Signon& = FindWindowEx(MDI&, 0, "AOL Child", "Goodbye from America Online!")
FindSignon& = Signon
End Function
Public Sub AddBuddyToList(List As ListBox)
Dim Buddy As Long, Tree As Long
Buddy = FindBuddyList
Tree = FindWindowEx(Buddy&, 0, "_AOL_Tree", vbNullString)
AddTreeToList Tree, List, False
End Sub
Function Chat_Hyperlink(Where As String, WhatToSay As String)
Chat_Hyperlink = "<a href=""""><a href=""""><a href=" & Where & ">" & WhatToSay
End Function
