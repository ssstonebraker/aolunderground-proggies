Attribute VB_Name = "Module1"
 'CKM'S Module for AOL 6.0
 'Working 6.0 addroom sub included
 
 Option Explicit

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function MciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function EnumWindows Lib "user32" (ByVal wndenmprc As Long, ByVal lParam As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Const ABS_ALWAYSONTOP = &H2
Public Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const EM_GETLINE = &HC4
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETLINECOUNT = &HBA

Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8

Public Const SWP_NOMOVE = &H2
Public Const SW_SHOWNOACTIVATE = 4
Public Const SWP_HIDEWINDOW = &H80

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT


Global Const SWP_NOSIZE = 1
Public Const SW_RESTORE = 9

Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Const LB_SETITEMDATA = &H19A
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const CB_GETCOUNT = &H146
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETITEMDATA = &H150

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3

Public Const VK_TAB = &H9
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
Public Const WM_MOVE = &HF012 '
Public Const WM_SETTEXT = &HC '
Public Const WM_SYSCOMMAND = &H112 '
Public Const ENTER_KEY = 13 '

Private Const PROCESS_READ = &H10
Private Const RIGHTS_REQUIRED = &HF0000



Public Sub click(Button)
Call SendMessageLong(Button, WM_LBUTTONDOWN, 0, 0)
Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub sendtext(text)
    Dim Room As Long, Rich As Long
    Room& = FindRoom()
    Rich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
    Rich& = FindWindowEx(Room&, Rich&, "RICHCNTL", vbNullString)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, text)
    Call SendMessageLong(Rich&, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub StayOnTop(Form)
Call SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Public Sub KillWin(window&)
Call PostMessage(window&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub StayOnline()
'kinda old 'use timer.
Dim modl As Long, stat As Long, Txt As String, buttn As Long
modl& = FindWindow("_AOL_Modal", vbNullString)
stat& = FindWindowEx(modl&, 0, "_AOL_Static", vbNullString)
Txt = wintxt(stat&)
If Txt Like "*Do you wish to continue to stay online?*" Then
buttn& = FindWindowEx(modl&, 0, "_AOL_Icon", vbNullString)
click buttn&
Exit Sub
End If
End Sub
Public Sub Playwav(WavFile As String)
Call SndPlaySound(WavFile$, SND_ASYNC)
End Sub
Public Sub TimeOut(duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= duration
        DoEvents
    Loop
End Sub
Public Sub Upchat()
Dim aol As Long, MDI As Long, uploading As Long, Txt As String, mail As Long
Dim savemail As Long, buttn As Long
aol& = FindWindow("aol frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
uploading& = FindWindow("_AOL_Modal", vbNullString)
Txt = wintxt(uploading&)

If Txt Like "*File Transfer*" Then
mail& = FindWindowEx(MDI&, 0, "AOL Child", "Write Mail")
KillWin mail&
Do
savemail& = FindWindow("#32770", "America Online")
buttn& = FindWindowEx(savemail&, 0, "Button", vbNullString)
buttn& = FindWindowEx(savemail&, buttn&, "Button", vbNullString)
TimeOut 0.1
Loop Until savemail& <> 0 And buttn& <> 0

Call SendMessage(buttn&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(buttn&, WM_KEYUP, VK_SPACE, 0&)

Call ShowWindow(uploading&, SW_MINIMIZE)
Call ShowWindow(mail&, SW_MINIMIZE)
Call EnableWindow(aol&, True)
Else
MsgBox "You are not uploading any files", 64, "Error"
End If
End Sub
Public Sub KillWait()
 Dim mdal As Long, buttn As Long
 Call RunMenuByString("About")
 Do
 mdal& = FindWindow("_AOL_Modal", vbNullString)
 buttn& = FindWindowEx(mdal&, 0, "_AOL_Icon", vbNullString)
 Loop Until mdal& <> 0 And buttn& <> 0
 click buttn&
End Sub
Public Sub InstantMessage(sn$, Saywhat$)
Dim aol As Long, tool As Long, Toolbar As Long, MDI As Long, imz As Long, rec As Long, icn As Long, ScreenName As Long
Dim Msgfield As Long
aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icn& = FindWindowEx(Toolbar&, 0, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(Toolbar&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(Toolbar&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(Toolbar&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(Toolbar&, icn&, "_AOL_Icon", vbNullString)
Call PostMessage(icn&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(icn&, WM_KEYUP, VK_SPACE, 0&)

aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)

Do
DoEvents
imz& = FindWindowEx(MDI&, 0, "AOL Child", "Send Instant Message")
rec& = FindWindowEx(imz&, 0, "_AOL_Edit", vbNullString)
Loop Until imz& <> 0 And rec& <> 0

ScreenName& = SendMessageByString(rec&, WM_SETTEXT, 0, sn$)

Msgfield& = FindWindowEx(imz&, 0, "Richcntl", vbNullString)
Call SendMessageByString(Msgfield&, WM_SETTEXT, 0, Saywhat$)

icn& = FindWindowEx(imz&, 0, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(imz&, icn&, "_AOL_Icon", vbNullString)
Call PostMessage(icn&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(icn&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub KeyWord(KW As String)
    Dim aol As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Sub RunMenuByString(SearchString As String)
    Dim aol As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Public Sub mail(Screename, Subj, Saywhat)
Dim aol As Long, tool As Long, Toolbar As Long, icn As Long, MDI As Long
Dim writemail As Long, EditWin As Long, subject As Long, Msgfield As Long, send As Long, stat As Long
clickagain:
aol& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
icn& = FindWindowEx(Toolbar&, 0, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(Toolbar&, icn&, "_AOL_Icon", vbNullString)
icn& = FindWindowEx(Toolbar&, icn&, "_AOL_Icon", vbNullString)
Call PostMessage(icn&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(icn&, WM_KEYUP, VK_SPACE, 0&)

aol& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)


Do
DoEvents
writemail& = FindWindowEx(MDI&, 0, "AOL Child", "Write Mail")
EditWin& = FindWindowEx(writemail&, 0, "_AOL_Edit", vbNullString)
Call ShowWindow(writemail&, SW_HIDE)
Loop Until writemail& And EditWin&


Call SendMessageByString(EditWin&, WM_SETTEXT, 0, Screename)

subject& = FindWindowEx(writemail&, 0, "_AOL_Edit", vbNullString)
subject& = FindWindowEx(writemail&, subject&, "_AOL_Edit", vbNullString)
subject& = FindWindowEx(writemail&, subject&, "_AOL_Edit", vbNullString)
Call SendMessageByString(subject&, WM_SETTEXT, 0, Subj)

Msgfield& = FindWindowEx(writemail&, 0, "RICHCNTL", vbNullString)
Call SendMessageByString(Msgfield&, WM_SETTEXT, 0, Saywhat)

send& = FindWindowEx(writemail&, 0, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
send& = FindWindowEx(writemail&, send&, "_AOL_Icon", vbNullString)
Call PostMessage(send&, WM_LBUTTONDOWN, 0, 0)
Call PostMessage(send&, WM_KEYUP, VK_SPACE, 0&)

Do
DoEvents
stat& = FindWindow("_AOL_Modal", vbNullString)
icn& = FindWindowEx(stat&, 0, "_AOL_Icon", vbNullString)
click icn&
Loop Until stat& And icn&

click icn&

End Sub
Public Function FindRoom()
Dim aol As Long, MDI As Long, child As Long, Rich As Long, roomlist As Long, send As Long
   
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    roomlist& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
        
    If Rich& <> 0& And roomlist& <> 0& Then
    FindRoom = child&
    Exit Function
    
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            roomlist& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            If Rich& <> 0& And roomlist& <> 0& Then
     FindRoom = child&
     Exit Function
    
            End If
        Loop Until child& = 0&
    End If

End Function
Public Function GetMdi()
Dim aol As Long, md As Long
aol& = FindWindow("aol frame25", vbNullString)
md& = FindWindowEx(aol&, 0, "Mdiclient", vbNullString)
GetMdi = md&
End Function
Public Function GetUser()
'
Dim welcome As Long, MDI As Long, mystring As String
    Dim child As Long, UserString As String
    MDI& = GetMdi()
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    mystring$ = wintxt(child&)
     If mystring$ Like "*Welcome, *" Then
        mystring$ = Mid$(mystring$, 10, (InStr(mystring$, "!") - 10))
        GetUser = mystring$
        Exit Function
   
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            mystring$ = wintxt(child&)
            If mystring$ Like "*Welcome, *" Then
                mystring$ = Mid$(mystring$, 10, (InStr(mystring$, "!") - 10))
                GetUser = mystring$
                Exit Function
            
            End If
        Loop Until child& = 0&
    End If
    
End Function

Function GetLast$(ByVal Txt As String)
On Error Resume Next
Dim X
Do
X = X + 1
Loop Until Mid(Txt, Len(Txt) - X, 1) = Chr(13)
GetLast$ = Right$(Txt, X)
End Function

Public Function getchatname()
Dim X&, Txt As String
X& = FindRoom()
Txt = wintxt(X&)
getchatname = Txt
End Function

Public Sub addsnbuddylist(sn As String)

Dim MDI As Long, buddylist As Long, buttn As Long, buttn1 As Long, buttn2 As Long, buttn3 As Long, buttn4 As Long, setup As Long, addnew As Long
Dim editbox As Long, mystring As String
Call KeyWord("bv")
MDI& = GetMdi()

Do
buddylist& = FindWindowEx(MDI&, 0, "AOL Child", "Buddy List")
DoEvents
buttn& = FindWindowEx(buddylist&, 0, "_AOL_Icon", vbNullString)
buttn1& = FindWindowEx(buddylist&, buttn&, "_AOL_Icon", vbNullString)
buttn2& = FindWindowEx(buddylist&, buttn1&, "_AOL_Icon", vbNullString)
buttn3& = FindWindowEx(buddylist&, buttn2&, "_AOL_Icon", vbNullString)
buttn4& = FindWindowEx(buddylist&, buttn3&, "_AOL_Icon", vbNullString)
click buttn4&
setup& = FindWindowEx(MDI&, 0, "AOL Child", "Buddy List Setup")
Loop Until buddylist& <> 0 And setup& <> 0

Do
setup& = FindWindowEx(MDI&, 0, "AOL Child", "Buddy List Setup")
buttn& = FindWindowEx(setup&, 0, "_AOL_Icon", vbNullString)
Loop Until setup& <> 0 And buttn& <> 0
TimeOut 2
click buttn&

Do
addnew& = FindWindow("_AOL_Modal", "Add New Buddy")
editbox& = FindWindowEx(addnew&, 0, "_AOL_Edit", vbNullString)
buttn& = FindWindowEx(addnew&, 0, "_AOL_Icon", vbNullString)
Loop Until addnew& <> 0 And editbox& <> 0 And buttn& <> 0

mystring = SendMessageByString(editbox&, WM_SETTEXT, 0, sn)
click buttn&
KillWin setup&
KillWin buddylist&
End Sub
Public Sub IMsOn()
Call InstantMessage("$im_on", "CKM wuz here!")
End Sub
Public Sub IMsOff()
Call InstantMessage("$im_off", "CKM wuz here!")
End Sub

Public Sub listbox_Save(Directory As String, lst As listbox)
'example: Call Savelistbox("c:\MyList.txt", List1)
Dim X, ListItem As String
Open Directory$ For Output As #1
For X = 0 To lst.ListCount - 1
ListItem = lst.List(X)
Print #1, ListItem
Next X
Close #1
End Sub
Public Sub listbox_Load(Directory As String, lst As listbox)
Dim X, mystring As String
On Error Resume Next
Open Directory$ For Input As #1
If Err Then
MsgBox "File Does not exist", vbCritical, "File not found."
Exit Sub
End If
Do
Input #1, mystring
lst.AddItem mystring
Loop Until EOF(1)
Close #1

End Sub

Public Sub Listbox_Append(Directory As String, lst As listbox)
'add more stuff to a saved list
'example Call Listbox_Append("c:\mylog.txt", List1)
Dim X, ListItem As String
On Error Resume Next
Open Directory$ For Append As #1

If Err Then
MsgBox "File not found", vbCritical, "File not Found."
Exit Sub
End If

For X = 0 To lst.ListCount - 1
ListItem = lst.List(X)
Print #1, ListItem
Next X
Close #1
End Sub



Public Sub Cdopen()
Dim returnstring As String
Dim retvalue As Long
retvalue = MciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub
Public Sub Cdclose()
Dim returnstring As String
Dim retvalue As Long
retvalue = MciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub

Public Function wintxt(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    Buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, Buffer$)
    wintxt$ = Buffer$
End Function
Public Function aolversion()
Dim aol As Long, aMenu As Long, mCount As Long
    Dim SearchString As String
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            
            SearchString$ = "What's New in AOL 6.0"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            aolversion = "America Online 6.0"
            Exit Function
            End If
            SearchString$ = "What's New in AOL 5.0"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            aolversion = "America Online 5.0"
            Exit Function
            End If
            
            SearchString$ = "What's New in AOL 4.0"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            aolversion = "America Online 4.0"
            Exit Function
            End If
        
            SearchString$ = "Send an Instant Message"
            If InStr(LCase(sString$), LCase(SearchString$)) Then
            aolversion = "America Online 3.0"
            Exit Function
            End If
        
        Next LookSub&
    Next LookFor&
End Function

Public Sub PlayMp3(strPath As String)
'Note: This sub will play .wav and .midi files
    Dim lngLen As Long, strShort As String * 255, strPlay As String
    Call MciSendString("stop mp3play", 0, 0, 0)
    Call MciSendString("close mp3play", 0, 0, 0)
    lngLen = GetShortPathName(strPath, strShort, 255)
    strPlay = Left(strShort, lngLen)
    Call MciSendString("open " & strPlay & " type mpegvideo alias mp3play", 0, 0, 0)
    Call MciSendString("play mp3play", 0, 0, 0)
End Sub
Public Sub StopMp3()
    Call MciSendString("stop mp3play", 0, 0, 0)
    Call MciSendString("close mp3play", 0, 0, 0)
End Sub
Public Sub ListKillDupes(listbox As listbox)
'Kills dublicite items in a listbox
        Dim Search1 As Long
        Dim Search2 As Long
        Dim KillDupe As Long
KillDupe = 0
For Search1& = 0 To listbox.ListCount - 1
For Search2& = Search1& + 1 To listbox.ListCount - 1
KillDupe = KillDupe + 1
If listbox.List(Search1&) = listbox.List(Search2&) Then
listbox.RemoveItem Search2&
Search2& = Search2& - 1
End If
Next Search2&
Next Search1&
End Sub
Public Sub computername()
'Your computers name
Dim strString As String
    
    strString = String(255, Chr$(0))
    GetComputerName strString, 255
    strString = Left$(strString, InStr(1, strString, Chr$(0)) - 1)
    MsgBox strString
End Sub

Public Sub internetconnection()
Dim FLAGS As Long
Dim result As Boolean

    result = InternetGetConnectedState(FLAGS, 0)
    If result Then
     MsgBox "connected"
            
    Else
    
     MsgBox "Not connected"
    End If
End Sub
Public Sub AddRoomToList(thelist As listbox, AddUser As Boolean)


' Only use this sub if you know that you it's AOL 6
    thelist.Clear
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, ScreenName As String
    Dim psnHold As Long, rBytes As Long, Index As Long, Room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Room = FindRoom
    If Room& = 0& Then Exit Sub
    rList& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 28
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal ScreenName$, 4)
            psnHold& = psnHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, ScreenName$, Len(ScreenName$), rBytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
            If ScreenName$ <> GetUser Or AddUser = True Then
                thelist.AddItem ScreenName$
            End If
        
        
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Function FindIM() As Long
    
    Dim aol As Long, MDI As Long, child As Long, Caption As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = wintxt(child&)
    If InStr(Caption$, ">IM From:") = 1 Or InStr(Caption$, ">IM From:") = 2 Or InStr(Caption$, ">IM From:") = 3 Then
        FindIM& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Caption$ = wintxt(child&)
            If InStr(Caption$, ">IM From:") = 1 Or InStr(Caption$, ">IM From:") = 2 Or InStr(Caption$, ">IM From:") = 3 Then
                FindIM& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindIM& = child&
End Function

Public Function MKI(X As Integer) As String
  ' I didn't write this Function, but I have
  ' modified it.
  On Error GoTo hErr
  
  Dim Y As Long

  Y = CLng(X) And &HFFFF&
  MKI = Chr(Y And &HFF&) & Chr((Y And &HFF00&) \ &H100&)
      
Exit Function
hErr:
  
End Function
Public Function GetLine(hwnd As Long, lngLineNumber As Long) As String
  On Error GoTo hErr
    
  Dim li As Long, liCnt As Long, X As String
  Dim result As Long
  
  li = SendMessage(hwnd, EM_LINEINDEX, lngLineNumber, ByVal 0&)
  liCnt = SendMessage(hwnd, EM_LINELENGTH, li, ByVal 0&)
  X = MKI(CInt(liCnt)) & String(liCnt + 1, 0)

  result = SendMessage(hwnd, EM_GETLINE, lngLineNumber, ByVal X)
  
  GetLine = Left(X, liCnt)
  
Exit Function
hErr:
  
End Function

Public Function GetTextFromRich(WindowHandle As Long) As String
   
    Dim Buffer As String, TextLength As Long, txtlen As Long, linenum As Long
    Dim X
    TextLength& = SendMessage(WindowHandle, EM_GETLINECOUNT, 0&, 0&)
    linenum = TextLength&
    X = GetLine(WindowHandle, linenum)
    GetTextFromRich$ = X
   
End Function

