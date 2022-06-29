Attribute VB_Name = "Shadow32"
'——————————————————————————————————————————————————————'
'                   Shadow 32  v1:                     '
'                                                      '
' First of all, I'd like to officially state that many '
'   of these procedures came from Jaguar32 and I just  '
' edited them to make them better.  So don't accuse me '
'   of pirating the BAS.  One of the things I did was  '
'clean up, so don't you AOL 3 proggers who live in the '
'  past throw this away.  There is much more order to  '
'   it and a few more procedures.  This BAS is mainly  '
' for everyone else.  AOL 4 proggers that need a BAS as'
'    as good as Jaguar, but for AOL 4.  And AOL 3/4    '
'proggers, that don't want to bother with detecting AOL'
' versions.  Well, this solves problems for both those '
'  groups.  It has has nearly flawless codes that edit '
'   themselves based on the AOL version of the Person  '
'                   using your prog.                   '
'                                                      '
'                   Shadow 32  v2:                     '
'                                                      '
'New in this version: small updates on a few codes, and'
'      a few more API and standard VB procedures.      '
'        If you have any questions, E-Mail me at       '
'              XIXShadow@hotmail.com.  Enjoy.             '
'——————————————————————————————————————————————————————'



Dim Going As Boolean
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal cmd As Long) As Long
Declare Function CreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function PostMessageBynum Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (Object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal Size As Integer, ByVal foreignPID As Long)
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
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Integer
Declare Function FillRect Lib "user32" (ByVal hdc As Integer, lpRect As RECT, ByVal hBrush As Integer) As Integer
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer

Public Const SPI_SCREENSAVERRUNNING = 97

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

Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

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

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

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

Enum MailBox
 NewMail = 1
 OldMail = 2
 SentMail = 3
End Enum

Enum PopupMenu
 MailCenter = 2
 MyFiles = 4
 MyAOL = 5
 Favorites = 6
 Internet = 7
 Channels = 8
 People = 9
End Enum

Sub AboutBox(Optional message As String)
'Shadow:
'this makes a Messagebox that displays your programs Name
'and your Name and your Message if you put it in
Msg = App.ProductName & " was created by " & App.CompanyName & "."
If message <> vbNullString Then Msg = Msg & "  " & message
MsgBox Msg, vbInformation
End Sub

Sub ADD_AOL_LB(item As String, List As ListBox)
'Jaguar:
'Add a list of Names to a VB ListBox
'This is usually called by another one of my functions

If List.ListCount = 0 Then
List.AddItem item
Exit Sub
End If
Do Until XX = (List.ListCount)
Let diss_Item$ = List.List(XX)
If Trim(LCase(diss_Item$)) = Trim(LCase(item)) Then Let do_it = "NO"
Let XX = XX + 1
Loop
If do_it = "NO" Then Exit Sub
List.AddItem item
End Sub

Sub AddListToString(List As ListBox)
'Jaguar:
'this will take a list and make it into a string
'and place a comma after each Item in the list
'this is good for a mass mailer
For DoList = 0 To List.ListCount - 1
AddListToStrin = AddListToStrin & List.List(DoList) & ", "
Next DoList
AddListToStrin = Mid(AddListToStrin, 1, Len(AddListToStrin) - 2)
End Sub

Sub AOLAddRoom(List As ListBox)
'Jaguar:
'This calls a function in 311.dll that retreives the Names
'from the AOL listbox.
'I have added some code so that it removes the
'garbage at the end of the listbox and also removes
'the user's SN from the listbox as well
Dim Index As Long
Dim I As Integer
If AOLGetChatName = vbNullString Then
 MsgBox "Not in a chat room.", vbExclamation
 Exit Sub
End If
For Index = 0 To 25
Namez$ = String$(256, " ")
ret = AOLGetList(Index, Namez$) ' & ErB$
Namez$ = Left$(Trim$(Namez$), Len(Trim(Namez$)))

ADD_AOL_LB Namez$, List
Next Index
end_addr:
List.RemoveItem List.ListCount - 1

I = GetListIndex(List, AOLGetUser())
If I <> -2 Then List.RemoveItem I
End Sub

Sub AddStringToList(Items As String, List As ListBox)
If Not Mid(Items, Len(Items), 1) = "," Then
Items = Items & ","
End If

For DoList = 1 To Len(Items)
thechars$ = thechars$ & Mid(Items, DoList, 1)

If Mid(Items, DoList, 1) = "," Then
List.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(Items, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList
End Sub

Sub AntiIdle()
'Jaguar:
'This is a sub that finds the AOL Modal window
' "You've been idle for a while"...blah blah blah
'and closes it for you.
AOL% = FindWindow("_AOL_Modal", vbNullString)
xstuff% = FindChildByTitle(AOL%, "Favorite Places")
If xstuff% Then Exit Sub
xstuff2% = FindChildByTitle(AOL%, "File Transfer *")
If xstuff2% Then Exit Sub
yes% = FindChildByClass(AOL%, "_AOL_Button")
APIClickButton yes%
End Sub

Sub AOLChangeCaption(NewCaption As String)
'Jaguar:
'This changes the "America  Online" to whatever
'you change newcaption to
'Shadow:
'I fixed AOLGetUser so that it gets the screenName
'when this has been changed
Call APISetText(AOLWindow(), NewCaption)
End Sub

Sub AOLChatSend(Text As String)
'Jaguar:
'sends "Text" to the chat room
room% = AOLFindRoom()
roombox% = FindChildByClass(room%, "RICHCNTL")
For I = 1 To 6
roombox% = APINextWin(roombox%)
Next I
Call APISetText(roombox%, Text)
DoEvents
roombox% = APINextWin(roombox%)
APIClickIcon roombox%
End Sub

Function AOLCheckIMs(Person As String) As Boolean
'Shadow:
'if the given Person can receive IMs, returns true
'if not, returns false
If AOLVersion = 3 Then
 Call RunMenuByString(AOLWindow(), "Send an Instant Message")
ElseIf AOLVersion = 4 Then
 Call AOLKeyword("aol://9293:" + Person)
 IMWin% = 0
 While IMWin% = 0: DoEvents
  IMWin% = FindChildByTitle(AOLMDI, "Send Instant Message")
 Wend
End If

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call APISetText(aoledit%, Person)
imsend% = FindChildByClass(im%, "_AOL_Icon")

For sends = 1 To 10
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends

APIClickIcon (imsend%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
OkWinStatic% = FindChildByTitle(OkWin%, Person)
OkWinText$ = APIGetText(OkWinStatic%)
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
If InStr(OkWinText$, "able") Then
 AOLCheckIMs = True
Else
 AOLCheckIMs = False
End If
End Function

Function AOLCheckIMs2(Person As String) As String
'Shadow:
'returns the text in the available for IMs box
If AOLVersion = 3 Then
 Call RunMenuByString(AOLWindow(), "Send an Instant Message")
ElseIf AOLVersion = 4 Then
 Call AOLKeyword("aol://9293:" + Person)
 IMWin% = 0
 While IMWin% = 0: DoEvents
  IMWin% = FindChildByTitle(AOLMDI, "Send Instant Message")
 Wend
End If

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call APISetText(aoledit%, Person)
imsend% = FindChildByClass(im%, "_AOL_Icon")

For sends = 1 To 10
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends

APIClickIcon (imsend%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
IMWin% = FindChildByTitle(MDI%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
OkWinStatic% = FindChildByTitle(OkWin%, Person)
OkWinText$ = APIGetText(OkWinStatic%)
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop
AOLCheckIMs2 = OkWinText$
End Function

Sub AOLClearChat()
'Shadow:
'This clears the chat room on AOL 4
Dim p As String
Dim p2 As String
If AOLVersion <> 4 Then
 Unavailable
 Exit Sub
End If
p = ".<p=                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 <br>"
p2 = ".<pre=                                                                                                                                                                                                                         <br> "
AOLChatSend ("<FONT COLOR=" & Chr$(34) & "#FFFFFE" & Chr$(34) & ">" & " " + p + "" + p2 + "" + p2 + "")
Pause 0.1
AOLChatSend ("<FONT COLOR=" & Chr$(34) & "#FFFFFE" & Chr$(34) & ">" & " " + p + "" + p2 + "" + p2 + "")
Pause 0.1
AOLChatSend ("<FONT COLOR=" & Chr$(34) & "#FFFFFE" & Chr$(34) & ">" & " " + p + "" + p2 + "" + p2 + "")
Pause 0.1
AOLChatSend ("<FONT COLOR=" & Chr$(34) & "#FFFFFE" & Chr$(34) & ">" & " " + p + "" + p2 + "" + p2 + "")
End Sub

Function AOLCountMail() As Long
'Jaguar:
'to use this properly, use it such as
'Msgbox "You have " & AOLCountMail & " mails"
themail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(themail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Sub AOLCursor()
'Jaguar:
'returns the hourglass cursor to the arrow cursor
'Shadow:
'this is the same as kill wait
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
SendMessage FindWindow("_AOL_Modal", vbNullString), WM_CLOSE, 0, 0
End Sub

Function AOLFindRoom()
'Jaguar:
'sets focus on the chat room window
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
childfocus% = GetWindow(MDI%, 5)

While childfocus%
If AOLVersion = 3 Then
 listers% = FindChildByClass(childfocus%, "_AOL_Edit")
 listere% = FindChildByClass(childfocus%, "_AOL_View")
ElseIf AOLVersion = 4 Then
 listers% = FindChildByClass(childfocus%, "RICHCNTL")
 listere% = FindChildByClass(childfocus%, "RICHCNTL")
End If
listerb% = FindChildByClass(childfocus%, "_AOL_Listbox")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then AOLFindRoom = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, GW_HWNDNEXT)
Wend
End Function

Function AOLGetChat() As String
'Jaguar:
'this gathers the text in the chat room window
'also can be used to make sure a user is in a
'chat room.  ex:  If AOLGetChat() = 0 Then user is
'not in a chat room
childs% = AOLFindRoom()
If AOLVersion = 3 Then
 child% = FindChildByClass(childs%, "_AOL_View")
ElseIf AOLVersion = 4 Then
 child% = FindChildByClass(childs%, "RICHCNTL")
End If
AOLGetChat = APIGetText(child%)
End Function

Function AOLGetChatName() As String
'Shadow:
'gets the Name of the current chat room
AOLGetChatName = APIGetText(AOLFindRoom)
End Function

Function AOLGetList(Index As Long, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

room% = AOLFindRoom()
aolhandle = FindChildByClass(room%, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function

Function AOLGetListString(Parent, Index, Buffer As String)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
    

aolhandle = Parent

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, Person$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal Person$, 4)
ListPersonHold = ListPersonHold + 6

Person$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, Person$, Len(Person$), ReadBytes)

Person$ = Left$(Person$, InStr(Person$, vbNullChar) - 1)
Call CloseHandle(AOLProcessThread)
End If

Buffer$ = Person$
End Function

Sub AOLGetMail(ListBox As ListBox)
'Shadow:
'this will read all the mail in the mailbox and copy it
'to the given listbox
Destination% = ListBox.hWnd
Source% = FindChildByClass(FindChildByClass(FindChildByClass(FindChildByTitle(AOLMDI, AOLGetUser & "'s Online Mailbox"), "_AOL_TabControl"), "_AOL_TabPage"), "_AOL_Tree")
counts = SendMessage(Source%, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source%, LB_GETTEXT, Adding, Buffer$)
char9 = InStr(Buffer$, Chr(9))
char9 = InStr(char9 + 1, Buffer$, Chr(9))
Buffer$ = Right(Buffer$, Len(Buffer$) - char9)
addstrings% = SendMessageByString(Destination%, LB_ADDSTRING, 0, Buffer$)
Next Adding
End Sub

Sub AOLGetMemberProfile(Person As String)
'Jaguar:
'This gets the profile of member "Person"
'Shadow v1:
'i couldn't convert this to work on AOL 4 too
'because it uses the popupmenus
'Shadow v2:
'AOL 4 proggers watch out, that code is buggy
If AOLVersion = 3 Then
 AOLRunMenuByString "Get a Member's Profile"
ElseIf AOLVersion = 4 Then
 AOLRunPopupMenu People, 14
End If
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
prof% = FindChildByTitle(MDI%, "Get a Member's Profile")
putPerson% = FindChildByClass(prof%, "_AOL_Edit")
Call APISetText(putPerson%, Person)
okbutton% = FindChildByClass(prof%, "_AOL_Icon")
APIClickIcon okbutton%
End Sub

Function AOLGetTopWindow()
'Jaguar:
'gets the topmost window
AOLGetTopWindow = GetTopWindow(AOLMDI())
End Function

Function AOLGetUser() As String
'Jaguar:
'Retrives the user's SN from the welcome window
'Shadow:
'I fixed this so that it gets the screenName
'when the AOL caption has been changed
On Error Resume Next
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
AOLGetUser = User
End Function

Sub AOLHide()
'Jaguar:
'hides the AOL window...doesn't close it
a = ShowWindow(AOLWindow, 0)
End Sub

Function AOLHyperLink(Address As String, Caption As String) As String
'Shadow:
'This creates a link to be put in an im or on AOL 4
'in the chat room.  This example shows how to make an
'im me link in the chat room:
'AOLChatSend AOLHyperLink("aol://9293:" & AOLGetUser, "IM ME")
AOLHyperLink = "< a href=" & Address & ">" & Caption & "</a>"
End Function

Function AOLIMSend(message As String) As Boolean
'Shadow:
'this sends the given Message to the first open IM
im% = FindChildByTitle(AOLMDI(), ">Instant Message ")
If im% Then GoTo CanSend
im% = FindChildByTitle(AOLMDI(), "  Instant Message ")
If im% Then GoTo CanSend
AOLIMSend = False
Exit Function
CanSend:
AOLIMSend = True
RTFBox% = FindChildByClass(im%, "RICHCNTL")
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
RTFBox% = APINextWin(RTFBox%)
Send% = APINextWin(RTFBox%)
APISetText RTFBox%, message
APIClickIcon Send%
End Function

Sub AOLIMsOff(message As String)
'Jaguar:
'Turns IM's off
Call AOLInstantMessage("$IM_OFF", message)
End Sub

Sub AOLIMsOn(message As String)
'Jaguar:
'turns IM's on
Call AOLInstantMessage("$IM_ON", message)
End Sub

Function AOLInstantMessage(Person As String, message As String) As Boolean
'Jaguar:
'sends an Instant Message to "Person" with the
'Message of "Message"
'Shadow:
'returns true if it was recieved, and false it they did
'not recieve it
If AOLVersion = 3 Then
 Call RunMenuByString(AOLWindow(), "Send an Instant Message")
ElseIf AOLVersion = 4 Then
 AOLKeyword "aol://9293:"
End If

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aoledit% = FindChildByClass(im%, "_AOL_Edit")
aolrich% = FindChildByClass(im%, "RICHCNTL")
imsend% = FindChildByClass(im%, "_AOL_Icon")
If aoledit% <> 0 And aolrich% <> 0 And imsend% <> 0 Then Exit Do
Loop

Call APISetText(aoledit%, Person)
Call APISetText(aolrich%, message)
imsend% = FindChildByClass(im%, "_AOL_Icon")

For sends = 1 To 9
imsend% = GetWindow(imsend%, GW_HWNDNEXT)
Next sends

APIClickIcon (imsend%)

AOLInstantMessage = True
Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
im% = FindChildByTitle(MDI%, "Send Instant Message")
aolcl% = FindWindow("#32770", "America Online")
If aolcl% <> 0 Then closer = SendMessage(aolcl%, WM_CLOSE, 0, 0): closer2 = SendMessage(im%, WM_CLOSE, 0, 0): AOLInstantMessage = False: Exit Do
If im% = 0 Then Exit Do
Loop
End Function

Sub AOLInvitation(People As String, message As String, Place As String, CloseWindow As Boolean)
'Shadow:
'sends an invitation with the given settings
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Buddy% = FindChildByTitle(MDI%, "Buddy List Window")
If Buddy% = 0 Then
 Call RunMenuByString(AOLWindow(), "Buddy List")
 Do: DoEvents
 Buddy% = FindChildByTitle(MDI%, "Buddy List Window")
 If Buddy% <> 0 Then Exit Do
 Loop
End If

budchat% = FindChildByTitle(Buddy%, "Buddy")
budchat% = GetWindow(budchat%, GW_HWNDNEXT)
APIClickIcon budchat%

Do: DoEvents
invit% = FindChildByTitle(MDI%, "Buddy Chat")
peeple% = FindChildByClass(invit%, "_AOL_Edit")
mesage% = GetWindow(peeple%, GW_HWNDNEXT)
plase% = GetWindow(mesage%, GW_HWNDNEXT)
Send% = FindChildByClass(invit%, "_AOL_Icon")
If peeple% <> 0 And mesage% <> 0 And plase% <> 0 And Send% <> 0 Then Exit Do
Loop

Call APISetText(peeple%, People)
Call APISetText(mesage%, message)
Call APISetText(plase%, Place)
Send% = FindChildByClass(invit%, "_AOL_Icon")

APIClickIcon (Send%)

If CloseWindow = True Then Exit Sub
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Do: DoEvents
retwin% = FindChildByTitle(MDI%, "Invitation From: ")
If retwin% <> 0 Then Exit Do
Loop

closer = SendMessage(retwin%, WM_CLOSE, 0, 0)
End Sub

Function AOLIsOnline(Notify As Boolean) As Boolean
'Jaguar:
'makes sure a user is signed on before using
'a certain feature
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
Welcome% = FindChildByTitle(MDI%, "Welcome, ")
If Welcome% = 0 Then
 AOLIsOnline = False
Else
 AOLIsOnline = True
End If
If Notify <> False And Welcome% = 0 Then
 MsgBox "Please sign on before using this feature.", 64
End If
End Function

Sub AOLKeyword(Text As String)
'Jaguar:
'goes to keyword "text"
If AOLVersion = 4 Then
 Toobar% = FindChildByClass(AOLWindow, "AOL Toolbar")
 Toobar% = FindChildByClass(Toobar%, "_AOL_Toolbar")
 GoBox% = FindChildByClass(Toobar%, "_AOL_Combobox")
 GoBox% = FindChildByClass(GoBox%, "Edit")
 APISetText GoBox%, Right(Text, Len(Text) - 1)
 SendCharNum GoBox%, Asc(Left(Text, 1))
 SendCharNum GoBox%, 13
ElseIf AOLVersion = 3 Then
 Call RunMenuByString(AOLWindow(), "Keyword...")
 Do: DoEvents
 AOL% = FindWindow("AOL Frame25", vbNullString)
 MDI% = FindChildByClass(AOL%, "MDIClient")
 keyw% = FindChildByTitle(MDI%, "Keyword")
 kedit% = FindChildByClass(keyw%, "_AOL_Edit")
 If kedit% Then Exit Do
 Loop

 editsend% = SendMessageByString(kedit%, WM_SETTEXT, 0, Text)
 pausing = DoEvents()
 Sending% = SendMessage(kedit%, 258, 13, 0)
 pausing = DoEvents()
End If
End Sub

Function AOLLastChatLine() As String
'Jaguar:
'returns the last line of chat in a chat room
theview$ = AOLGetChat


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(theview$, lastlen, Len(thechars$))
AOLLastChatLine = lastline
End Function

Function AOLLastMessage() As String
ChatText = AOLLastChatLine
ChatTrimNum = Len(AOLLastTalker)
ChatTrim$ = Mid$(ChatText, ChatTrimNum + 4, Len(ChatText) - Len(AOLLastTalker))
AOLLastMessage = ChatTrim$
End Function

Function AOLLastTalker() As String
ChatText$ = AOLLastChatLine
ChatTrim$ = Left$(ChatText$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
AOLLastTalker = SN
End Function

Sub AOLLocateMember(name As String)
'Jaguar:
'locates, if possible, member "Name"
'Shadow v1:
'i couldn't convert this to work on AOL 4 too
'because it uses the popupmenus
'Shadow v2:
'AOL 4 proggers watch out, that code is buggy
If AOLVersion <> 4 Then
 AOLRunMenuByString "Locate a Member Online"
Else
 AOLRunPopupMenu People, 13
End If
DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
prof% = FindChildByTitle(MDI%, "Locate Member Online")
putName% = FindChildByClass(prof%, "_AOL_Edit")
Call APISetText(putName%, name)
okbutton% = FindChildByClass(prof%, "_AOL_Icon")
APIClickIcon okbutton%
closes = SendMessage(prof%, WM_CLOSE, 0, 0)
End Sub

Sub AOLMail(People As String, Subject As String, message As String)
'Jaguar:
'opens a blank mail and sends it to "People" with
'the subject of "subject" and body of "Message"
AOLRunTool 1

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
icone% = FindChildByClass(mailwin%, "_AOL_Icon")
peepz% = FindChildByClass(mailwin%, "_AOL_Edit")
subjt% = FindChildByTitle(mailwin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(mailwin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, People)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, Subject)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)

APIClickIcon (icone%)

Do: DoEvents
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
mailwin% = FindChildByTitle(MDI%, "Compose Mail")
erro% = FindChildByTitle(MDI%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If mailwin% = 0 Then Exit Do
If aolw% <> 0 Then
'a = SendMessage(aolw%, WM_CLOSE, 0, 0)
APIClickButton (FindChildByTitle(aolw%, "OK"))
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(mailwin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Public Sub AOLMassIM(List As ListBox, Text As String)
'Jaguar:
'This was made by DouBT
'The one that was already here was all screwed up!
If List.ListCount = 0 Then
 MsgBox "List is empty.", vbExclamation
 Exit Sub
End If
For I% = 0 To List.ListCount - 1
Call AOLInstantMessage(List.List(I%), Text)
Next I%
End Sub

Function AOLMDI()
'Jaguar:
'this can be used instead of typing out the two
'lines of code below
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL%, "MDIClient")
End Function

Function AOLMessageFromIM() As String
im% = FindChildByTitle(AOLMDI(), ">Instant Message ")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message ")
If im% Then GoTo Greed
Exit Function
Greed:
imtext% = FindChildByClass(im%, "RICHCNTL")
IMMessage = APIGetText(imtext%)
SN = AOLWindow()
snlen = Len(AOLWindow()) + 3
blah = Mid(IMMessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(IMMessage, Len(IMMessage) - 1) 'Left(blah, Len(blah) - 1)
End Function

Sub AOLOpenMailBox(Box As MailBox)
'Shadow:
'opens given mailbox
If AOLVersion = 4 Then
 If Box = 1 Then
  Call AOLRunTool(0)
 ElseIf Box = 2 Then
  Call AOLRunPopupMenu(MailCenter, 4)
 Else
  Call AOLRunPopupMenu(MailCenter, 5)
 End If
Else
 If Box = 1 Then
  Call AOLRunMenuByString("Read &New Mail")
 ElseIf Box = 2 Then
  Call AOLRunMenuByString("Check Mail You've &Read")
 Else
  Call AOLRunMenuByString("Check Mail You've &Sent")
 End If
End If
End Sub

Sub AOLPhish(List As ListBox, Text As String)
'Shadow:
'this sends a mass im to all the people in the listbox
'with your Message, but it makes it look like the online
'host says it
AOLMassIM List, "<BR><BR><BR><BR><BR><FONT COLOR=#0000FF SIZE=2><B> OnlineHost:  </B></FONT><FONT COLOR=#000000>" & Text
End Sub

Function AOLReplaceWithSN(Text As String) As String
'Jaguar:
'will turn "*" in a string into the current
'user's Screen Name
'Example:  current user's SN is GreedieFly
'it will turn "* is da bomb!" into
'"GreedieFly is da bomb!"
Let inptxt$ = Text
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextChr$ = "*" Then Let nextChr$ = AOLGetUser()

Let newsent$ = newsent$ + nextChr$

If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
AOLReplaceWithSN = newsent$
End Function

Function AOLRespondIM(message As String) As String
'Jaguar:
'This finds an IM sent to you, answers it with a
'Message of "Message", sends it and then closes the
'IM window
'Shadow32:
'i made it return the Message inside the IM
im% = FindChildByTitle(AOLMDI(), ">Instant Message ")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message ")
If im% Then GoTo Greed
Exit Function
Greed:
E% = FindChildByClass(im%, "RICHCNTL")
AOLRespondIM = APIGetText(E%)
E% = GetWindow(E%, GW_HWNDNEXT)
E% = GetWindow(E%, GW_HWNDNEXT)
E% = GetWindow(E%, GW_HWNDNEXT)
E% = GetWindow(E%, GW_HWNDNEXT)
E% = GetWindow(E%, GW_HWNDNEXT)
E% = GetWindow(E%, GW_HWNDNEXT)
E% = GetWindow(E%, GW_HWNDNEXT)
E% = GetWindow(E%, GW_HWNDNEXT)
e2% = GetWindow(E%, GW_HWNDNEXT) 'Send Text
E% = GetWindow(e2%, GW_HWNDNEXT) 'Send Button
Call APISetText(e2%, message)
APIClickIcon (E%)
Pause 0.8
DoEvents
APIClose im%
End Function

Function AOLRoomCount() As Integer
'Jaguar:
'returns the number of people in the chatroom
TheChild% = AOLFindRoom()
lister% = FindChildByClass(TheChild%, "_AOL_Listbox")

getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
AOLRoomCount = getcount
End Function

Sub AOLRunMenuByString(stringer As String)
'Jaguar:
'This will run the drop down menus.
'To use this you have to type it exactly as it is
'on the drop down menus.  Such as:
'if you wanted to click the compose mail in the
'drop down menu under mail you would put
'AOLRunMenuByString("&Compose Mail")
'you must put an & before the letter that is
'underlined
Call RunMenuByString(AOLWindow(), stringer)
End Sub

Sub AOLRunPopupMenu(HIndex As PopupMenu, VIndex As Integer)
'Shadow:
'This will run the popup menus on AOL 4.  Sorry I could
'not do run by string, but that would be very hard if
'possible.  It is also a little buggy so, be careful.
'Note: On the Indexes, they start with, 0 not 1, so 2 is
'mail center...  Also remember to count separator lines
'on menus as an Index number
If AOLVersion <> 4 Then
 Unavailable
 Exit Sub
End If

toolbar% = FindChildByClass(AOLWindow, "AOL Toolbar")
toolbar% = FindChildByClass(toolbar%, "_AOL_Toolbar")
hIcon% = FindChildByClass(toolbar%, "_AOL_Icon")
For X = 0 To HIndex - 1
hIcon% = GetWindow(hIcon%, GW_HWNDNEXT)
Next X
PostMessageBynum hIcon, WM_LBUTTONDOWN, 0, 0&
PostMessageBynum hIcon, WM_LBUTTONUP, 0, 0&

While hMenu% = 0: DoEvents
 hMenu% = FindWindow("#32768", vbNullString)
Wend

SendMessage AOLWindow, WM_COMMAND, VIndex + 7169, 0
BringWindowToTop FindWindow("Shell_TrayWnd", vbNullString)
End Sub

Sub AOLRunTool(Tool As Integer)
'Jaguar:
'this clicks on the toolbar icons
'the first one...mailbox...is 0
'compose mail is 1
'channels is 2 etc...
toolbar% = FindChildByClass(AOLWindow(), "AOL Toolbar")
If AOLVersion = 4 Then toolbar% = FindChildByClass(toolbar%, "_AOL_Toolbar")
iconz% = FindChildByClass(toolbar%, "_AOL_Icon")
For X = 1 To Tool - 1
iconz% = GetWindow(iconz%, GW_HWNDNEXT)
Next X
isen% = IsWindowEnabled(iconz%)
If isen% = 0 Then Exit Sub
APIClickIcon (iconz%)
End Sub

Sub AOLSetBackPre()
'Shadow v1:
'i couldn't convert this to work on AOL 4 too
'because it uses the popupmenus
'Shadow v2:
'AOL 4 proggers watch out, that code is buggy
If AOLVersion <> 4 Then
 RunMenuByString AOLWindow, "Preferences"
Else
 AOLRunPopupMenu MyAOL, 3
End If

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
APIClickIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
APICloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByClass(aolmod%, "_AOL_Icon")
If aolOK% <> 0 And APICloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(APICloses%, BM_SETCHECK, 0, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

APIClickIcon (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)
End Sub

Sub AOLSetFocus()
'Jaguar:
'SetFocusAPI doesn't work AOL because AOL has added
'a safeguard against other programs calling certain
'API functions (like owner-drawn things and like.)
'This is the only way known for setting the focus
'to AOL.  This is a normal VB command!
'Shadow:
'Well i found something that does work
BringWindowToTop AOLWindow
End Sub

Sub AOLSetPreference()
'Shadow v1:
'i couldn't convert this to work on AOL 4 too
'because it uses the popupmenus
'Shadow v2:
'AOL 4 proggers watch out, that code is buggy
If AOLVersion <> 4 Then
 RunMenuByString AOLWindow, "Preferences"
Else
 AOLRunPopupMenu MyAOL, 3
End If

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

Pause (0.2)
APIClickIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
APICloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByClass(aolmod%, "_AOL_Icon")
If aolOK% <> 0 And APICloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(APICloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

APIClickIcon (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)
End Sub

Sub AOLSignOff()
AOLRunMenuByString "&Sign Off"
End Sub

Function AOLSNFromIM() As String
'Jaguar:
'this will return a Screen Name from an IM
im% = FindChildByTitle(AOLMDI(), ">Instant Message ")
If im% Then GoTo Greed
im% = FindChildByTitle(AOLMDI(), "  Instant Message ")
If im% Then GoTo Greed
Exit Function
Greed:
heh$ = APIGetText(im%)
naw$ = Mid(heh$, InStr(heh$, ":") + 2)
AOL = naw$
End Function

Sub AOLUnHide()
'Jaguar:
'Unhides AOL window after it was hidden
a = ShowWindow(AOLWindow, 5)
End Sub

Sub AOLUnUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(AOL%, 0)
End Sub

Sub AOLUpchat()
AOL% = FindWindow("AOL Frame25", vbNullString)
AOModal% = FindChildByClass(AOL%, "_AOL_Modal")
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(AOL%, 1)
Call EnableWindow(Upp%, 0)
End Sub

Function AOLVersion()
'Jaguar:
'returns What version the User is using
AOL% = FindWindow("AOL Frame25", vbNullString)
hMenu% = GetMenu(AOL%)

submenu% = GetSubMenu(hMenu%, 0)
subItem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")

FindString% = GetMenuString(submenu%, subItem%, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("Print Set&up...") & "*" Then
AOLVersion = 4
ElseIf UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 2.5
End If
End Function

Sub AOLWaitMail()
'Jaguar:
'this waits until the user's mail window has stopped
'listing mail
'Shadow:
'in theory this works, but in reality, not so often
mailwin% = GetTopWindow(AOLMDI())
AOLTree% = FindChildByClass(mailwin%, "_AOL_Tree")

Do: DoEvents
firstcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
Pause (3)
secondcount = SendMessage(AOLTree%, LB_GETCOUNT, 0, 0)
If firstcount = secondcount Then Exit Do
Loop
End Sub

Function AOLWelcome()
AOL% = FindWindow("AOL Frame25", vbNullString)
MDI% = FindChildByClass(AOL%, "MDIClient")
AOLWelcome = FindChildByTitle(MDI%, "Welcome, ")
End Function

Function AOLWindow()
'Jaguar:
'finds the AOL window
AOL% = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL%
End Function

Function APIChildWin(hWnd%)
'Shadow:
'this gets the first child window
APIChildWin = GetWindow(hWnd%, GW_CHILD)
End Function

Public Sub APIClickButton(hWnd%)
'Jaguar:
'clicks a button
SendMessage hWnd%, WM_KEYDOWN, VK_SPACE, 0
SendMessage hWnd%, WM_KEYUP, VK_SPACE, 0
End Sub

Sub APIClickIcon(hWnd%)
'Jaguar:
'clicks an icon....such as the toolbar buttons
Click% = SendMessage(hWnd%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(hWnd%, WM_LBUTTONUP, 0, 0&)
End Sub

Function APIClickList(hWnd%)
ClickList% = SendMessageByNum(hWnd%, &H203, 0, 0&)
End Function

Sub APIClose(hWnd%)
'Jaguar:
'closes the hWnd window...same as clicking the X
closes = SendMessage(hWnd%, WM_CLOSE, 0, 0)
End Sub

Function APIGetClass(hWnd%) As String
Buffer$ = String$(250, 0)
getclas% = GetClassName(hWnd%, Buffer$, 250)

APIGetClass = Buffer$
End Function

Function APIGetCursorXPos() As Long
'Shadow:
'This returns the X position of the mouse on the screen.
'Remember, this is pixels, not twips.  To convert between
'twips and pixels do this:
'Pixels = Twips * Screen.TwipsPerPixelX
'Twips = Pixels / Screen.TwipsPerPixelX
'Or you can just change your forms scalemode to pixels and
'there will be no conversion.
'Note: I've never seen anything but 15 twips per pixel,
'but if you're a perfectionist you can use that code.
Dim CurPos As POINTAPI
GetCursorPos CurPos
APIGetCursorXPos = CurPos.X
End Function

Function APIGetCursorYPos() As Long
'Shadow:
'This returns the Y position of the mouse on the screen.
Dim CurPos As POINTAPI
GetCursorPos CurPos
APIGetCursorYPos = CurPos.Y
End Function

Function APIGetParent(hWnd%)
'Shadow:
'this gets the parent window
APIGetParent = GetWindow(hWnd%, GW_OWNER)
End Function

Function APIGetText(hWnd%) As String
'Jaguar:
'this will get the text of the hWnd window
GetTrim = SendMessageByNum(hWnd%, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(hWnd%, 13, GetTrim + 1, TrimSpace$)

APIGetText = TrimSpace$
End Function

Function APINextWin(hWnd%)
'Shadow:
'this gets the next window
APINextWin = GetWindow(hWnd%, GW_HWNDNEXT)
End Function

Function APIPrevWin(hWnd%)
'Shadow:
'this gets the previous window
APIPrevWin = GetWindow(hWnd%, GW_HWNDPREV)
End Function

Sub APISetCursorPos(X As Long, Y As Long)
'Shadow:
'This will set the mouse cursor position, but remember
'that it's pixels, not twips.
SetCursorPos X, Y
End Sub

Sub APISetFocus(hWnd%)
'Shadow:
'sets the focus to the given window
SetFocusAPI hWnd%
End Sub

Sub APISetParent(hWnd%, Parent%)
'Shadow:
'this can be used to make your form an MDI Child
'To make it an AOL Child do this:
'ParentChange(Form1.hWnd, AOLMDI())
doparent% = SetParent(hWnd%, Parent%)
End Sub

Sub APISetText(hWnd%, Text As String)
'Jaguar:
'this will put "txt" in the window of "win"
'this can be used to change the text in _AOL_Static,
'RICHCNTL and _AOL_Edit windows or the Window caption
thetext% = SendMessageByString(hWnd%, WM_SETTEXT, 0, Text)
End Sub

Function APISpy()
'Shadow:
'This is for making a spy prog. It will return the window
'that the mouse is over when it's called.
'Note: use APIGetClass, APIGetText, APIGetFileName, and
'APIGetParent for other things on a spy prog
APISpy = WindowFromPoint(APIGetCursorXPos, APIGetCursorYPos)
End Function

Sub CenterForm(Form As Form, CenterX As Boolean, CenterY As Boolean)
'Jaguar:
'this will center you form in the very center of
'the users screen
'Shadow:
'i updated this so you can center the form only
'horizontally or only vertically
With Form
 If CenterX Then .Left = (Screen.Width - .Width) / 2
 If CenterY Then .Top = (Screen.Height - .Height) / 2
End With
End Sub

Function DescrambleText(Text As String)
'Jaguar:
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(Text, Len(Text), 1)

If Not findlastspace = " " Then
Text = Text & " "
Else
Text = Text
End If

'Descrambles the text
For scrambling = 1 To Len(Text)
thechar$ = Mid(Text, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
Chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(Chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo city
lastchar$ = Mid(Chars$, 2, 1)
'finds what is inbetween the last and first character
midchar$ = Mid(Chars$, 3, Len(Chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
city:
Scrambled$ = Scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
Scrambled$ = Scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed Buffers
sniff:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = Scrambled$

End Function

Public Sub DisableCTRL_ALT_DEL()
'Jaguar:
'Disables the Ctrl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Public Sub EnableCTRL_ALT_DEL()
'Jaguar:
'Enables the Ctrl+Alt+Del
 Dim ret As Integer
 Dim pOld As Boolean
 ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Function EncryptType(Text As String, types As Integer)
'Jaguar:
'to encrypt, example:
'encrypted$ = EncryptType("Messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = EncryptType("decryptedMessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(Text)
If types = 0 Then
current$ = Asc(Mid(Text, God, 1)) - 1
Else
current$ = Asc(Mid(Text, God, 1)) + 1
End If
Process$ = Process$ & Chr(current$)
Next God

EncryptType = Process$
End Function

Function FindChildByClass(ParentWnd%, Class As String)
firs% = GetWindow(ParentWnd%, GW_MAX)
If UCase(Mid(APIGetClass(firs%), 1, Len(Class))) Like UCase(Class) Then GoTo Greed
firs% = GetWindow(ParentWnd%, GW_CHILD)
If UCase(Mid(APIGetClass(firs%), 1, Len(Class))) Like UCase(Class) Then GoTo Greed

While firs%
firss% = GetWindow(ParentWnd%, GW_MAX)
If UCase(Mid(APIGetClass(firss%), 1, Len(Class))) Like UCase(Class) Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(Mid(APIGetClass(firs%), 1, Len(Class))) Like UCase(Class) Then GoTo Greed
Wend
FindChildByClass = 0

Greed:
room% = firs%
FindChildByClass = room%
End Function

Function FindChildByTitle(ParentWnd%, Title As String)
firs% = GetWindow(ParentWnd%, 5)
If UCase(APIGetText(firs%)) Like UCase(Title) Then GoTo Greed
firs% = GetWindow(ParentWnd%, GW_CHILD)

While firs%
firss% = GetWindow(ParentWnd%, 5)
If UCase(APIGetText(firss%)) Like UCase(Title) & "*" Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(APIGetText(firs%)) Like UCase(Title) & "*" Then GoTo Greed
Wend
FindChildByTitle = 0

Greed:
room% = firs%
FindChildByTitle = room%
End Function

Function FindFwdWin(dosloop)
'Jaguar:
'FindFwdWin = GetParent(FindChildByTitle(FindChildByClass(AOLMDI(), "AOL Child"), "Forward"))
'Exit Function
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo Greed
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, GW_HWNDNEXT)
forw% = FindChildByTitle(firs%, "Forward")
If forw% <> 0 Then GoTo Greed
If dosloop = 1 Then Exit Do
Loop
Exit Function
Greed:
FindFwdWin = firs%

Exit Function
begis:
FindFwdWin = firss%
End Function


Function FindSendWin(dosloop)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo Greed
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo Greed
If dosloop = 1 Then Exit Do
Loop
Exit Function
Greed:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function

Function FreeProcess()
'Jaguar:
'frees process of freezes in your program
'and other stuff that makes your program
'slow down.  Works great.
'Shadow:
'this works good for making punters go faster
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Public Function GetChildCount(hWnd%) As Long
Dim hChild As Long

Dim I As Integer
   
If hWnd% = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd%, GW_CHILD)
   

While hChild
hChild = GetWindow(hChild, GW_HWNDNEXT)
I = I + 1
Wend

GetChildCount = I
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Function GetLineCount(Text As String)

theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)

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

Function GetListIndex(List As ListBox, Text As String) As Long

Dim iIndex As Integer

With List
 For iIndex = 0 To .ListCount - 1
   If .List(iIndex) = Text Then
    GetListIndex = iIndex
    Exit Function
   End If
 Next iIndex
End With

GetListIndex = -2   '  if Item isnt found
'( I didnt want to use -1 as it evaluates to True)

End Function

Sub HideWindow(hWnd%)
'Jaguar:
'hides the "hWnd" window
Hi = ShowWindow(hWnd%, SW_HIDE)
End Sub

Function IntegerToString(Number As Integer) As String
IntegerToString = str(Number)
End Function

Function KTEncrypt(ByVal password, ByVal strng, force%)
'Jaguar:
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(password) = 0 Then Error 31100
  
  'Is password too long
  If Len(password) > 255 Then Error 31100

  'Is there a strng$ to work with?
  If Len(strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
          strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
        End If
        look = look + 1
      Loop
      
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
    End If
  Else
    'force% flag set, ecrypt string regardless of tag
    EncryptFlag% = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(password)
  
  
  'Tack on leading characters to prevent repetative recognition
  password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
  password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
  password = password + Chr$(Asc(Right$(password, 1)) Xor PassMax)
  password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(tochange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
      look = InStr(look, strng, Chr$(1))
      If look > 0 Then
        strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
      look = InStr(strng, Chr$(0))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
      look = InStr(strng, Chr$(10))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
      look = InStr(strng, Chr$(13))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
      look = InStr(strng, Chr$(26))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      strng = Mid$(strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function
End Function

Function LineFromText(Text As String, Line As Long) As String
theview$ = Text


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
C = C + 1
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
If Line = C Then GoTo ex
thechars$ = ""
End If

Next FindChar
Exit Function
ex:
thechatext$ = ReplaceCharacter(thechatext$, Chr(13), "")
thechatext$ = ReplaceCharacter(thechatext$, Chr(10), "")
LineFromText = thechatext$
End Function

Sub ListToList(Source%, Destination%)
'Shadow:
'Jaguar had this in there and its alright for many
'things, but some listboxes, like the people here
'listbox in a chatroom, don't like being manipulated by
'API functions and will give you an AOL error Message.
counts = SendMessage(Source%, LB_GETCOUNT, 0, 0)

For Adding = 0 To counts - 1
Buffer$ = String$(250, 0)
getstrings% = SendMessageByString(Source%, LB_GETTEXT, Adding, Buffer$)
addstrings% = SendMessageByString(Destination%, LB_ADDSTRING, 0, Buffer$)
Next Adding
End Sub

Function MakeSpaceInGoto(Text As String)
'Jaguar:
'this is for Room Busters.  this will put
'a %20 for a space in the goto menu or keyword
'to make sure if the user puts in "M M" as the room
'Name that the user will end up in "M M" and not
'"MM"
'Shadow:
'i don't see the point for this because "M M" is the
'same room as "MM", just displayed differently
Let inptxt$ = Text
Let lenth% = Len(inptxt$)

Do While numspc% <= lenth%
DoEvents
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextchrr$ = Mid$(inptxt$, numspc%, 2)
If nextChr$ = " " Then Let nextChr$ = "%20"
Let newsent$ = newsent$ + nextChr$

If Crapp% > 0 Then Let Crapp% = Crapp% - 1
DoEvents
Loop
MakeSpaceInGoto = newsent$
End Function

Sub MaxWindow(hWnd%)
'Jaguar:
'makes "hWnd" window Maximized
ma = ShowWindow(hWnd%, SW_MAXIMIZE)
End Sub

Sub MiniWindow(hWnd%)
'Jaguar:
'minimizes the "hWnd" window
mi = ShowWindow(hWnd, SW_MINIMIZE)
End Sub

Sub MoveForm(Form As Form)
'Shadow:
'moves a form and shows the little trails
DoEvents
ReleaseCapture
SendMessage Form.hWnd, &HA1, 2, 0
End Sub

Sub NotOnTop(Form As Form)
SetWinOnTop = SetWindowPos(Form.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
End Sub

Sub Pause(Interval)
'Jaguar:
'pause/waits for "interval" seconds
current = Timer
Do While Timer - current < Val(Interval)
DoEvents
Loop
End Sub

Sub Playwav(file As String)
'Jaguar:
'will play a .wav file.
'example:  Playwav("fileName.wav")
SoundName$ = file
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   X% = sndPlaySound(SoundName$, wFlags%)
End Sub

Function RandomNumber(Range As Long) As Long
Randomize
RandomNumber = Int((Val(Range) * Rnd) + 1)
End Function

Function ReplaceCharacter(Text As String, Find As String, Change As String) As String
'Shadow:
'Replaces a single character in a string to what you say
If InStr(Text, Find) = 0 Then
ReplaceCharacter = Text
Exit Function
End If

For Replace = 1 To Len(Text)
thechar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Find Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + Change
End If
Next Replace

ReplaceCharacter = thechars$
End Function

Function ReplaceText(Text As String, Find As String, Change As String) As String
'Shadow:
'Replaces text in a string to what you say
a = InStr(Text, Find)
If a = 0 Then
ReplaceText = Text
Else
ReplaceText = Left(Text, a - 1) + Change + Right(Text, Len(Text) - a - Len(Find) + 1)
End If
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

Function ScrambleText(Text As String) As String
'Jaguar:
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(Text, Len(Text), 1)

If Not findlastspace = " " Then
Text = Text & " "
Else
Text = Text
End If

'Scrambles the text
For scrambling = 1 To Len(Text)
thechar$ = Mid(Text, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
Chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(Chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(Chars$, Len(Chars$), 1)

'finds what is inbetween the last and first character
midchar$ = Mid(Chars$, 2, Len(Chars$) - 2)
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

'clears character and reversed Buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
ScrambleText = Scrambled$

Exit Function
End Function

Sub SendCharNum(hWnd%, Char As String)
'Shadow:
'this sends the first character of the given string to the
'given window. use this to type stuff in a text box, etc.
E = SendMessageByNum(hWnd%, WM_CHAR, Asc(Char), 0)
End Sub

Sub SizeFormToWindow(Form As Form, hWnd%)
'Jaguar:
'this will make your form(Form) into the exact size
'of the given window(hWnd%)
'example:  SizeFormToWindow Me, AOLMDI()
'that would make a very large window
Dim wndRect As RECT, lRet As Long

lRet = GetWindowRect(hWnd%, wndRect)

With Form
  .Top = wndRect.Top * Screen.TwipsPerPixelY
  .Left = wndRect.Left * Screen.TwipsPerPixelX
  .Height = ((wndRect.Bottom) - (wndRect.Top)) * Screen.TwipsPerPixelY
  .Width = ((wndRect.Right) - (wndRect.Left)) * Screen.TwipsPerPixelX
End With
End Sub

Sub StayOnline()
'Jaguar:
'this finds that 46 min box and closes it whenever
'it pops up
hWndz% = FindWindow("_AOL_Palette", "America Online Timer")
ChildhWnd% = FindChildByTitle(hWndz%, "OK")
APIClickButton (ChildhWnd%)
End Sub

Sub StayOnTop(Form As Form)
'Jaguar:
'sets your form to be Form topmost window all Form
'time. Example:  StayOnTop Me
SetWinOnTop = SetWindowPos(Form.hWnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
End Sub

Function StringToInteger(Text As String) As Integer
StringToInteger = Val(Text)
End Function

Function TrimCharacter(Text As String, Chars As String) As String
TrimCharacter = ReplaceCharacter(Text, Chars, "")
End Function

Function TrimReturns(Text As String) As String
takechr13$ = ReplaceCharacter(Text, Chr$(13), "")
takechr10$ = ReplaceCharacter(takechr13$, Chr$(10), "")
TrimReturns = takechr10$
End Function

Function TrimSpaces(Text As String) As String
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

Function TrimText(Text As String, Chars As String) As String
'Shadow:
'Removes text from a string
ReplaceText Text, Chars, ""
End Function

Function TXTBackwards(Text As String) As String
'Jaguar:
'Returns the string backwards
Let inptxt$ = Text
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let newsent$ = nextChr$ & newsent$
Loop
TXTBackwards = newsent$
End Function

Function TXTElite(Text As String) As String
'Jaguar:
'Returns the string elite
Let inptxt$ = Text
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

If nextChr$ = "A" Then Let nextChr$ = "/\"
If nextChr$ = "a" Then Let nextChr$ = "å"
If nextChr$ = "B" Then Let nextChr$ = "ß"
If nextChr$ = "C" Then Let nextChr$ = "Ç"
If nextChr$ = "c" Then Let nextChr$ = "¢"
If nextChr$ = "D" Then Let nextChr$ = "Ð"
If nextChr$ = "d" Then Let nextChr$ = "ð"
If nextChr$ = "E" Then Let nextChr$ = "Ê"
If nextChr$ = "e" Then Let nextChr$ = "è"
If nextChr$ = "f" Then Let nextChr$ = "ƒ"
If nextChr$ = "H" Then Let nextChr$ = "|-|"
If nextChr$ = "I" Then Let nextChr$ = "‡"
If nextChr$ = "i" Then Let nextChr$ = "î"
If nextChr$ = "k" Then Let nextChr$ = "|‹"
If nextChr$ = "L" Then Let nextChr$ = "£"
If nextChr$ = "M" Then Let nextChr$ = "]V["
If nextChr$ = "m" Then Let nextChr$ = "^^"
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
If nextChr$ = "W" Then Let nextChr$ = "VV"
If nextChr$ = "w" Then Let nextChr$ = "vv"
If nextChr$ = "X" Then Let nextChr$ = "X"
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
TXTElite = newsent$
End Function

Function TXTHacker(Text As String) As String
'Jaguar:
'Returns the Text hacker style
Let inptxt$ = Text
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
If nextChr$ = " " Then Let nextChr$ = " "
Let newsent$ = newsent$ + nextChr$
Loop
TXTHacker = newsent$
End Function

Function TXTSpaced(Text As String) As String
'Jaguar:
'Returns the Text spaced
Let inptxt$ = Text
Let lenth% = Len(inptxt$)
Do While numspc% <= lenth%
Let numspc% = numspc% + 1
Let nextChr$ = Mid$(inptxt$, numspc%, 1)
Let nextChr$ = nextChr$ + " "
Let newsent$ = newsent$ + nextChr$
Loop
TXTSpaced = newsent$
End Function

Sub Unavailable()
MsgBox "This feature does not work for your version of AOL.", vbInformation
End Sub

Sub UnHideWindow(hWnd%)
'Jaguar:
'unhides the "hWnd" window that has been hidden
un = ShowWindow(hWnd%, SW_SHOW)
End Sub

Function UntilWindowClass(parentw, childhand)
GoBack:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(Mid(APIGetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(APIGetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(APIGetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(Mid(APIGetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Greed
Wend
GoTo GoBack
FindClassLike = 0

Greed:
room% = firs%
UntilWindowClass = room%
End Function

Function UntilWindowTitle(parentw, childhand)
GoBac:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(APIGetText(firs%)) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(APIGetText(firss%)) Like UCase(childhand) Then GoTo Greed
firs% = GetWindow(firs%, GW_HWNDNEXT)
If UCase(APIGetText(firs%)) Like UCase(childhand) Then GoTo Greed
Wend
GoTo GoBac
FindWindowLike = 0

Greed:
room% = firs%
UntilWindowTitle = room%
End Function

Sub waitforok()
'Jaguar:
'Waits for the AOL OK Messages that pops up
Do
DoEvents
okw% = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw% <> 0
   
    okb% = FindChildByTitle(okw%, "OK")
    okd = SendMessageByNum(okb%, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub WinClickStartMenu()
'Shadow:
'This clicks the start menu button. It's not very useful,
'but it's a good example of three api calls.
Taskbar% = FindWindow("Shell_TrayWnd", VBNUllstrring)
StartBtn% = FindChildByClass(Taskbar%, "Button")
SendMessage StartBtn%, WM_KEYDOWN, VK_SPACE, 0
End Sub

Sub WinCreateShortcut(PutLinkInFolder As String, LinkName As String, LinkToFile As String)
'Shadow:
'this creates a shortcut to put on the start menu or
'desktop.  if it creates file not found error, mail me
'and i'll send you STKIT432.dll
CreateShellLink PutLinkInFolder, LinkName, LinkToFile, vbNullString
End Sub

Function WinGetDesktopDir() As String
'Shadow:
'finds the folder where shortcuts on the desktop are
'stored
WinGetDesktopDir = WinGetWinDir & "Desktop\"
End Function

Function WinGetStartDir() As String
'Shadow:
'finds the folder where shortcuts on the start menu are
'stored
WinGetStartDir = WinGetWinDir & "Start Menu\"
End Function

Function WinGetSystemDir() As String
'Shadow:
'finds the windows system directory
Buffer$ = String$(255, 0)
X = GetSystemDirectory(Buffer$, 255)
Buffer$ = TrimCharacter(Buffer$, Chr(0))
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
WinGetSystemDir = Buffer$
End Function

Function WinGetTempDir() As String
'Shadow:
'finds the window's temporary directory
Buffer$ = String$(255, 0)
X = GetTempPath(255, Buffer$)
Buffer$ = TrimCharacter(Buffer$, Chr(0))
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
WinGetTempDir = Buffer$
End Function

Function WinGetWinDir() As String
'Jaguar:
'finds the window's directory
Buffer$ = String$(255, 0)
X = GetWindowsDirectory(Buffer$, 255)
Buffer$ = TrimCharacter(Buffer$, Chr(0))
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
WinGetWinDir = Buffer$
End Function

Sub WinProtectINI()
'Shadow:
'Protects the win.ini file from being wrote to easily
inifile = GetWindowsDir & "win.ini"
SetAttr inifile, vbArchive + vbReadOnly + vbSystem
End Sub

Sub WinUnProtectINI()
'Shadow:
'Unprotects the win.ini file from being wrote to easily
inifile = GetWindowsDir & "win.ini"
SetAttr inifile, vbArchive
End Sub
