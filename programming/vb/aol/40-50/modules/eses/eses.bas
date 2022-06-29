Attribute VB_Name = "wh0re"
'eses.bas by eses eses00@yahoo.com
Option Explicit
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Public Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal DWreserved As Long)
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd&, ByVal wMsg&, ByVal wParam&, ByVal lparam&)
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wflags As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Const CB_GETCOUNT& = &H146
Public Const CB_SETCURSEL& = &H14E
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const HWND_TOPMOST = -1
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_SETCURSEL = &H186
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const WM_CLOSE = &H10
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
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SW_MAXIMIZE& = 3
Public Const SW_MINIMIZE& = 6
Public Const SW_RESTORE& = 9
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const WM_CHAR = &H102
Public Const WM_COMMAND = &H111
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const ENTER_KEY = 13
Public Const Op_Flags = PROCESS_READ Or RIGHTS_REQUIRED
Public Function aol()
aol = FindWindow("AOL Frame25", vbNullString)
End Function
Public Function mdi()
mdi = FindWindowEx(aol, 0&, "MDIClient", vbNullString)
End Function
Public Sub KW(KeyWord$)
'this method uses the edit\combo box on the aol toolbar
Dim toolbar1&, toolbar2&, ComboBox&, editwin&
toolbar1& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
ComboBox& = FindWindowEx(toolbar2&, 0&, "_AOL_ComboBox", vbNullString)
editwin& = FindWindowEx(ComboBox&, 0&, "Edit", vbNullString)
Call SendMessageByString(editwin&, WM_SETTEXT, 0, KeyWord)
Call SendMessageLong(editwin&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(editwin&, WM_CHAR, VK_RETURN, 0&)
TimeOut (0.2)
Call SendMessageByString(editwin&, WM_SETTEXT, 0, "Type Keyword or Web Address here and click Go")
End Sub
Public Sub TimeOut(length&)
Dim time As Long
time = Timer
Do
DoEvents
Loop Until Timer - time >= length
End Sub
Public Sub Pr(room$)
Call KW2("aol://2719:2-2-" & room$)
Call WaitForOkOrChat(room$)
End Sub
Public Sub IM(sn$, wuttosay$)
Dim daim&, text&, recipient&, send&, errorwin&, count&, errorbut&
Call KW2("im")
Do
DoEvents
daim& = FindWindowEx(mdi, 0&, "AOL Child", "Send Instant Message")
text& = FindWindowEx(daim&, 0&, "RICHCNTL", vbNullString)
send& = FindWindowEx(daim&, 0&, "_AOL_Icon", vbNullString)
recipient& = FindWindowEx(daim&, 0&, "_AOL_Edit", vbNullString)
For count& = 0 To 7
send& = FindWindowEx(daim&, send&, "_AOL_Icon", vbNullString)
Next count&
Loop Until daim& <> 0& And send& <> 0& And text& <> 0&
Call SendMessageByString(recipient&, WM_SETTEXT, 0&, sn$)
Call SendMessageByString(text&, WM_SETTEXT, 0&, wuttosay$)
Call SendMessage(send&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(send&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
errorwin& = FindWindow("#32770", "America Online")
daim& = FindWindowEx(mdi, 0&, "AOL Child", "Send Instant Message")
Loop Until errorwin& <> 0 Or daim& = 0
If errorwin <> 0 Then
errorbut& = FindWindowEx(errorwin&, 0&, "Button", vbNullString)
Call PostMessage(errorbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(errorbut&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(daim&, WM_CLOSE, 0&, 0&)
End If
End Sub
Public Sub IM2(sn$, wuttosay$)
Dim daim&, text&, recipient&, send&, errorwin&, count&, errorbut&
Call KW("im")
Do
DoEvents
daim& = FindWindowEx(mdi, 0&, "AOL Child", "Send Instant Message")
text& = FindWindowEx(daim&, 0&, "RICHCNTL", vbNullString)
send& = FindWindowEx(daim&, 0&, "_AOL_Icon", vbNullString)
recipient& = FindWindowEx(daim&, 0&, "_AOL_Edit", vbNullString)
For count& = 0 To 7
send& = FindWindowEx(daim&, send&, "_AOL_Icon", vbNullString)
Next count&
Loop Until daim& <> 0& And send& <> 0& And text& <> 0&
Call SendMessageByString(recipient&, WM_SETTEXT, 0&, sn$)
Call SendMessageByString(text&, WM_SETTEXT, 0&, wuttosay$)
Call SendMessage(send&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(send&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub RunToolbar(IconNumber&, letter$)
Dim menu&, toolbar1&, toolbar2&, icon&, count&, found&
toolbar1& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1, 0&, "_AOL_Toolbar", vbNullString)
icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
For count& = 1 To IconNumber&
icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
Next count&
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
menu& = FindWindow("#32768", vbNullString)
found& = IsWindowVisible(menu&)
Loop Until found& <> 0
letter$ = Asc(letter$)
Call PostMessage(menu&, WM_CHAR, letter$, 0&)

End Sub

Public Function FindChat()
Dim window&, Chat&, list&, text&
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
Chat& = FindWindowEx(window&, 0&, "RICHCNTL", vbNullString)
text& = FindWindowEx(window&, Chat&, "RICHCNTL", vbNullString)
list& = FindWindowEx(window&, 0&, "_AOL_Listbox", vbNullString)
If Chat& <> 0 And text& <> 0 And list& <> 0 Then
FindChat = window&
Exit Function
End If
window& = FindWindowEx(mdi, window&, "AOL Child", vbNullString)
Loop Until window& = 0
FindChat = 0
End Function
Public Sub SendText(wuttosay$)
'send text to the room
Dim text&, count&, room&
room& = FindChat
If room& = 0 Then Exit Sub
text& = FindWindowEx(room&, 0&, "RICHCNTL", vbNullString)
text& = FindWindowEx(room&, text&, "RICHCNTL", vbNullString)
Call SendMessageByString(text&, WM_SETTEXT, 0&, wuttosay$)
Call SendMessageLong(text&, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Function GetText(window&)
Dim text$, length&
length& = SendMessage(window&, WM_GETTEXTLENGTH, 0&, 0&)
text$ = String(length&, 0&)
Call SendMessageByString(window&, WM_GETTEXT, length& + 1, text$)
GetText = text$
End Function
Public Function FirstLetterBold(thestring$)
Dim count&, first$, Second$, sentence$, Third$
first$ = Mid(thestring$, 1&, 1)
sentence$ = "<b>" & first$ & "</b>"
For count& = 2 To Len(thestring$)
first$ = Mid(thestring$, count&, 1)
Second$ = Mid(thestring$, count& - 1, 1)
Third$ = Mid(thestring$, count& + 1, 1)
If first$ = " " And Second$ = " " And Third$ = " " Then
sentence$ = sentence$ & first$
GoTo bottom:
End If
If first$ = " " Then
sentence$ = sentence$ & first$ & "<b>"
GoTo bottom:
End If
If Second$ = " " And first$ <> " " Then
sentence$ = sentence$ & first$ & "</b>"
GoTo bottom:
End If
sentence$ = sentence$ & first$
bottom:
Next count&
FirstLetterBold = sentence$
End Function
Public Sub CollectMemDir(searchfor$, thelist As ListBox, onlyonline As Boolean)
Dim dirwin&, resultswin&, more&, list&, nowin&, nobutton&
If onlyonline = True Then
Call KW2("aol://4950:0000010000|all:" & searchfor$ & "|online:")
Else
Call KW2("aol://4950:0000010000|all:" & searchfor$)
End If
Do
DoEvents
dirwin& = FindWindowEx(mdi, 0&, "AOL Child", "Member Directory")
resultswin& = FindWindowEx(mdi, 0&, "AOL Child", "Member Directory Search Results")
list& = FindWindowEx(resultswin&, 0&, "_AOL_Listbox", vbNullString)
more& = FindWindowEx(resultswin&, 0&, "_AOL_Icon", vbNullString)
nowin& = FindWindow("#32770", "America Online")
Loop Until (nowin& <> 0) Or (resultswin& <> 0 And list& <> 0 And more& <> 0 And dirwin& <> 0)
If nowin& <> 0 Then
nobutton& = FindWindowEx(nowin&, 0&, "Button", vbNullString)
Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
Loop Until dirwin& <> 0
Call SendMessage(dirwin&, WM_CLOSE, 0&, 0&)
Exit Sub
End If

Call WaitForListToLoad(list&)
If SendMessage(list&, LB_GETCOUNT, 0&, 0&) < 20 Then GoTo add:
Call PostMessage(more&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(more&, WM_LBUTTONUP, 0, 0&)
Call WaitForListToLoad(list&)
If nowin& <> 0 Then
nobutton& = FindWindowEx(nowin&, 0&, "Button", vbNullString)
Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
GoTo add:
End If
If SendMessage(list&, LB_GETCOUNT, 0&, 0&) < 40 Then GoTo add:
Call PostMessage(more&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(more&, WM_LBUTTONUP, 0, 0&)
Call WaitForListToLoad(list&)
If nowin& <> 0 Then
nobutton& = FindWindowEx(nowin&, 0&, "Button", vbNullString)
Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
GoTo add:
End If
If SendMessage(list&, LB_GETCOUNT, 0&, 0&) < 60 Then GoTo add:
Call PostMessage(more&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(more&, WM_LBUTTONUP, 0, 0&)
Call WaitForListToLoad(list&)
If nowin& <> 0 Then
nobutton& = FindWindowEx(nowin&, 0&, "Button", vbNullString)
Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
GoTo add:
End If
If SendMessage(list&, LB_GETCOUNT, 0&, 0&) < 80 Then GoTo add:
Call PostMessage(more&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(more&, WM_LBUTTONUP, 0, 0&)
Call WaitForListToLoad(list&)
If nowin& <> 0 Then
nobutton& = FindWindowEx(nowin&, 0&, "Button", vbNullString)
Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
GoTo add:
End If
add:
Call AddNames(list&, thelist)
Call SendMessage(resultswin&, WM_CLOSE, 0&, 0&)
Call SendMessage(dirwin&, WM_CLOSE, 0&, 0&)

End Sub

Public Sub WaitForListToLoad(thelist&)
Dim one&, two&, three&
Do
one& = SendMessage(thelist&, LB_GETCOUNT, 0&, 0&)
TimeOut 1
two& = SendMessage(thelist&, LB_GETCOUNT, 0&, 0&)
TimeOut 1
three& = SendMessage(thelist&, LB_GETCOUNT, 0&, 0&)
Loop Until one& = two& And two& = three&
End Sub
Public Function ReplaceString(thestring$, find$, ReplaceWith$)
Dim Number$, all$, leftside$, rightside$
If thestring$ = "" Then Exit Function
If InStr(thestring$, find$) = 0 Then Exit Function
all$ = thestring$
Do
DoEvents
Number$ = InStr(all$, find$)
leftside$ = Left(all$, Number$ - 1)
rightside$ = Right(all$, Len(all$) - Number$ - Len(find$) + 1)
all$ = leftside$ & ReplaceWith$ & rightside$
Loop Until InStr(all$, find$) = 0
ReplaceString = all$
End Function
Public Sub AddNames(aollist&, vblist As ListBox)
'/because i dunno wtf procces and thread is
'/ i stole this part for dos's addroomtolistbox sub
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, index As Long, room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim Ta As Long, Ta2 As Long
rList& = aollist&
If rList& = 0& Then Exit Sub
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
screenname$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal screenname$, 4)
psnHold& = psnHold& + 6
screenname$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
Ta& = InStr(1, screenname$, Chr(9))
Ta2& = InStr(Ta& + 1, screenname$, Chr(9))
screenname$ = Mid(screenname$, Ta& + 1, Ta2& - 2)
screenname$ = Right(screenname$, Len(screenname$) - InStr(screenname$, Chr(9)))
vblist.AddItem (LCase(RemoveSpaces(screenname$)))
Next index&
Call CloseHandle(mThread)
End If
'//----
End Sub
Public Sub AddRoom(thelist As ListBox)
Dim room&, list&
room& = FindChat
If room& = 0 Then Exit Sub
list& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
Call AddNames(list&, thelist)
End Sub
Public Function RemoveSpaces(thestring$)
Dim count&, letter$, sentence$
If thestring$ = "" Then Exit Function
For count& = 1 To Len(thestring$)
letter$ = Mid(thestring$, count&, 1)
If letter$ = " " Then
GoTo bottom
End If
sentence$ = sentence$ & letter$
bottom:
Next count&
RemoveSpaces = sentence$
End Function
Public Sub AddMemDir(thelist As ListBox)
Dim dirwin&, resultswin&, list&
dirwin& = FindWindowEx(mdi, 0&, "AOL Child", "Member Directory")
resultswin& = FindWindowEx(mdi, 0&, "AOL Child", "Member Directory Search Results")
list& = FindWindowEx(resultswin&, 0&, "_AOL_Listbox", vbNullString)
If dirwin& = 0 Then Exit Sub
Call AddNames(list&, thelist)
End Sub
Public Sub AddWhosChatting(thelist As ListBox)
Dim findachat&, chatting&, list&, processthread&, name$, searchindex&, bytesread&, process&, listhandle3&, listholditem&
findachat& = FindWindowEx(mdi, 0&, "AOL Child", "Find a Chat")
chatting& = FindWindowEx(mdi, 0&, "AOL Child", "Who's Chatting")
list& = FindWindowEx(chatting&, 0&, "_AOL_Listbox", vbNullString)
If chatting& = 0 Then Exit Sub
listhandle3& = list&
'taken from cronx3.bas
Call GetWindowThreadProcessId(listhandle3&, process&)
processthread& = OpenProcess(Op_Flags, False, process&)
If processthread& Then
For searchindex& = 0 To SendMessageLong(listhandle3&, LB_GETCOUNT, 0&, 0&) - 1
name$ = String(4, vbNullChar)
listholditem& = SendMessage(listhandle3&, LB_GETITEMDATA, ByVal CLng(searchindex&), 0&)
listholditem& = listholditem& + 24
Call ReadProcessMemory(processthread&, listholditem&, name$, 4, bytesread&)
Call RtlMoveMemory(listholditem&, ByVal name$, 4)
listholditem& = listholditem& + 6
name$ = String(25, vbNullChar)
Call ReadProcessMemory(processthread&, listholditem&, name$, Len(name$), bytesread&)
name$ = Mid(name$, 3, InStr(name$, vbNullChar))
If name$ <> User Then
thelist.AddItem LCase(RemoveSpaces(name$))
End If
Next searchindex&
Call CloseHandle(processthread)
End If
'-/-
End Sub
Public Sub AddBLtoList(thelist As ListBox)
Dim buddylist&, edit&, editicon&, caption2$, caption$, window&, firstlist&, secondlist&, count&
Call KW2("buddylist")
TimeOut 1.5
buddylist& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
caption2$ = InStr(GetCaption(buddylist&), "Buddy Lists")
If caption2$ <> 0 Then GoTo start:
Do
DoEvents
buddylist& = FindWindowEx(mdi, buddylist&, "AOL Child", vbNullString)
caption2$ = InStr(GetCaption(buddylist&), "Buddy Lists")
Loop Until buddylist& <> 0 & caption2$ <> 0
start:
firstlist& = FindWindowEx(buddylist&, 0&, "_AOL_Listbox", vbNullString)
Call WaitForListToLoad(firstlist&)
For count& = 0 To SendMessage(firstlist&, LB_GETCOUNT, 0, 0) - 1
Call SendMessage(firstlist&, LB_SETCURSEL, count&, 0&)
editicon& = FindWindowEx(buddylist&, 0&, "_AOL_Icon", vbNullString)
editicon& = FindWindowEx(buddylist&, editicon&, "_AOL_Icon", vbNullString)
Call SendMessage(editicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(editicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
caption$ = GetCaption(window&)
If InStr(caption$, "Edit List") <> 0 Then
edit& = window&
End If
Loop Until edit& = window&
secondlist& = FindWindowEx(edit&, 0&, "_AOL_Listbox", vbNullString)
Call WaitForListToLoad(secondlist&)
Call AddNames(secondlist&, thelist)
TimeOut 0.2
Call SendMessage(edit&, WM_CLOSE, 0&, 0&)
TimeOut 0.2
Next count&
Call SendMessage(edit&, WM_CLOSE, 0&, 0&)
Call SendMessage(buddylist&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub AddNameToBL(name$)
Dim buddylist&, edit&, editicon&, addicon&, okwin&, okbut&, saveicon&, caption2$, namebox&, caption$, window&, firstlist&, secondlist&, count&
Call KW2("buddylist")
TimeOut 1.5
Do
DoEvents
buddylist& = FindBuddyListEdit
Loop Until buddylist& <> 0
editicon& = FindWindowEx(buddylist&, 0&, "_AOL_Icon", vbNullString)
editicon& = FindWindowEx(buddylist&, editicon&, "_AOL_Icon", vbNullString)
Call SendMessage(editicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(editicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
caption$ = GetCaption(window&)
If InStr(caption$, "Edit List") <> 0 Then
edit& = window&
End If
Loop Until edit& = window&
namebox& = FindWindowEx(edit&, 0&, "_AOL_Edit", vbNullString)
namebox& = FindWindowEx(edit&, namebox&, "_AOL_Edit", vbNullString)
Call SendMessageByString(namebox&, WM_SETTEXT, 0&, name$)
addicon& = FindWindowEx(edit&, 0&, "_AOL_Icon", vbNullString)
Call SendMessage(addicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(addicon&, WM_LBUTTONUP, 0, 0&)
saveicon& = FindWindowEx(edit&, editicon&, "_AOL_Icon", vbNullString)
saveicon& = FindWindowEx(edit&, saveicon&, "_AOL_Icon", vbNullString)
saveicon& = FindWindowEx(edit&, saveicon&, "_AOL_Icon", vbNullString)
saveicon& = FindWindowEx(edit&, saveicon&, "_AOL_Icon", vbNullString)
TimeOut 0.8
Call SendMessage(saveicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(saveicon&, WM_LBUTTONUP, 0, 0&)
Call WindowClose(buddylist&)
Do
DoEvents
okwin& = FindWindow("#32770", "America Online")
okbut& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until (okwin& <> 0 And okbut& <> 0)
Call SendMessage(okbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(okbut&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub AddListToBL(thelist As ListBox)
Dim buddylist&, edit&, editicon&, addicon&, okwin&, okbut&, saveicon&, caption2$, namebox&, caption$, window&, firstlist&, secondlist&, count&
Call KW2("buddylist")
TimeOut 1.5
Do
DoEvents
buddylist& = FindBuddyListEdit
Loop Until buddylist& <> 0
editicon& = FindWindowEx(buddylist&, 0&, "_AOL_Icon", vbNullString)
editicon& = FindWindowEx(buddylist&, editicon&, "_AOL_Icon", vbNullString)
Call SendMessage(editicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(editicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
caption$ = GetCaption(window&)
If InStr(caption$, "Edit List") <> 0 Then
edit& = window&
End If
Loop Until edit& = window&
For count& = 0 To thelist.ListCount - 1
namebox& = FindWindowEx(edit&, 0&, "_AOL_Edit", vbNullString)
namebox& = FindWindowEx(edit&, namebox&, "_AOL_Edit", vbNullString)
Call SendMessageByString(namebox&, WM_SETTEXT, 0&, thelist.list(count&))
addicon& = FindWindowEx(edit&, 0&, "_AOL_Icon", vbNullString)
Call SendMessage(addicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(addicon&, WM_LBUTTONUP, 0, 0&)
saveicon& = FindWindowEx(edit&, editicon&, "_AOL_Icon", vbNullString)
saveicon& = FindWindowEx(edit&, saveicon&, "_AOL_Icon", vbNullString)
saveicon& = FindWindowEx(edit&, saveicon&, "_AOL_Icon", vbNullString)
saveicon& = FindWindowEx(edit&, saveicon&, "_AOL_Icon", vbNullString)
TimeOut 0.8
Next count&

Call SendMessage(saveicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(saveicon&, WM_LBUTTONUP, 0, 0&)
Call WindowClose(buddylist&)
Do
DoEvents
okwin& = FindWindow("#32770", "America Online")
okbut& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until (okwin& <> 0 And okbut& <> 0)
Call SendMessage(okbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(okbut&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Function FindBuddyListEdit()
Dim buddylist&, caption$
buddylist& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
caption$ = InStr(GetCaption(buddylist&), "Buddy Lists")
If caption$ <> 0 Then
FindBuddyListEdit = buddylist&
Exit Function
End If
buddylist& = FindWindowEx(mdi, buddylist&, "AOL Child", vbNullString)
Loop Until buddylist& = 0
FindBuddyListEdit = 0
End Function
Public Function FindBuddyList()
FindBuddyList = FindWindowEx(mdi, 0&, "AOL Child", "Buddy List Window")
End Function
Public Sub OpenNewMail()
Call RunToolbar(2, "R")
End Sub
Public Sub OpenOldMail()
Call RunToolbar(2, "O")
End Sub
Public Sub OpenSentMail()
Call RunToolbar(2, "S")
End Sub

Public Function ProfileOpen(sn$) As String
Dim getwin&, prowin&, edit&, geticon&, nopro&, nobut&
Call RunToolbar(9, "G")
Do
DoEvents
getwin& = FindWindowEx(mdi, 0&, "AOL Child", "Get a Member's Profile")
edit& = FindWindowEx(getwin&, 0&, "_AOL_Edit", vbNullString)
geticon& = FindWindowEx(getwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until getwin& <> 0 And edit& <> 0 And geticon& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, sn$)
Call SendMessage(geticon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(geticon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
prowin& = FindWindowEx(mdi, 0&, "AOL Child", "Member Profile")
nopro& = FindWindow("#32770", "America Online")
Loop Until nopro& <> 0 Or prowin <> 0
If nopro <> 0 Then
nobut& = FindWindowEx(nopro&, 0&, "Button", vbNullString)
Call PostMessage(nobut&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(nobut&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(getwin&, WM_CLOSE, 0&, 0&)
ProfileOpen$ = 0
End If
ProfileOpen$ = prowin&
End Function
Public Function ProfileScroll(sn$) As String
Dim getwin&, prowin&, edit&, geticon&, profile$, nopro&, nobut&, text&
Call RunToolbar(9, "G")
Do
DoEvents
getwin& = FindWindowEx(mdi, 0&, "AOL Child", "Get a Member's Profile")
edit& = FindWindowEx(getwin&, 0&, "_AOL_Edit", vbNullString)
geticon& = FindWindowEx(getwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until getwin& <> 0 And edit& <> 0 And geticon& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, sn$)
Call SendMessage(geticon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(geticon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
prowin& = FindWindowEx(mdi, 0&, "AOL Child", "Member Profile")
nopro& = FindWindow("#32770", "America Online")
Loop Until nopro& <> 0 Or prowin <> 0
If nopro <> 0 Then
nobut& = FindWindowEx(nopro&, 0&, "Button", vbNullString)
Call PostMessage(nobut&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(nobut&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(getwin&, WM_CLOSE, 0&, 0&)
ProfileScroll$ = 0
Exit Function
End If
Do
DoEvents
text& = FindWindowEx(prowin&, 0&, "RICHCNTL", vbNullString)
Loop Until text& <> 0
TimeOut 1
profile$ = GetText(text&)
TimeOut 0.5
Call PostMessage(getwin&, WM_CLOSE, 0&, 0&)
Call PostMessage(prowin&, WM_CLOSE, 0&, 0&)
ProfileScroll$ = profile$
SendMacro (profile$)
End Function
Public Sub SendMacro(themacro$, Optional pause$ = 0.8)
Dim count&, letter$, sentence$
For count& = 1 To Len(themacro$)
letter$ = Mid(themacro$, count&, 1)
sentence$ = sentence$ & letter$
If letter$ = Chr(13) Then
SendText (ReplaceString(sentence$, Chr(13), ""))
sentence$ = ""
TimeOut (pause$)
End If
Next count&
End Sub
Public Function GetCaption(window&)
Dim caption$, length&
length& = GetWindowTextLength(window&)
caption$ = String(length&, 0)
Call GetWindowText(window&, caption$, length& + 1)
GetCaption = caption$
End Function
Public Sub SetCaption(window&, newcaption$)
Call SetWindowText(window&, newcaption$)
End Sub
Public Function CheckIfMaster()
Dim firstwin&, secondwin&, count&, image&, icon&, icon2&, subwin&, subbut&
Call RunToolbar(5, "C")
Do
DoEvents
firstwin& = FindWindowEx(mdi, 0&, "AOL Child", " AOL Parental Controls")
icon& = FindWindowEx(firstwin&, 0&, "_AOL_Icon", vbNullString)
image& = FindWindowEx(firstwin&, 0&, "_AOL_Image", vbNullString)
Loop Until firstwin& <> 0 And icon& <> 0 And image& <> 0
TimeOut 1.7
icon& = FindWindowEx(firstwin&, icon&, "_AOL_Icon", vbNullString)
Call SendMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon&, WM_LBUTTONUP, 0, 0&)
Do
secondwin& = FindWindowEx(mdi, 0&, "AOL Child", "Parental Controls")
subwin& = FindWindow("#32770", "America Online")
subbut& = FindWindowEx(subwin&, 0&, "Button", vbNullString)
Loop Until secondwin <> 0 Or (subwin& <> 0 And subbut& <> 0)
TimeOut 0.7
If subwin& <> 0 Then
Call PostMessage(subbut&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(subbut&, WM_LBUTTONUP, 0&, 0&)
CheckIfMaster = False
Exit Function
Else
Call SendMessage(firstwin&, WM_CLOSE, 0&, 0&)
Call SendMessage(secondwin&, WM_CLOSE, 0&, 0&)
CheckIfMaster = True
End If
End Function
Public Sub SendMail(sn$, subject$, body$)
Dim toolbar1&, toolbar2&, icon&, Rich&, savewin&, savebut&, count&, mail&, send&, edit&, modal&, ModalIcon&, unknown&
toolbar1& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1, 0&, "_AOL_Toolbar", vbNullString)
icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
Call SendMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
mail& = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
edit& = FindWindowEx(mail&, 0&, "_AOL_Edit", vbNullString)
send& = FindWindowEx(mail&, 0&, "_AOL_icon", vbNullString)
Loop Until mail& <> 0 And send& <> 0 And edit& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, sn$)
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, subject$)
Rich& = FindWindowEx(mail&, Rich&, "RICHCNTL", vbNullString)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, body$)
For count& = 1 To 13
send& = FindWindowEx(mail&, send&, "_AOL_icon", vbNullString)
Next count&
Call SendMessage(send&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(send&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
modal& = FindWindow("_AOL_Modal", vbNullString)
unknown& = FindWindowEx(mdi, 0&, "AOL Child", "Error")
Loop Until modal& <> 0 Or unknown& <> 0
If unknown <> 0 Then
Call PostMessage(unknown&, WM_CLOSE, 0&, 0&)
DoEvents
Call PostMessage(mail&, WM_CLOSE, 0&, 0&)
DoEvents
Do
DoEvents
savewin& = FindWindow("#32770", "America Online")
savebut& = FindWindowEx(savewin&, 0&, "Button", "&No")
Loop Until savewin& <> 0 And savebut& <> 0

Call SendMessage(savebut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(savebut&, WM_KEYUP, VK_SPACE, 0&)
Exit Sub
End If
modal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Function CheckIfDead(sn$)
Dim toolbar1&, toolbar2&, icon&, Rich&, errorview&, ErrorText$, savewin&, savebut&, count&, mail&, send&, edit&, unknown&
toolbar1& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1, 0&, "_AOL_Toolbar", vbNullString)
icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
Call SendMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
mail& = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
edit& = FindWindowEx(mail&, 0&, "_AOL_Edit", vbNullString)
send& = FindWindowEx(mail&, 0&, "_AOL_icon", vbNullString)
Loop Until mail& <> 0 And send& <> 0 And edit& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, sn$ & ",*")
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, "boo")
Rich& = FindWindowEx(mail&, Rich&, "RICHCNTL", vbNullString)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, "-=P")
For count& = 1 To 13
send& = FindWindowEx(mail&, send&, "_AOL_icon", vbNullString)
Next count&
Call SendMessage(send&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(send&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
unknown& = FindWindowEx(mdi, 0&, "AOL Child", "Error")
errorview& = FindWindowEx(unknown, 0&, "_AOL_View", vbNullString)
Loop Until unknown& <> 0 And errorview& <> 0
ErrorText$ = GetText(errorview&)
Call PostMessage(unknown&, WM_CLOSE, 0&, 0&)
DoEvents
Call PostMessage(mail&, WM_CLOSE, 0&, 0&)
DoEvents
Do
DoEvents
savewin& = FindWindow("#32770", "America Online")
savebut& = FindWindowEx(savewin&, 0&, "Button", "&No")
Loop Until savewin& <> 0 And savebut& <> 0
Call SendMessage(savebut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(savebut&, WM_KEYUP, VK_SPACE, 0&)
ErrorText$ = LCase(RemoveSpaces(ErrorText$))
If InStr(ErrorText$, LCase(RemoveSpaces(sn$))) <> 0 Then
CheckIfDead = True
Else
CheckIfDead = False
End If
End Function
Public Sub RunMenu(topmenu&, submenu&)
Dim menu&, menu2&, themenu&
menu& = GetMenu(aol)
menu2& = GetSubMenu(menu&, topmenu&)
themenu& = GetMenuItemID(menu2&, submenu&)
Call SendMessageLong(aol, WM_COMMAND, themenu&, 0&)
End Sub
Public Sub WaitForOkOrChat(ChatRoom$)
Dim fullwin&, fullbut&, room$
Do
DoEvents
fullwin& = FindWindow("#32770", "America Online")
fullbut& = FindWindowEx(fullwin&, 0&, "Button", vbNullString)
If LCase(RemoveSpaces(ChatRoom$)) = LCase(RemoveSpaces(GetRoomName)) Then Exit Sub
If fullwin& <> 0 And fullbut& <> 0 Then Exit Sub: Call ClickButton(fullbut&)
Loop
End Sub

Public Sub FormOnTop(frm As Form)
Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, flags)
End Sub
Public Sub FormMove(frm As Form)
    Call ReleaseCapture
    Call SendMessage(frm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub IMsOff()
Call IM("$im_off", "{s goodbye")
End Sub
Public Sub IMsOn()
Call IM("$im_on", "{s welcome")
End Sub
Public Function GetFromINI(section As String, key As String, directory As String) As String
'//-//taken from dos32
Dim strBuffer As String
strBuffer = String(750, Chr(0))
key$ = LCase$(key$)
GetFromINI$ = Left(strBuffer, GetPrivateProfileString(section$, ByVal key$, "", strBuffer, Len(strBuffer), directory$))
End Function
Public Sub WriteToINI(section As String, key As String, keyvalue As String, directory As String)
'//-//taken from dos32
Call WritePrivateProfileString(section$, UCase$(key$), keyvalue$, directory$)
End Sub
Public Sub ChatXbyString(thestring$)
Dim daroom&, lilbox&, check&, glyph&, window&, index&, snlist&, dastatic&, icon&, icon2&
daroom& = FindChat
If daroom& = 0 Then Exit Sub
snlist& = FindWindowEx(daroom&, 0&, "_AOL_Listbox", vbNullString)
'//from dos32 addroomtolistbox sub
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim Ta As Long, Ta2 As Long
rList& = snlist&
If rList& = 0& Then Exit Sub
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
screenname$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal screenname$, 4)
psnHold& = psnHold& + 6
screenname$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
Ta& = InStr(1, screenname$, Chr(9))
Ta2& = InStr(Ta& + 1, screenname$, Chr(9))
screenname$ = Mid(screenname$, Ta& + 1, Ta2& - 2)
screenname$ = RemoveSpaces(screenname$)
'/- end of dos's code
If InStr(LCase(RemoveSpaces(screenname$)), LCase(RemoveSpaces(thestring$))) <> 0 And screenname$ <> User Then
snlist& = FindWindowEx(daroom&, 0&, "_AOL_Listbox", vbNullString)

SendText ("°· storage · ignored: " & LCase(RemoveSpaces(screenname$)))

Call SendMessage(snlist&, LB_SETCURSEL, index&, 0&)
Call PostMessage(snlist&, WM_LBUTTONDBLCLK, 0&, 0&)
findlilbox:
DoEvents
lilbox& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
check& = FindWindowEx(lilbox&, 0&, "_AOL_Checkbox", vbNullString)
glyph& = FindWindowEx(lilbox&, 0&, "_AOL_Glyph", vbNullString)
dastatic& = FindWindowEx(lilbox&, 0&, "_AOL_Static", vbNullString)
icon& = FindWindowEx(lilbox&, 0&, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(lilbox&, icon&, "_AOL_Icon", vbNullString)
If check& <> 0 And glyph& <> 0 And dastatic& <> 0 And icon& <> 0 And icon2& <> 0 Then
window& = lilbox&
GoTo ignore:
End If
lilbox& = FindWindowEx(mdi, lilbox&, "AOL Child", vbNullString)
Loop Until lilbox = 0
GoTo findlilbox:
ignore:
check& = FindWindowEx(window&, 0&, "_AOL_Checkbox", vbNullString)
Do Until SendMessage(check&, BM_GETCHECK, 0&, 0&) <> 0
DoEvents
Call PostMessage(check&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(check&, WM_LBUTTONUP, 0&, 0&)
Loop
Call PostMessage(window&, WM_CLOSE, 0&, 0&)
TimeOut 1
End If
Next index&
Call CloseHandle(mThread)
End If
End Sub
Public Function FindWelcome()
Dim window&, caption$
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
caption$ = GetCaption(window&)
If InStr(caption$, "Welcome,") <> 0 Then
FindWelcome = window&
Exit Function
End If
window& = FindWindowEx(mdi, window&, "AOL Child", vbNullString)
Loop Until window& = 0
FindWelcome = 0
End Function
Public Function User()
Dim window&, caption$, ex&
window& = FindWelcome
If window& = 0 Then Exit Function
caption$ = GetCaption(window&)
ex& = InStr(caption$, "!")
User = ReplaceString(Left(caption$, ex& - 1), "Welcome, ", "")
End Function
Public Sub ChatUnXbyString(thestring$)
Dim daroom&, lilbox&, check&, glyph&, window&, index&, snlist&, dastatic&, icon&, icon2&
daroom& = FindChat
If daroom& = 0 Then Exit Sub
snlist& = FindWindowEx(daroom&, 0&, "_AOL_Listbox", vbNullString)
'//from dos32 addroomtolistbox sub
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, room As Long
Dim rList As Long, sThread As Long, mThread As Long
Dim Ta As Long, Ta2 As Long
rList& = snlist&
If rList& = 0& Then Exit Sub
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
For index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
screenname$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal screenname$, 4)
psnHold& = psnHold& + 6
screenname$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
Ta& = InStr(1, screenname$, Chr(9))
Ta2& = InStr(Ta& + 1, screenname$, Chr(9))
screenname$ = Mid(screenname$, Ta& + 1, Ta2& - 2)
screenname$ = RemoveSpaces(screenname$)
'/- end of dos's code
If InStr(LCase(RemoveSpaces(screenname$)), LCase(RemoveSpaces(thestring$))) <> 0 And screenname$ <> User Then
SendText ("°· storage · un-ignored: " & LCase(RemoveSpaces(screenname$)))


snlist& = FindWindowEx(daroom&, 0&, "_AOL_Listbox", vbNullString)
Call SendMessage(snlist&, LB_SETCURSEL, index&, 0&)
Call PostMessage(snlist&, WM_LBUTTONDBLCLK, 0&, 0&)
lilbox& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
findlilbox:
DoEvents
Do
DoEvents
check& = FindWindowEx(lilbox&, 0&, "_AOL_Checkbox", vbNullString)
glyph& = FindWindowEx(lilbox&, 0&, "_AOL_Glyph", vbNullString)
dastatic& = FindWindowEx(lilbox&, 0&, "_AOL_Static", vbNullString)
icon& = FindWindowEx(lilbox&, 0&, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(lilbox&, icon&, "_AOL_Icon", vbNullString)
If check& <> 0 And glyph& <> 0 And dastatic& <> 0 And icon& <> 0 And icon2& <> 0 Then
window& = lilbox&
GoTo ignore:
End If
lilbox& = FindWindowEx(mdi, lilbox&, "AOL Child", vbNullString)
Loop Until lilbox = 0
GoTo findlilbox:
ignore:
check& = FindWindowEx(window&, 0&, "_AOL_Checkbox", vbNullString)
Do Until SendMessage(check&, BM_GETCHECK, 0&, 0&) = 0
DoEvents
Call PostMessage(check&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(check&, WM_LBUTTONUP, 0&, 0&)
Loop
Call PostMessage(window&, WM_CLOSE, 0&, 0&)
TimeOut 1
End If
Next index&
Call CloseHandle(mThread)
End If
End Sub
Public Function IsUserOnline()
If FindWelcome = 0 Then
IsUserOnline = False
Else
IsUserOnline = True
End If
End Function
Public Sub PlayWav(file$)
If Dir(file$) = "" Then Exit Sub
Call sndPlaySound(file$, SND_FLAG)
End Sub
Public Function FindSentIM()
'-/-this finds the IM u sent
Dim window&, caption$
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
caption$ = GetCaption(window&)
If InStr(caption$, "Instant Message") <> 0 And InStr(caption$, ">") = 0 Then
FindSentIM = window&
Exit Function
End If
window& = FindWindowEx(mdi, window&, "AOL Child", vbNullString)
Loop Until window& = 0
FindSentIM = 0
End Function
Public Function FindRecievedIM()
'-/-this finds the IM sumone sent you
'-/-good for baiters
Dim window&, caption$
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
caption$ = GetCaption(window&)
If InStr(caption$, "Instant Message") <> 0 And InStr(caption$, ">") <> 0 Then
FindRecievedIM = window&
Exit Function
End If
window& = FindWindowEx(mdi, window&, "AOL Child", vbNullString)
Loop Until window& = 0
FindRecievedIM = 0
End Function
Public Function GetIMSender()
'/-if you respond to the IM this sub will not work
Dim IM&, caption$
IM& = FindRecievedIM
If IM& = 0 Then Exit Function
caption$ = GetCaption(IM&)
GetIMSender = Mid(caption$, InStr(caption$, ":") + 1, Len(caption$))
'-/-if you still want the IM sender. even tho you responded
'-/-un-qoute this. then erase the above. DO NOT un-quote
'-/- this if you are making a baiter.
'Dim im&, Caption$
'im& = FindRecievedIM
'If im& = 0 Then
'im& = FindSentIM
'If im& = 0 Then Exit Function
'Caption$ = GetCaption(im&)
'GetIMSender = Mid(Caption$, InStr(Caption$, ":") + 1, Len(Caption$))
End Function
Public Sub SendChatLink(url$, wuttosay$, underlined As Boolean)
If underlined = True Then
SendText ("< a href=http://" & url$ & ">" & wuttosay$ & "</a>")
Else
SendText ("< a href=http://" & url$ & "></u>" & wuttosay$ & "</a>")
End If
End Sub
Public Sub CloseAllIMs()
Dim IM&
Do
DoEvents
IM& = FindSentIM
Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
Loop Until IM& = 0
Do
DoEvents
IM& = FindRecievedIM
Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
Loop Until IM& = 0
End Sub
Public Function GetTimeOnline()
Dim dastatic&, modal&, text$, icon&
Call RunToolbar(5, "O")
Do
DoEvents
modal& = FindWindow("_AOL_Modal", vbNullString)
dastatic& = FindWindowEx(modal&, 0&, "_AOL_Static", vbNullString)
icon& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
Loop Until modal& <> 0
text$ = GetText(dastatic&)
GetTimeOnline = ReplaceString(Mid(text$, InStr(text$, "for ") + 4, Len(text$)), ".", "")
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Sub SignOff()
If FindWelcome = 0 Then Exit Sub
Call RunMenu(3, 1)
End Sub
Public Function CheckIMs(sn$)
Dim daim&, text&, recipient&, availible&, send&, errorwin&, count&, errorbut&, dastatic&, datext$
Call RunToolbar(9, "I")
Do
DoEvents
daim& = FindWindowEx(mdi, 0&, "AOL Child", "Send Instant Message")
text& = FindWindowEx(daim&, 0&, "RICHCNTL", vbNullString)
recipient& = FindWindowEx(daim&, 0&, "_AOL_Edit", vbNullString)
send& = FindWindowEx(daim&, 0&, "_AOL_Icon", vbNullString)
availible& = FindWindowEx(daim&, send&, "_AOL_Icon", vbNullString)
Loop Until daim& <> 0& And text& <> 0& And send& <> 0 And availible& <> 0
Call SendMessageByString(recipient&, WM_SETTEXT, 0&, sn$)
Call SendMessageByString(text&, WM_SETTEXT, 0&, "can ya accept?")
TimeOut 0.5
send& = FindWindowEx(daim&, 0&, "_AOL_Icon", vbNullString)
availible& = FindWindowEx(daim&, send&, "_AOL_Icon", vbNullString)
For count& = 0 To 7
availible& = FindWindowEx(daim&, availible&, "_AOL_Icon", vbNullString)
Next count&
Call SendMessage(availible&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(availible&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
errorwin& = FindWindow("#32770", "America Online")
dastatic& = FindWindowEx(errorwin&, 0&, "Static", vbNullString)
dastatic& = FindWindowEx(errorwin&, dastatic&, "Static", vbNullString)
errorbut& = FindWindowEx(errorwin&, 0&, "Button", "OK")
Loop Until errorwin& <> 0 And dastatic& <> 0 And errorbut& <> 0
datext$ = GetText(dastatic&)
If InStr(datext$, "is online and able") <> 0 Then
CheckIMs = True
Else
CheckIMs = False
End If
Call PostMessage(errorbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call PostMessage(errorbut&, WM_KEYUP, VK_SPACE, 0&)
Call PostMessage(daim&, WM_CLOSE, 0&, 0&)
End Function
Public Function TrimTime(Optional AddAmOrPM As Boolean = True)
Dim datime$, min$, hour$, other$
datime$ = time
If InStr(datime$, "PM") <> 0 Then
other$ = "pm"
Else
other$ = "am"
End If
If AddAmOrPM = False Then
TrimTime = Left(datime$, Len(datime$) - 6)
Else
TrimTime = Left(datime$, Len(datime$) - 6) & " " & other$
End If
End Function
Public Function FullDate()
Dim damonth$, daday$, dayear$, daday2$, dafulldate$
If Month(Date) = 1 Then
damonth$ = "january"
ElseIf Month(Date) = 2 Then
damonth$ = "february"
ElseIf Month(Date) = 3 Then
damonth$ = "march"
ElseIf Month(Date) = 4 Then
damonth$ = "april"
ElseIf Month(Date) = 5 Then
damonth$ = "may"
ElseIf Month(Date) = 6 Then
damonth$ = "june"
ElseIf Month(Date) = 7 Then
damonth$ = "july"
ElseIf Month(Date) = 8 Then
damonth$ = "august"
ElseIf Month(Date) = 9 Then
damonth$ = "september"
ElseIf Month(Date) = 10 Then
damonth$ = "october"
ElseIf Month(Date) = 11 Then
damonth$ = "november"
ElseIf Month(Date) = 12 Then
damonth$ = "december"
End If
If WeekDay(Date) = 1 Then
daday$ = "sunday"
ElseIf WeekDay(Date) = 2 Then
daday$ = "monday"
ElseIf WeekDay(Date) = 3 Then
daday$ = "tuesday"
ElseIf WeekDay(Date) = 4 Then
daday$ = "wednesday"
ElseIf WeekDay(Date) = 5 Then
daday$ = "thursday"
ElseIf WeekDay(Date) = 6 Then
daday$ = "friday"
ElseIf WeekDay(Date) = 7 Then
daday$ = "saturday"
End If
If Day(Date) = 1 Or Day(Date) = 21 Or Day(Date) = 31 Then
daday2$ = Day(Date) & "st"
ElseIf Day(Date) = 2 Or Day(Date) = 22 Then
daday2$ = Day(Date) & "nd"
ElseIf Day(Date) = 3 Or Day(Date) = 23 Then
daday2$ = Day(Date) & "rd"
ElseIf Day(Date) = 4 Or Day(Date) = 24 Or Day(Date) = 5 Or Day(Date) = 6 Or Day(Date) = 7 Or Day(Date) = 8 Or Day(Date) = 9 Or Day(Date) = 10 Or Day(Date) = 11 Or Day(Date) = 12 Or Day(Date) = 13 Or Day(Date) = 14 Or Day(Date) = 15 Or Day(Date) = 16 Or Day(Date) = 17 Or Day(Date) = 18 Or Day(Date) = 19 Or Day(Date) = 20 Or Day(Date) = 25 Or Day(Date) = 26 Or Day(Date) = 27 Or Day(Date) = 28 Or Day(Date) = 29 Or Day(Date) = 30 Then
daday2$ = Day(Date) & "th"
End If
FullDate = daday$ & ", " & damonth$ & " " & daday2$ & ", " & Year(Date)
End Function
Public Sub SendMailWithAttach(sn$, subject$, body$, filepath$)
Dim toolbar1&, toolbar2&, atmodal&, atwin&, atbut&, atedit&, aticon2&, aticon&, count2&, icon&, Rich&, savewin&, savebut&, count&, mail&, send&, edit&, modal&, ModalIcon&, unknown&
toolbar1& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1, 0&, "_AOL_Toolbar", vbNullString)
icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
Call SendMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
mail& = FindWindowEx(mdi, 0&, "AOL Child", "Write Mail")
edit& = FindWindowEx(mail&, 0&, "_AOL_Edit", vbNullString)
send& = FindWindowEx(mail&, 0&, "_AOL_icon", vbNullString)
Loop Until mail& <> 0 And send& <> 0 And edit& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, sn$)
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
edit& = FindWindowEx(mail&, edit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(edit&, WM_SETTEXT, 0&, subject$)
Rich& = FindWindowEx(mail&, Rich&, "RICHCNTL", vbNullString)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, body$)
aticon& = FindWindowEx(mail&, 0&, "_AOL_icon", vbNullString)
For count2& = 1 To 12
aticon& = FindWindowEx(mail&, aticon&, "_AOL_icon", vbNullString)
Next count2&
Call SendMessage(aticon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(aticon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
atmodal& = FindWindow("_AOL_Modal", vbNullString)
aticon2& = FindWindowEx(atmodal&, 0&, "_AOL_icon", vbNullString)
Loop Until atmodal& <> 0 And aticon2& <> 0
Call SendMessage(aticon2&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(aticon2&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
atwin& = FindWindow("#32770", "Attach")
atbut& = FindWindowEx(atwin&, 0&, "Button", "&Open")
atedit& = FindWindowEx(atwin&, 0&, "Edit", vbNullString)
Loop Until atwin& <> 0 & atbut& <> 0 And atedit& <> 0
atwin& = FindWindow("#32770", "Attach")
atbut& = FindWindowEx(atwin&, 0&, "Button", "&Open")
atedit& = FindWindowEx(atwin&, 0&, "Edit", vbNullString)
TimeOut 0.8
Call SendMessageByString(atedit&, WM_SETTEXT, 0&, filepath$)
Call PostMessage(atbut&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(atbut&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
atmodal& = FindWindow("_AOL_Modal", vbNullString)
aticon2& = FindWindowEx(atmodal&, 0&, "_AOL_icon", vbNullString)
aticon2& = FindWindowEx(atmodal&, aticon2&, "_AOL_icon", vbNullString)
aticon2& = FindWindowEx(atmodal&, aticon2&, "_AOL_icon", vbNullString)
Loop Until atmodal& <> 0 And aticon2& <> 0
Call SendMessage(aticon2&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(aticon2&, WM_LBUTTONUP, 0, 0&)
For count& = 1 To 13
send& = FindWindowEx(mail&, send&, "_AOL_icon", vbNullString)
Next count&
Call SendMessage(send&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(send&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
modal& = FindWindow("_AOL_Modal", vbNullString)
unknown& = FindWindowEx(mdi, 0&, "AOL Child", "Error")
Loop Until modal& <> 0 Or unknown& <> 0
If unknown <> 0 Then
Call PostMessage(unknown&, WM_CLOSE, 0&, 0&)
Call PostMessage(mail&, WM_CLOSE, 0&, 0&)
Do
DoEvents
savewin& = FindWindow("#32770", "America Online")
savebut& = FindWindowEx(savewin&, 0&, "Button", "&No")
Loop Until savewin& <> 0 And savebut& <> 0

Call SendMessage(savebut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(savebut&, WM_KEYUP, VK_SPACE, 0&)
Exit Sub
End If
modal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Function FindSignon()
Dim window&, dastatic&, dastatic2&, dacombo&, dacombo2&, daicon&, daicon2&, daicon3&, daicon4&, errorwin&, errorbut&, welcome&
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
dastatic& = FindWindowEx(window&, 0&, "_AOL_Static", vbNullString)
dastatic2& = FindWindowEx(window&, dastatic&, "_AOL_Static", vbNullString)
dacombo& = FindWindowEx(window&, 0&, "_AOL_Combobox", vbNullString)
dacombo2& = FindWindowEx(window&, dacombo&, "_AOL_Combobox", vbNullString)
daicon& = FindWindowEx(window&, 0&, "_AOL_Icon", vbNullString)
daicon2& = FindWindowEx(window&, daicon&, "_AOL_Icon", vbNullString)
daicon3& = FindWindowEx(window&, daicon2&, "_AOL_Icon", vbNullString)
daicon4& = FindWindowEx(window&, daicon3&, "_AOL_Icon", vbNullString)
If window& <> 0 And dastatic& <> 0 And dastatic2& <> 0 And dacombo& <> 0 And dacombo2& <> 0 And daicon& <> 0 And daicon2& <> 0 And daicon3& <> 0 And daicon4& <> 0 Then
FindSignon = window&
Exit Function
End If
window& = FindWindowEx(mdi, window&, "AOL Child", vbNullString)
Loop Until window& = 0
FindSignon = 0
End Function
Public Sub SignonGuest(sn$, pw$)
Dim window&, icon&, dacombo&, count&, modal&, snedit&, pwedit&, snicon&, errorwin&, errorbut&, welcome&
window& = FindSignon
If window& = 0 Then Exit Sub
dacombo& = FindWindowEx(window&, 0&, "_AOL_Combobox", vbNullString)
Call PostMessage(dacombo&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(dacombo&, WM_LBUTTONUP, 0&, 0&)
Call SendMessageLong(dacombo&, CB_SETCURSEL, SendMessageLong(dacombo&, CB_GETCOUNT, 0&, 0&) - 1, 0&)
Call PostMessage(dacombo&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(dacombo&, WM_LBUTTONUP, 0&, 0&)
TimeOut 0.8
icon& = FindWindowEx(window&, 0&, "_AOL_Icon", vbNullString)
For count& = 1 To 3
icon& = FindWindowEx(window&, icon&, "_AOL_Icon", vbNullString)
Next count&
Do
DoEvents
Call SendMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon&, WM_LBUTTONUP, 0, 0&)
TimeOut 0.6
Loop Until IsWindowVisible(window&) = False
Do
DoEvents
modal& = FindWindow("_AOL_Modal", vbNullString)
snedit& = FindWindowEx(modal&, 0&, "_AOL_Edit", vbNullString)
pwedit& = FindWindowEx(modal&, snedit&, "_AOL_Edit", vbNullString)
snicon& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
Loop Until modal& <> 0 And snedit& <> 0 And pwedit& <> 0 And snicon& <> 0
modal& = FindWindow("_AOL_Modal", vbNullString)
snedit& = FindWindowEx(modal&, 0&, "_AOL_Edit", vbNullString)
pwedit& = FindWindowEx(modal&, snedit&, "_AOL_Edit", vbNullString)
snicon& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
Call SendMessageByString(snedit&, WM_SETTEXT, 0&, sn$)
Call SendMessageByString(pwedit&, WM_SETTEXT, 0&, pw$)
Call SendMessage(snicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(snicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
errorwin& = FindWindow("#32770", "America Online")
errorbut& = FindWindowEx(errorwin&, 0&, "Button", vbNullString)
welcome& = FindWelcome
Loop Until (errorwin& <> 0 And errorbut& <> 0) Or welcome& <> 0
If errorwin& <> 0 Then
Call SendMessage(errorbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(errorbut&, WM_KEYUP, VK_SPACE, 0&)
Call SendMessage(snicon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(snicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
errorwin& = FindWindow("#32770", "America Online")
errorbut& = FindWindowEx(errorwin&, 0&, "Button", vbNullString)
Loop Until errorwin& <> 0 And errorbut& <> 0
Call SendMessage(errorbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(errorbut&, WM_KEYUP, VK_SPACE, 0&)
End If
End Sub
Public Function GetIMMessage()
Dim IM&, text$, edit&, sentence$, letter$, count&
IM& = FindRecievedIM
If IM& = 0 Then
IM& = FindSentIM
End If
If IM& = 0 Then Exit Function
edit& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
text$ = GetText(edit&)
For count& = 1 To Len(text$)
letter$ = Mid(text$, count&, 1)
If letter$ = Chr(13) Then
sentence$ = ""
End If
sentence$ = sentence$ & letter$
Next count&
GetIMMessage = Mid(ReplaceString(sentence$, Chr(13), ""), InStr(ReplaceString(sentence$, Chr(13), ""), ":") + 3, Len(ReplaceString(sentence$, Chr(13), "")))
End Function
Public Function RandomNumber(HighestNumber$)
Randomize
RandomNumber = Int(Rnd * HighestNumber$) + 1
End Function
Public Function RandomLetter()
Dim Number$
Number$ = RandomNumber(26)
If Number$ = 1 Then
RandomLetter = "a"
ElseIf Number$ = 2 Then
RandomLetter = "b"
ElseIf Number$ = 3 Then
RandomLetter = "c"
ElseIf Number$ = 4 Then
RandomLetter = "d"
ElseIf Number$ = 5 Then
RandomLetter = "e"
ElseIf Number$ = 6 Then
RandomLetter = "f"
ElseIf Number$ = 7 Then
RandomLetter = "g"
ElseIf Number$ = 8 Then
RandomLetter = "h"
ElseIf Number$ = 9 Then
RandomLetter = "i"
ElseIf Number$ = 10 Then
RandomLetter = "j"
ElseIf Number$ = 11 Then
RandomLetter = "k"
ElseIf Number$ = 12 Then
RandomLetter = "l"
ElseIf Number$ = 13 Then
RandomLetter = "m"
ElseIf Number$ = 14 Then
RandomLetter = "n"
ElseIf Number$ = 15 Then
RandomLetter = "o"
ElseIf Number$ = 16 Then
RandomLetter = "p"
ElseIf Number$ = 17 Then
RandomLetter = "q"
ElseIf Number$ = 18 Then
RandomLetter = "r"
ElseIf Number$ = 19 Then
RandomLetter = "s"
ElseIf Number$ = 20 Then
RandomLetter = "t"
ElseIf Number$ = 21 Then
RandomLetter = "u"
ElseIf Number$ = 22 Then
RandomLetter = "v"
ElseIf Number$ = 23 Then
RandomLetter = "w"
ElseIf Number$ = 24 Then
RandomLetter = "x"
ElseIf Number$ = 25 Then
RandomLetter = "y"
ElseIf Number$ = 26 Then
RandomLetter = "z"
End If
End Function
Public Sub KillWait()
Dim window&, Button&
Call RunMenu(4, 10)
Do
DoEvents
window& = FindWindow("_AOL_Modal", vbNullString)
Button& = FindWindowEx(window&, 0&, "_AOL_Icon", vbNullString)
Loop Until window& <> 0 And Button& <> 0
Call SendMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Button&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub KillModal()
Dim window&, Button&
DoEvents
window& = FindWindow("_AOL_Modal", vbNullString)
Button& = FindWindowEx(window&, 0&, "_AOL_Icon", vbNullString)
If window& = 0 Then Exit Sub
Call SendMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(Button&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub WindowClose(window&)
Call PostMessage(window&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub WindowHide(window&)
Call ShowWindow(window&, SW_HIDE)
End Sub
Public Sub WindowShow(window&)
Call ShowWindow(window&, SW_SHOW)
End Sub
Public Sub HideWelcome()
Call ShowWindow(FindWelcome, SW_HIDE)
End Sub
Public Sub HideAOL()
Call ShowWindow(aol, SW_HIDE)
End Sub
Public Sub ShowWelcome()
Call ShowWindow(FindWelcome, SW_SHOW)
End Sub
Public Sub ShowAOL()
Call ShowWindow(aol, SW_SHOW)
End Sub
Public Function FindUploadWin()
Dim window&, dagauge&, dastatic&, caption&
window& = FindWindow("_AOL_Modal", vbNullString)
Do
DoEvents
dagauge& = FindWindowEx(window&, 0&, "_AOL_Gauge", vbNullString)
dastatic& = FindWindowEx(window&, 0&, "_AOL_Static", vbNullString)
caption& = InStr(GetText(dastatic&), " Upload")
If window& <> 0 And dagauge& <> 0 And dastatic& <> 0 And caption& <> 0 Then
FindUploadWin = window&
Exit Function
End If
window& = FindWindowEx(aol, window&, "_AOL_Modal", vbNullString)
Loop Until window& = 0
FindUploadWin = 0
End Function
Public Sub WindowMinimize(window&)
Call ShowWindow(window&, SW_MINIMIZE)
End Sub
Public Sub WindowMaximize(window&)
Call ShowWindow(FindUploadWin, SW_MAXIMIZE)
End Sub
Public Sub WindowRestore(window&)
Call ShowWindow(window&, SW_RESTORE)
End Sub
Public Sub UpchatOn()
Call EnableWindow(FindUploadWin, 0)
Call EnableWindow(aol, 1)
Call ShowWindow(FindUploadWin, SW_MINIMIZE)
End Sub
Public Sub UpchatOff()
Call EnableWindow(FindUploadWin, 1)
Call EnableWindow(aol, 0)
Call ShowWindow(FindUploadWin, SW_RESTORE)
End Sub
Public Function UploadTime()
'this returns time left in the upload
Dim upwin&, dastatic&, datime$
upwin& = FindUploadWin
If upwin& = 0 Then
UploadTime = "Not Uploading"
Exit Function
End If
dastatic& = FindWindowEx(upwin&, 0&, "_AOL_Static", vbNullString)
dastatic& = FindWindowEx(upwin&, dastatic&, "_AOL_Static", vbNullString)
datime$ = GetText(dastatic&)
datime$ = ReplaceString(datime$, "About ", "")
datime$ = ReplaceString(datime$, " remaining.", "")
UploadTime = datime$
End Function
Public Function FindDownloadWin()
Dim window&, dagauge&, dastatic&, caption&
window& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
dagauge& = FindWindowEx(window&, 0&, "_AOL_Gauge", vbNullString)
dastatic& = FindWindowEx(window&, 0&, "_AOL_Static", vbNullString)
caption& = InStr(GetText(dastatic&), "Download")
If window& <> 0 And dagauge& <> 0 And dastatic& <> 0 And caption& <> 0 Then
FindDownloadWin = window&
Exit Function
End If
window& = FindWindowEx(mdi, window&, "AOL Child", vbNullString)
Loop Until window& = 0
FindDownloadWin = 0
End Function
Public Function DownloadTime()
'this returns time left in the d\l
Dim downwin&, dastatic&, datime$
downwin& = FindDownloadWin
If downwin& = 0 Then
DownloadTime = "Not Downloading"
Exit Function
End If
dastatic& = FindWindowEx(downwin&, 0&, "_AOL_Static", vbNullString)
dastatic& = FindWindowEx(downwin&, dastatic&, "_AOL_Static", vbNullString)
datime$ = GetText(dastatic&)
datime$ = ReplaceString(datime$, "About ", "")
datime$ = ReplaceString(datime$, " remaining.", "")
DownloadTime = datime$
End Function
Public Sub HideAim()
Dim window&
window& = FindWindow("_Oscar_BuddylistWin", vbNullString)
Call ShowWindow(window&, SW_HIDE)
End Sub
Public Sub ShowAim()
Dim window&
window& = FindWindow("_Oscar_BuddylistWin", vbNullString)
Call ShowWindow(window&, SW_SHOW)
End Sub
Public Function DownloadPercent()
'returns the percent the download is at
Dim downwin&, dastatic&, datime$, Number$
downwin& = FindDownloadWin
If downwin& = 0 Then
DownloadPercent = "Not Downloading"
Exit Function
End If
dastatic& = FindWindowEx(downwin&, 0&, "_AOL_Static", vbNullString)
dastatic& = FindWindowEx(downwin&, dastatic&, "_AOL_Static", vbNullString)
datime$ = GetCaption(downwin&)
Number$ = InStr(datime$, " - ")
datime$ = Mid(datime$, Number$ + 3, Len(datime$))
DownloadPercent = datime$
End Function
Public Function UploadPercent()
'returns the percent the upload is at
Dim upwin&, dastatic&, datime$, Number$
upwin& = FindUploadWin
If upwin& = 0 Then
UploadPercent = "Not Uploading"
Exit Function
End If
dastatic& = FindWindowEx(upwin&, 0&, "_AOL_Static", vbNullString)
dastatic& = FindWindowEx(upwin&, dastatic&, "_AOL_Static", vbNullString)
datime$ = GetCaption(upwin&)
Number$ = InStr(datime$, " - ")
datime$ = Mid(datime$, Number$ + 3, Len(datime$))
UploadPercent = datime$
End Function
Public Sub CollectFindAChat(thelist As ListBox, townsquare As Boolean, arts As Boolean, friends As Boolean, life As Boolean, news As Boolean, places As Boolean, romance As Boolean, specialintrest As Boolean, uk As Boolean, canada As Boolean, japan As Boolean, brazil As Boolean, Optional pause$ = 0.8)
Dim findachat&, chatting&, list&, abort As Boolean, iconcount&, rightcount&, leftlist&, rightlist&, wcicon&, nobutton&, nowindow&
Call RunToolbar(9, "F")
Do
DoEvents
If FindSignon <> 0 Then Exit Sub
findachat& = FindWindowEx(mdi, 0&, "AOL Child", "Find a Chat")
leftlist& = FindWindowEx(findachat&, 0&, "_AOL_Listbox", vbNullString)
rightlist& = FindWindowEx(findachat&, leftlist&, "_AOL_Listbox", vbNullString)
wcicon& = FindWindowEx(findachat&, 0&, "_AOL_Icon", vbNullString)
For iconcount& = 1 To 8
wcicon& = FindWindowEx(findachat&, wcicon&, "_AOL_Icon", vbNullString)
Next iconcount&
Loop Until findachat& <> 0 And leftlist& <> 0 And rightlist& <> 0 And wcicon& <> 0
Call WaitForListToLoad(leftlist&)
Call WaitForListToLoad(rightlist&)
choose:
If townsquare = True Then
townsquare = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 0, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf arts = True Then
arts = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 1, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf friends = True Then
friends = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 2, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf life = True Then
life = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 3, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf news = True Then
news = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 4, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf places = True Then
places = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 5, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf romance = True Then
romance = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 6, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf specialintrest = True Then
specialintrest = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 7, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf uk = True Then
uk = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 8, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf canada = True Then
canada = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 9, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf japan = True Then
japan = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 10, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
GoTo collect:
ElseIf brazil = True Then
brazil = False
Call SendMessageLong(leftlist&, LB_SETCURSEL, 11, 0&)
Call PostMessage(leftlist&, WM_LBUTTONDBLCLK, 0&, 0&)
Call WaitForListToLoad(rightlist&)
abort = True
GoTo collect:
End If
Exit Sub
collect:
For rightcount& = 0 To SendMessageLong(rightlist&, LB_GETCOUNT, 0&, 0&) - 1
If GetWCCount(rightlist&, rightcount&) = "" Then GoTo skipit:
Call SendMessageLong(rightlist&, LB_SETCURSEL, rightcount&, 0&)
Call PostMessage(wcicon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(wcicon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
nowindow& = FindWindow("#32770", "America Online")
nobutton& = FindWindowEx(nowindow&, 0&, "Button", "OK")
chatting& = FindWindowEx(mdi, 0&, "AOL Child", "Who's Chatting")
list& = FindWindowEx(chatting&, 0&, "_AOL_Listbox", vbNullString)
Loop Until (chatting& <> 0 And list& <> 0) Or (nowindow& <> 0 And nobutton& <> 0)
If nowindow& <> 0 Then
Call SendMessage(nobutton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(nobutton&, WM_KEYUP, VK_SPACE, 0&)
TimeOut (pause$)
GoTo emptyy:
End If
Call WaitForListToLoad(list&)
Call AddWhosChatting(thelist)
Call WindowClose(chatting&)
emptyy:
TimeOut (pause$)
skipit:
Next rightcount&
GoTo choose:
End Sub

Public Function Locate(sn$)
Dim reswin&, resbut&, win&, caption$, dastatic&, where$, offwin&, offbut&
Call KW2("aol://3548:" & sn$)
look:
win& = FindWindowEx(mdi, 0&, "AOL Child", vbNullString)
Do
DoEvents
caption$ = InStr(GetCaption(win&), "Locate")
If caption <> 0 Then
reswin& = win&
GoTo ok:
End If
offwin& = FindWindow("#32770", "America Online")
offbut& = FindWindowEx(offwin&, 0&, "Button", vbNullString)
If offwin& <> 0 And offbut& <> 0 Then
GoTo ok:
End If
win& = FindWindowEx(win&, 0&, "AOL Child", vbNullString)
Loop Until win& = 0
GoTo look:
ok:
If offwin& <> 0 Then
Call SendMessage(offbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(offbut&, WM_KEYUP, VK_SPACE, 0&)
Locate = sn$ & " is not signed on."
Exit Function
End If
Do
DoEvents
dastatic& = FindWindowEx(reswin&, 0&, "_AOL_Static", vbNullString)
where$ = GetText(dastatic&)
Loop Until where$ <> ""
Locate = where$
Call WindowClose(reswin&)
End Function
Public Function GetAIMsn()
Dim window&, caption$
window& = FindWindow("_Oscar_BuddylistWin", vbNullString)
If window& = 0 Then
GetAIMsn = "not signed on"
Exit Function
End If
caption$ = GetCaption(window&)
GetAIMsn = Left(caption$, InStr(caption$, "'") - 1)
End Function

Public Sub SaveListBox(directory As String, thelist As ListBox)
    'dos32 to the rescue
    Dim SaveList As Long
    On Error Resume Next
    Open directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.list(SaveList&)
    Next SaveList&
    Close #1
End Sub

Public Sub Save2ListBoxes(directory As String, ListA As ListBox, ListB As ListBox)
    'curtosy of dos
    Dim SaveLists As Long
    On Error Resume Next
    Open directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.list(SaveLists&) & "*" & ListB.list(SaveLists)
    Next SaveLists&
    Close #1
End Sub
Public Sub Loadlistbox(directory As String, thelist As ListBox)
    '::yawns:: thanks dos
    Dim MyString As String
    On Error Resume Next
    Open directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        thelist.AddItem MyString$
    Wend
    Close #1
End Sub

Public Sub Load2listboxes(directory As String, ListA As ListBox, ListB As ListBox)
    'dos once again
    Dim MyString As String, aString As String, bString As String
    On Error Resume Next
    Open directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        aString$ = Left(MyString$, InStr(MyString$, "*") - 1)
        bString$ = Right(MyString$, Len(MyString$) - InStr(MyString$, "*"))
        DoEvents
        ListA.AddItem aString$
        ListB.AddItem bString$
    Wend
    Close #1
End Sub
Sub LoadText(txtLoad As TextBox, path As String)
'dos32
    Dim TextString As String
    On Error Resume Next
    Open path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txtLoad.text = TextString$
End Sub

Sub SaveText(txtSave As TextBox, path As String)
'dos32
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.text
    Open path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
Public Sub SendListMail(dalist As ListBox, subject$, message$)
Dim name&, names$
For name& = 0 To dalist.ListCount - 1
names$ = names$ & dalist.list(name&) & ", "
Next name&
Call SendMail(Left(names$, Len(names$) - 1), subject$, message$)
End Sub
Public Sub SendListMailBCC(dalist As ListBox, subject$, message$)
Dim name&, names$
For name& = 0 To dalist.ListCount - 1
names$ = names$ & "(" & dalist.list(name&) & "), "
Next name&
Call SendMail(Left(names$, Len(names$) - 1), subject$, message$)
End Sub
Public Function HackerText(dastring$)
Dim letter$, sentence$, count&
dastring$ = UCase(dastring$)
For count& = 1 To Len(dastring$)
letter$ = Mid(dastring$, count&, 1)
If letter$ = "A" Then
letter$ = "a"
ElseIf letter$ = "E" Then
letter$ = "e"
ElseIf letter = "I" Then
letter$ = "i"
ElseIf letter$ = "O" Then
letter$ = "o"
ElseIf letter$ = "U" Then
letter$ = "u"
End If
sentence$ = sentence$ & letter$
Next count&
HackerText = sentence$
End Function
Public Function CheckIfGuest()
Dim guestwin&, guestbut&, notwin&
Call RunToolbar(2, "A")
Do
DoEvents
guestwin& = FindWindow("#32770", "America Online")
guestbut& = FindWindowEx(guestwin&, 0&, "Button", vbNullString)
notwin& = FindWindowEx(mdi, 0&, "AOL Child", "Address Book")
Loop Until (guestwin& <> 0 And guestbut& <> 0) Or notwin& <> 0
If guestwin& <> 0 Then
Call SendMessage(guestbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(guestbut&, WM_KEYUP, VK_SPACE, 0&)
CheckIfGuest = True
Else
Call WindowClose(notwin&)
CheckIfGuest = False
End If
End Function
Public Function CheckIfSNisAvailible(sn$)
Dim first&, subwin&, subbut&, firsticon&, Second&, secondicon&, Third&, thirdicon&, edit&, Fourth&, taken&, takenbut&
If CheckIfGuest = True Then
MsgBox "You cannot use this feature while signed on as guest"
CheckIfSNisAvailible = False
Exit Function
End If
Call RunToolbar(5, "n")
Do
DoEvents
subwin& = FindWindow("_AOL_Modal", vbNullString)
subbut& = FindWindowEx(subwin&, 0&, "_AOL_Icon", vbNullString)
first& = FindWindowEx(mdi, 0&, "AOL Child", "AOL Screen Names")
Loop Until first& <> 0 Or (subwin& <> 0 And subbut& <> 0)
If subwin& <> 0 Then
Call SendMessage(subbut&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(subbut&, WM_LBUTTONUP, 0, 0&)
MsgBox "You cannot use this feature while signed on a sub acct."
CheckIfSNisAvailible = False
Exit Function
End If
Do
DoEvents
TimeOut 1
first& = FindWindowEx(mdi, 0&, "AOL Child", "AOL Screen Names")
firsticon& = FindWindowEx(first&, 0&, "_AOL_Icon", vbNullString)
firsticon& = FindWindowEx(first&, firsticon&, "_AOL_Icon", vbNullString)
Loop Until first& <> 0 And firsticon& <> 0
Call PostMessage(firsticon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(firsticon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
Second& = FindWindow("_AOL_Modal", "Create a Screen Name")
secondicon& = FindWindowEx(Second&, 0&, "_AOL_Icon", vbNullString)
Loop Until Second& <> 0 And secondicon& <> 0
Call SendMessage(secondicon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(secondicon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
Third& = FindWindow("_AOL_Modal", "Step 1 of 4: Choose a Screen Name")
thirdicon& = FindWindowEx(Third&, 0&, "_AOL_Icon", vbNullString)
edit& = FindWindowEx(Third&, 0&, "_AOL_Edit", vbNullString)
Loop Until Third& <> 0 And thirdicon& <> 0 And edit& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, sn$)
Call SendMessage(thirdicon&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessage(thirdicon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
taken& = FindWindow("#32770", "America Online")
takenbut& = FindWindowEx(taken&, 0&, "Button", vbNullString)
Fourth& = FindWindow("_AOL_Modal", "Step 2 of 4: Choose a password")
Loop Until Fourth& <> 0 Or (taken& <> 0 And takenbut& <> 0)
If taken& <> 0 Then
Call SendMessage(takenbut&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(takenbut&, WM_KEYUP, VK_SPACE, 0&)
thirdicon& = FindWindowEx(Third&, thirdicon&, "_AOL_Icon", vbNullString)
Call PostMessage(thirdicon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(thirdicon&, WM_LBUTTONUP, 0&, 0&)
Call WindowClose(first&)
CheckIfSNisAvailible = False
Exit Function
End If
Call WindowClose(Fourth&)
secondicon& = FindWindowEx(Second&, secondicon&, "_AOL_Icon", vbNullString)
Call PostMessage(secondicon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(secondicon&, WM_LBUTTONUP, 0&, 0&)
Call WindowClose(first&)
CheckIfSNisAvailible = True
End Function
Public Function FirstAndLastLetterBold(thestring$)
Dim count&, first$, Second$, sentence$, Third$, char&, sentence2$
sentence$ = "<b>" & Left(thestring, 1) & "</b>"
For count& = 2 To Len(thestring$) - 1
first$ = Mid(thestring, count&, 1)
Second$ = Mid(thestring, count& + 1, 1)
If count& = 2 Then GoTo here:
Third$ = Mid(thestring, count& - 1, 1)
here:
If Second$ = " " Then
sentence$ = sentence$ & "<b>" & first
GoTo bottom:
End If
If Third$ = " " And first$ <> " " Then
sentence$ = sentence$ & first$ & "</b>"
GoTo bottom:
End If
sentence$ = sentence$ & first$
bottom:
Next count&
FirstAndLastLetterBold = sentence$ & "<b>" & Right(thestring$, 1)
End Function
Public Sub ChangeAOLCaption(newcaption$)
Dim amer&
amer& = aol
If amer& = 0 Then Exit Sub
Call SetCaption(amer&, newcaption$)
End Sub
Public Sub ChangeRoomCaption(newcaption$)
Dim room&
room& = FindChat
If room& = 0 Then Exit Sub
Call SetCaption(room&, newcaption$)
End Sub
Public Sub GhostOn()
Dim blwin&, icon&, check&, check1&, check2&, check3&, check4&, check5&, check6&, save&, save1&, save2&, save3&, count&, editwin&, okwin&, okbut&
Call KW2("buddylist")
Do
DoEvents
blwin& = FindBuddyListEdit
Loop Until blwin& <> 0
icon& = FindWindowEx(blwin&, 0&, "_AOL_Icon", vbNullString)
For count& = 1 To 4
icon& = FindWindowEx(blwin&, icon&, "_AOL_Icon", vbNullString)
Next count&
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
editwin& = FindWindowEx(mdi, 0&, "AOL Child", "Privacy Preferences")
check& = FindWindowEx(editwin&, 0&, "_AOL_Checkbox", vbNullString)
check1& = FindWindowEx(editwin&, check&, "_AOL_Checkbox", vbNullString)
check2& = FindWindowEx(editwin&, check1&, "_AOL_Checkbox", vbNullString)
check3& = FindWindowEx(editwin&, check2&, "_AOL_Checkbox", vbNullString)
check4& = FindWindowEx(editwin&, check3&, "_AOL_Checkbox", vbNullString)
save& = FindWindowEx(editwin&, 0&, "_AOL_Icon", vbNullString)
save1& = FindWindowEx(editwin&, save&, "_AOL_Icon", vbNullString)
save2& = FindWindowEx(editwin&, save1&, "_AOL_Icon", vbNullString)
save3& = FindWindowEx(editwin&, save2&, "_AOL_Icon", vbNullString)
check5& = FindWindowEx(editwin&, check4&, "_AOL_Checkbox", vbNullString)
check6& = FindWindowEx(editwin&, check5&, "_AOL_Checkbox", vbNullString)
Loop Until editwin& <> 0 And check& <> 0 And check1& <> 0 And check2& <> 0 And check3& <> 0 And check4& <> 0 And save& <> 0 And save1& <> 0 And save2& <> 0 And save3& <> 0 And check5& <> 0 And check6& <> 0
DoEvents
Call PostMessage(check4&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(check4&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(check6&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(check6&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(save3&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(save3&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
editwin& = FindWindowEx(mdi, 0&, "AOL Child", "Privacy Preferences")
Loop Until IsWindowVisible(editwin&) = False
Do
DoEvents
okwin& = FindWindow("#32770", "America Online")
okbut& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until okwin& <> 0 And okbut& <> 0
Call PostMessage(okbut&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(okbut&, WM_LBUTTONUP, 0&, 0&)
Call WindowClose(blwin&)
End Sub
Public Sub GhostOff()
Dim blwin&, icon&, editwin&, count&, save&, save1&, save2&, save3&, check&, okwin&, okbut&
Call KW2("buddylist")
Do
DoEvents
blwin& = FindBuddyListEdit
Loop Until blwin& <> 0
icon& = FindWindowEx(blwin&, 0&, "_AOL_Icon", vbNullString)
For count& = 1 To 4
icon& = FindWindowEx(blwin&, icon&, "_AOL_Icon", vbNullString)
Next count&
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
editwin& = FindWindowEx(mdi, 0&, "AOL Child", "Privacy Preferences")
check& = FindWindowEx(editwin&, 0&, "_AOL_Checkbox", vbNullString)
save& = FindWindowEx(editwin&, 0&, "_AOL_Icon", vbNullString)
save1& = FindWindowEx(editwin&, save&, "_AOL_Icon", vbNullString)
save2& = FindWindowEx(editwin&, save1&, "_AOL_Icon", vbNullString)
save3& = FindWindowEx(editwin&, save2&, "_AOL_Icon", vbNullString)
Loop Until editwin& <> 0 And check& <> 0 And save& <> 0 And save1& <> 0 And save2& <> 0 And save3& <> 0
DoEvents
Call PostMessage(check&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(check&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(save3&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(save3&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
editwin& = FindWindowEx(mdi, 0&, "AOL Child", "Privacy Preferences")
Loop Until IsWindowVisible(editwin&) = False
Do
DoEvents
okwin& = FindWindow("#32770", "America Online")
okbut& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until okwin& <> 0 And okbut& <> 0
Call PostMessage(okbut&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(okbut&, WM_LBUTTONUP, 0&, 0&)
Call WindowClose(blwin&)
End Sub
Public Sub ProfileTag(name$, location$, birthday$, marital$, hobbies$, computers$, occupation$, quote$)
Dim prowin&, nameedit&, locationedit&, birthdayedit&, maritaledit&, hobbiesedit&, computersedit&, occupationedit&, quoteedit&, icon&, update&, okwin&, okbut&
Call RunToolbar(5, "y")
Do
DoEvents
prowin& = FindWindowEx(mdi, 0&, "AOL Child", "Edit Your Online Profile")
nameedit& = FindWindowEx(prowin&, 0&, "_AOL_Edit", vbNullString)
locationedit& = FindWindowEx(prowin&, nameedit&, "_AOL_Edit", vbNullString)
birthdayedit& = FindWindowEx(prowin&, locationedit&, "_AOL_Edit", vbNullString)
maritaledit& = FindWindowEx(prowin&, birthdayedit&, "_AOL_Edit", vbNullString)
hobbiesedit& = FindWindowEx(prowin&, maritaledit&, "_AOL_Edit", vbNullString)
computersedit& = FindWindowEx(prowin&, hobbiesedit&, "_AOL_Edit", vbNullString)
occupationedit& = FindWindowEx(prowin&, computersedit&, "_AOL_Edit", vbNullString)
quoteedit& = FindWindowEx(prowin&, occupationedit&, "_AOL_Edit", vbNullString)
icon& = FindWindowEx(prowin&, 0&, "_AOL_Icon", vbNullString)
update& = FindWindowEx(prowin&, icon&, "_AOL_Icon", vbNullString)
Loop Until prowin& <> 0 And nameedit& <> 0 And locationedit& <> 0 And birthdayedit& <> 0 And maritaledit& <> 0 And hobbiesedit& <> 0 And computersedit& <> 0 And occupationedit& <> 0 And quoteedit& <> 0 And icon& <> 0 And update& <> 0
Call SendMessageByString(nameedit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(locationedit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(birthdayedit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(maritaledit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(hobbiesedit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(computersedit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(occupationedit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(quoteedit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(nameedit&, WM_SETTEXT, 0&, name$)
Call SendMessageByString(locationedit&, WM_SETTEXT, 0&, location$)
Call SendMessageByString(birthdayedit&, WM_SETTEXT, 0&, birthday$)
Call SendMessageByString(maritaledit&, WM_SETTEXT, 0&, marital$)
Call SendMessageByString(hobbiesedit&, WM_SETTEXT, 0&, hobbies$)
Call SendMessageByString(computersedit&, WM_SETTEXT, 0&, computers$)
Call SendMessageByString(occupationedit&, WM_SETTEXT, 0&, occupation$)
Call SendMessageByString(quoteedit&, WM_SETTEXT, 0&, quote$)
Call PostMessage(update&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(update&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
okwin& = FindWindow("#32770", "America Online")
okbut& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until okwin& <> 0 And okbut& <> 0
Call PostMessage(okbut&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(okbut&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub ChangePW(oldpw$, newpw$)
Dim modal&, win&, ModalIcon&, modalicon2&, winicon&, old&, new1&, new2&, okwin&, okbut&
Call RunToolbar(5, "a")
Do
DoEvents
modal& = FindWindow("_AOL_Modal", vbNullString)
ModalIcon& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
modalicon2& = FindWindowEx(modal&, ModalIcon&, "_AOL_Icon", vbNullString)
Loop Until modal& <> 0 And ModalIcon& <> 0 And modalicon2& <> 0
Call PostMessage(ModalIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ModalIcon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
win& = FindWindow("_AOL_Modal", "Change Your Password")
winicon& = FindWindowEx(win&, 0&, "_AOL_Icon", vbNullString)
old& = FindWindowEx(win&, 0&, "_AOL_Edit", vbNullString)
new1& = FindWindowEx(win&, old&, "_AOL_Edit", vbNullString)
new2& = FindWindowEx(win&, new1&, "_AOL_Edit", vbNullString)
Loop Until win& <> 0 And winicon& <> 0 And old& <> 0 And new1& <> 0 And new2& <> 0
Call SendMessageByString(old&, WM_SETTEXT, 0&, oldpw$)
Call SendMessageByString(new1&, WM_SETTEXT, 0&, newpw$)
Call SendMessageByString(new2&, WM_SETTEXT, 0&, newpw$)
Call PostMessage(winicon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(winicon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
okwin& = FindWindow("#32770", "America Online")
okbut& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
Loop Until okwin& <> 0 And okbut& <> 0
Call PostMessage(okbut&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(okbut&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(modalicon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(modalicon2&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub IMx(sn$)
'IM ignore sumone
Call IM("$IM_OFF " & sn$, "{S goodbye")
End Sub
Public Sub IMunx(sn$)
'IM unignore sumone
Call IM("$IM_ON " & sn$, "{S welcome")
End Sub
Public Sub ListRemoveBlanks(thelist As ListBox)
Dim count&, count2&
If thelist.ListCount = 0 Then Exit Sub
Do
DoEvents
count& = 1
Do
DoEvents
If thelist.list(count&) = "" Then thelist.RemoveItem (count&)
count& = count& + 1
count2& = thelist.ListCount
Loop Until count& >= count2&
Loop Until InStr(thelist.hwnd, "") = 0
End Sub
Public Sub ListRemoveDupes(thelist As ListBox)
Dim count&, count2&, count3&
If thelist.ListCount = 0 Then Exit Sub
For count& = 0 To thelist.ListCount - 1
DoEvents
For count2& = count& + 1 To thelist.ListCount - 1
DoEvents
If thelist.list(count&) = thelist.list(count2&) Then thelist.RemoveItem (count2&)
Next count2&
Next count&
End Sub
Public Function IsAOL40()
'checks if user is on 4.0
Dim toolbar1&, toolbar2&, ComboBox&, editwin&
If aol = 0 Then
IsAOL40 = False
Exit Function
End If
toolbar1& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
ComboBox& = FindWindowEx(toolbar2&, 0&, "_AOL_ComboBox", vbNullString)
editwin& = FindWindowEx(ComboBox&, 0&, "Edit", vbNullString)
If toolbar1& <> 0 And toolbar2& <> 0 And ComboBox& <> 0 And editwin& <> 0 Then
IsAOL40 = True
Exit Function
Else
IsAOL40 = False
End If
End Function
Public Sub KW2(thekw$)
'this method calls a 'keyword' window and does it that way
Dim window&, edit&, icon2&, toolbar1&, toolbar2&, icon&, count&
toolbar1& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
icon2& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
For count& = 0 To 18
icon2& = FindWindowEx(toolbar2&, icon2&, "_AOL_Icon", vbNullString)
Next count&
Call PostMessage(icon2&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon2&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
window& = FindWindowEx(mdi, 0&, "AOL Child", "Keyword")
edit& = FindWindowEx(window&, 0&, "_AOL_Edit", vbNullString)
icon& = FindWindowEx(window&, 0&, "_AOL_Icon", vbNullString)
Loop Until window& <> 0 And edit& <> 0 And icon& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, thekw$)
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function DoubleText(thestring$)
Dim count&, letter$, sentence$
For count& = 1 To Len(thestring$)
letter$ = Mid(thestring$, count&, 1)
sentence$ = sentence$ & letter$ & letter$
Next count&
DoubleText = sentence$
End Function
Public Function SpacedText(thestring$)
Dim count&, letter$, sentence$
For count& = 1 To Len(thestring$)
letter$ = Mid(thestring$, count&, 1)
sentence$ = sentence$ & letter$ & " "
Next count&
SpacedText = sentence$
End Function
Public Sub SpiralScroll(thestring$, Optional pause$ = 0.8)
Dim sentence$
sentence$ = thestring$
SendText (sentence$)
Do
DoEvents
sentence$ = Right(sentence$, Len(sentence$) - 1) & Left(sentence$, 1)
TimeOut (pause$)
SendText (sentence$)
Loop Until sentence$ = thestring$
End Sub
Public Sub WindowsShutdown()
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub
Public Sub WindowsReboot()
Call ExitWindowsEx(EWX_REBOOT, 0)
End Sub
Public Function UploadName()
'this returns the name of the file being uploaded
Dim upwin&, dastatic&, daname$
upwin& = FindUploadWin
If upwin& = 0 Then
UploadName = "Not Uploading"
Exit Function
End If
dastatic& = FindWindowEx(upwin&, 0&, "_AOL_Static", vbNullString)
daname$ = GetText(dastatic&)
daname$ = ReplaceString(daname$, "Now Uploading ", "")
UploadName = daname$
End Function
Public Function DownloadName()
'returns the name of the file being downloaded
Dim downwin&, dastatic&, daname$
downwin& = FindDownloadWin
If downwin& = 0 Then
DownloadName = "Not Downloading"
Exit Function
End If
dastatic& = FindWindowEx(downwin&, 0&, "_AOL_Static", vbNullString)
daname$ = GetText(dastatic&)
daname$ = ReplaceString(daname$, "Now Downloading ", "")
DownloadName = daname$
End Function
Public Sub NewUserReset(thesn$, thepath$)
'heh, very unhappy with this sub. it DOES work but...
'it takes about 2-3 minutes to reset it. although its slow
'it works on all aol versions. Case and space sensitive.
Dim path$, text$, newu$
path$ = thepath$ & "\idb\main.idx"
If Dir(path$) = "" Then Exit Sub
newu$ = "NewUser"
If Len(thesn$) < 7 Then MsgBox "The name must be 7 or more letters": Exit Sub
If Len(thesn$) > 7 Then
Do
DoEvents
newu$ = newu$ & " "
Loop Until Len(newu$) = Len(thesn$)
End If
Open path$ For Binary As #1
text$ = Input(LOF(1), #1)
Close #1
If InStr(text$, thesn$) = 0 Then Exit Sub
text$ = ReplaceString(text$, thesn$, newu$)
Open path$ For Output As #1
Print #1, text$
Close #1
End Sub

Public Function GetRoomName()
Dim room$
room$ = GetCaption(FindChat)
If room$ = "" Then
GetRoomName = "Not in a room"
Else
GetRoomName = room$
End If
End Function
Public Sub ClickButton(dabutton&)
Call SendMessage(dabutton&, WM_KEYDOWN, VK_SPACE, 0&)
Call SendMessage(dabutton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Function GetWCCount(aollist&, index&)
'/parts taken from dos
On Error Resume Next
Dim cProcess As Long, itmHold As Long, screenname As String
Dim psnHold As Long, rBytes As Long, room As Long, thecount$
Dim rList As Long, sThread As Long, mThread As Long, lett$, num&, eh&
Dim Ta As Long, Ta2 As Long
rList& = aollist&
If rList& = 0& Then Exit Function
sThread& = GetWindowThreadProcessId(rList, cProcess&)
mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
screenname$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
itmHold& = itmHold& + 24
Call ReadProcessMemory(mThread&, itmHold&, screenname$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal screenname$, 4)
psnHold& = psnHold& + 6
screenname$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, screenname$, Len(screenname$), rBytes&)
Ta& = InStr(1, screenname$, Chr(9))
Ta2& = InStr(Ta& + 1, screenname$, Chr(9))
screenname$ = Mid(screenname$, Ta& + 1, Ta2& - 2)
screenname$ = Mid(screenname$, 4, 2)
For num& = 1 To Len(screenname$)
lett$ = Mid(screenname$, num&, 1)
If IsNumeric(lett$) = True Then
thecount$ = thecount$ & lett$
End If
Next num&
GetWCCount = thecount$
thecount$ = ""
Call CloseHandle(mThread)
End If
'//---
End Function
Public Sub SendBuddyInvite(thelist As ListBox, message$, url$)
'this will send a url invite. not a chatroom
Dim bud&, icon&, inwin&, edit&, edit2&, edit3&, check&, names$, count&
Call KW2("buddy chat")
Do
DoEvents
bud& = FindWindowEx(mdi, 0&, "AOL Child", "Buddy Chat")
edit& = FindWindowEx(bud&, 0&, "_AOL_Edit", vbNullString)
edit2& = FindWindowEx(bud&, edit&, "_AOL_Edit", vbNullString)
edit3& = FindWindowEx(bud&, edit2&, "_AOL_Edit", vbNullString)
edit3& = FindWindowEx(bud&, edit3&, "_AOL_Edit", vbNullString)
icon& = FindWindowEx(bud&, 0&, "_AOL_Icon", vbNullString)
check& = FindWindowEx(bud&, 0&, "_AOL_Checkbox", vbNullString)
Loop Until bud& <> 0 And edit& <> 0 And edit2& <> 0 And check& <> 0 And edit3& <> 0 And icon& <> 0
For count& = 0 To thelist.ListCount - 1
names$ = names$ & thelist.list(count&) & ","
Next count&
Call SendMessageByString(edit&, WM_SETTEXT, 0&, names$)
Call SendMessageByString(edit2&, WM_SETTEXT, 0&, message$)
check& = FindWindowEx(bud&, check&, "_AOL_Checkbox", vbNullString)
Call PostMessage(check&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(check&, WM_LBUTTONUP, 0&, 0&)
TimeOut 0.6
Call SendMessageByString(edit3&, WM_SETTEXT, 0&, url$)
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
inwin& = FindWindowEx(mdi, 0&, "AOL Child", "Invitation from: " & User)
Loop Until inwin& <> 0
Call WindowClose(inwin&)
End Sub
Public Sub SignOn25(sn$, pw$)
Dim aol&, mdi&, win&, pic&, combo&, icon&, icon2&, icon3&, modal&, micon&, micon2&, medit&, medit2&
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
win& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Do
DoEvents
pic& = FindWindowEx(win&, 0&, "_AOL_Glyph", vbNullString)
combo& = FindWindowEx(win&, 0&, "_AOL_Combobox", vbNullString)
icon& = FindWindowEx(win&, 0&, "_AOL_Icon", vbNullString)
icon2& = FindWindowEx(win&, icon&, "_AOL_Icon", vbNullString)
icon3& = FindWindowEx(win&, icon2&, "_AOL_Icon", vbNullString)
If win& <> 0 And pic& <> 0 And combo& <> 0 And icon& <> 0 And icon2& <> 0 And icon3& <> 0 Then GoTo foundit:
win& = FindWindowEx(win&, 0&, "AOL Child", vbNullString)
Loop Until win& = 0
foundit:
Do
DoEvents
Call SendMessage(icon&, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon&, WM_LBUTTONUP, 0, 0&)
TimeOut 0.6
Loop Until IsWindowVisible(icon&) = False
Do
DoEvents
modal& = FindWindow("_AOL_Modal", vbNullString)
micon& = FindWindowEx(modal&, 0&, "_AOL_Button", vbNullString)
micon2& = FindWindowEx(modal&, micon&, "_AOL_Button", vbNullString)
medit& = FindWindowEx(modal&, 0&, "_AOL_Edit", vbNullString)
medit2& = FindWindowEx(modal&, medit&, "_AOL_Edit", vbNullString)
Loop Until modal& <> 0 And micon& <> 0 And micon2& <> 0 And medit& <> 0 And medit2& <> 0
Call SendMessageByString(medit&, WM_SETTEXT, 0&, sn$)
Call SendMessageByString(medit2&, WM_SETTEXT, 0&, pw$)
Call PostMessage(micon2&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(micon2&, WM_LBUTTONUP, 0, 0&)
End Sub
Public Sub SendText25(wuttosay$)
Dim aol&, mdi&, room&, pic&, view&, list&, edit&
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
room& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Do
DoEvents
pic& = FindWindowEx(room&, 0&, "_AOL_Glyph", vbNullString)
view& = FindWindowEx(room&, 0&, "_AOL_View", vbNullString)
list& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
edit& = FindWindowEx(room&, 0&, "_AOL_Edit", vbNullString)
If room& <> 0 And pic& <> 0 And edit& <> 0 And view& <> 0 And list& <> 0 Then GoTo foundit:
room& = FindWindowEx(mdi&, room&, "AOL Child", vbNullString)
Loop Until room& = 0
room& = 0: Exit Sub
foundit:
Call SendMessageByString(edit&, WM_SETTEXT, 0&, wuttosay$)
Call SendMessageLong(edit&, WM_CHAR, ENTER_KEY, 0&)
End Sub

Public Function FindChat25()
Dim aol&, mdi&, room&, pic&, view&, list&, edit&
aol& = FindWindow("AOL Frame25", vbNullString)
mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
room& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
Do
DoEvents
pic& = FindWindowEx(room&, 0&, "_AOL_Glyph", vbNullString)
view& = FindWindowEx(room&, 0&, "_AOL_View", vbNullString)
list& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
edit& = FindWindowEx(room&, 0&, "_AOL_Edit", vbNullString)
If room& <> 0 And pic& <> 0 And edit& <> 0 And view& <> 0 And list& <> 0 Then
FindChat25 = room&
Exit Function
End If
room& = FindWindowEx(mdi&, room&, "AOL Child", vbNullString)
Loop Until room& = 0
FindChat25 = 0
End Function
Public Function LastChatLine25()
Dim room&, view&, text$, letter$, sentence$, count&
room& = FindChat25
If room& = 0 Then
LastChatLine25 = ""
Exit Function
End If
view& = FindWindowEx(room&, 0&, "_AOL_View", vbNullString)
text$ = GetText(view&)
For count& = 1 To Len(text$)
letter$ = Mid(text$, count&, 1)
If letter$ = Chr(13) Then
sentence$ = ""
GoTo bottom:
End If
sentence$ = sentence$ & letter$
bottom:
Next count&
LastChatLine25 = sentence$
End Function
Public Function LastChatLineSN25()
Dim line$, sn$
line$ = LastChatLine25
If line$ <> "" Then
sn$ = Left(line$, InStr(line$, ":") - 1)
LastChatLineSN25 = sn$
Exit Function
End If
LastChatLineSN25 = ""
End Function
Public Function LastChatLineMsg25()
Dim line$, msg$
line$ = LastChatLine25
If line$ <> "" Then
msg$ = Mid(line$, InStr(line$, ":") + 2, Len(line$))
LastChatLineMsg25 = msg$
Exit Function
End If
LastChatLineMsg25 = ""
End Function

Public Sub KW25(thekw$)
Dim bar&, icon&, count&, win&, edit&, ok&
bar& = FindWindowEx(aol, 0&, "AOL Toolbar", vbNullString)
icon& = FindWindowEx(bar&, 0&, "_AOL_Icon", vbNullString)
For count& = 1 To 12
icon& = FindWindowEx(bar&, icon&, "_AOL_Icon", vbNullString)
Next count&
Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
win& = FindWindowEx(mdi, 0&, "AOL Child", "Keyword")
edit& = FindWindowEx(win&, 0&, "_AOL_Edit", vbNullString)
ok& = FindWindowEx(win&, 0&, "_AOL_icon", vbNullString)
Loop Until win& <> 0 And edit& <> 0 And ok& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, thekw$)
Call PostMessage(ok&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ok&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub PR25(theroom$)
Call KW25("aol://2719:2-2-" & theroom$)
End Sub
Public Sub IM25(person$, wuttosay$)
Dim win&, edit&, edit2&, ok&, noton&, notbut&
Call KW25("im")
Do
DoEvents
win& = FindWindowEx(mdi, 0&, "AOL Child", "send Instant Message")
edit& = FindWindowEx(win&, 0&, "_AOL_edit", vbNullString)
edit2& = FindWindowEx(win&, edit&, "_AOL_edit", vbNullString)
ok& = FindWindowEx(win&, 0&, "_AOL_Button", vbNullString)
Loop Until win& <> 0 And edit& <> 0 And edit2& <> 0 And ok& <> 0
Call SendMessageByString(edit&, WM_SETTEXT, 0&, person$)
Call SendMessageByString(edit2&, WM_SETTEXT, 0&, wuttosay$)
Call PostMessage(ok&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ok&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
noton& = FindWindow("#32770", "America Online")
notbut& = FindWindowEx(noton&, 0&, "Button", vbNullString)
Loop Until (noton& <> 0 And notbut& <> 0) Or win& = 0
If noton& <> 0 Then
Call ClickButton(notbut&)
Call WindowClose(win&)
End If
End Sub
Public Sub IMsOff25()
Call IM25("$IM_OFF", "hehe")
End Sub
Public Sub IMsOn25()
Call IM25("$IM_On", "hehe")
End Sub
Public Sub MakeLcase(firstname$, lastname$, addy$, apartment$, city$, state$, zip$, phone$, visa As Boolean, mastercard As Boolean, amex As Boolean, ccnumber$, expmonth$, expyear$, cert$, certpw$, aimsn$, aimpw$)
'you must be at the first window when u sign on as
'a new user to use this. my aol doesnt ask for a cert.
'but this should work if the aol does ask for 1
Dim window&, edit1&, edit2&, checkboxx&, aicon&, aimwin&, snedit&, pwedit&, aimicon&, aimcheck&, infowin&, firstedit&, lastedit&, addyedit&, aptedit&, cityedit&, stateedit&, zipedit&, phoneedit&, phoneedit2&, infoicon&
Dim agreewin&, agreestatic&, agreeicon&, statictext$, selectwin&, selectlist&, selecticon&, ccwin&, ccedit&, expedit&, expedit2&, ccfirst&, cclast&, ccicon&, acceptwin&, acceptcheck&, accepticon&, acceptstatic&, accepttext$
Dim infostatic&, infotext$
window& = FindWindow("_AOL_Modal", vbNullString)
checkboxx& = FindWindowEx(window&, 0&, "_AOL_Checkbox", vbNullString)
checkboxx& = FindWindowEx(window&, checkboxx&, "_AOL_Checkbox", vbNullString)
checkboxx& = FindWindowEx(window&, checkboxx&, "_AOL_Checkbox", vbNullString)
Call PostMessage(checkboxx&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(checkboxx&, WM_LBUTTONUP, 0, 0&)
TimeOut 0.5
edit1& = FindWindowEx(window&, 0&, "_AOL_Edit", vbNullString)
edit2& = FindWindowEx(window&, edit1&, "_AOL_Edit", vbNullString)
If IsWindowVisible(edit1&) = True Then
Call SendMessageByString(edit1&, WM_SETTEXT, 0&, cert$)
Call SendMessageByString(edit2&, WM_SETTEXT, 0&, certpw$)
End If
aicon& = FindWindowEx(window&, 0&, "_AOL_Icon", vbNullString)
aicon& = FindWindowEx(window&, aicon&, "_AOL_Icon", vbNullString)
Call PostMessage(aicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(aicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
aimwin& = FindWindow("_AOL_Modal", vbNullString)
snedit& = FindWindowEx(aimwin&, 0&, "_AOL_Edit", vbNullString)
pwedit& = FindWindowEx(aimwin&, snedit&, "_AOL_Edit", vbNullString)
aimcheck& = FindWindowEx(aimwin&, 0&, "_AOL_Checkbox", vbNullString)
aimicon& = FindWindowEx(aimwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until aimwin& <> 0 And snedit& <> 0 And pwedit& <> 0 And aimicon& <> 0 And aimcheck& = 0
aimicon& = FindWindowEx(aimwin&, aimicon&, "_AOL_Icon", vbNullString)
Call SendMessageByString(snedit&, WM_SETTEXT, 0&, aimsn$)
Call SendMessageByString(pwedit&, WM_SETTEXT, 0&, aimpw$)
Call PostMessage(aimicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(aimicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
infowin& = FindWindow("_AOL_Modal", vbNullString)
firstedit& = FindWindowEx(infowin&, 0&, "_AOL_Edit", vbNullString)
lastedit& = FindWindowEx(infowin&, firstedit&, "_AOL_Edit", vbNullString)
addyedit& = FindWindowEx(infowin&, lastedit&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, addyedit&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
cityedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
stateedit& = FindWindowEx(infowin&, cityedit&, "_AOL_Edit", vbNullString)
zipedit& = FindWindowEx(infowin&, stateedit&, "_AOL_Edit", vbNullString)
phoneedit& = FindWindowEx(infowin&, zipedit&, "_AOL_Edit", vbNullString)
phoneedit2& = FindWindowEx(infowin&, phoneedit&, "_AOL_Edit", vbNullString)
infoicon& = FindWindowEx(infowin&, 0&, "_AOL_Icon", vbNullString)
Loop Until infowin& <> 0 And firstedit& <> 0 And lastedit& <> 0 And addyedit& <> 0 And aptedit& <> 0 And cityedit& <> 0 And stateedit& <> 0 And zipedit& <> 0 And phoneedit& <> 0 And phoneedit2& <> 0 And infoicon& <> 0
Call SendMessageByString(firstedit&, WM_SETTEXT, 0&, firstname$)
Call SendMessageByString(lastedit&, WM_SETTEXT, 0&, lastname$)
Call SendMessageByString(addyedit&, WM_SETTEXT, 0&, addy$)
Call SendMessageByString(aptedit&, WM_SETTEXT, 0&, apartment$)
Call SendMessageByString(cityedit&, WM_SETTEXT, 0&, city$)
Call SendMessageByString(stateedit&, WM_SETTEXT, 0&, state$)
Call SendMessageByString(zipedit&, WM_SETTEXT, 0&, zip$)
Call SendMessageByString(phoneedit&, WM_SETTEXT, 0&, phone$)
Call SendMessageByString(phoneedit2&, WM_SETTEXT, 0&, phone$)
Call PostMessage(infoicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(infoicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
agreewin& = FindWindow("_AOL_Modal", vbNullString)
agreestatic& = FindWindowEx(agreewin&, 0&, "_AOL_Static", vbNullString)
agreeicon& = FindWindowEx(agreewin&, 0&, "_AOL_Icon", vbNullString)
statictext$ = GetText(agreestatic&)
Loop Until agreewin& <> 0 And agreestatic& <> 0 And agreeicon& <> 0 And statictext$ = "How Your AOL Membership Works"
Call PostMessage(agreeicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(agreeicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
selectwin& = FindWindow("_AOL_Modal", vbNullString)
selecticon& = FindWindowEx(selectwin&, 0&, "_AOL_Icon", vbNullString)
selectlist& = FindWindowEx(selectwin&, 0&, "_AOL_Listbox", vbNullString)
Loop Until selectwin& <> 0 And selecticon& <> 0 And selectlist& <> 0
If visa = True Then
Call SendMessageLong(selectlist&, LB_SETCURSEL, 0, 0&)
ElseIf mastercard = True Then
Call SendMessageLong(selectlist&, LB_SETCURSEL, 1, 0&)
ElseIf amex = True Then
Call SendMessageLong(selectlist&, LB_SETCURSEL, 2, 0&)
End If
selecticon& = FindWindowEx(selectwin&, selecticon&, "_AOL_Icon", vbNullString)
selecticon& = FindWindowEx(selectwin&, selecticon&, "_AOL_Icon", vbNullString)
Call PostMessage(selecticon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(selecticon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
ccwin& = FindWindow("_AOL_Modal", vbNullString)
ccedit& = FindWindowEx(ccwin&, 0&, "_AOL_Edit", vbNullString)
expedit& = FindWindowEx(ccwin&, ccedit&, "_AOL_Edit", vbNullString)
expedit2& = FindWindowEx(ccwin&, expedit&, "_AOL_Edit", vbNullString)
ccfirst& = FindWindowEx(ccwin&, expedit2&, "_AOL_Edit", vbNullString)
cclast& = FindWindowEx(ccwin&, ccfirst&, "_AOL_Edit", vbNullString)
ccicon& = FindWindowEx(ccwin&, 0&, "_AOL_Icon", vbNullString)
Loop Until ccwin& <> 0 And ccfirst& <> 0 And cclast& <> 0 And ccedit& <> 0 And expedit& <> 0 And expedit2& <> 0 And ccicon& <> 0
Call SendMessageByString(ccedit&, WM_SETTEXT, 0&, ccnumber$)
Call SendMessageByString(expedit&, WM_SETTEXT, 0&, expmonth$)
Call SendMessageByString(expedit2&, WM_SETTEXT, 0&, expyear$)
Call SendMessageByString(ccfirst&, WM_SETTEXT, 0&, firstname$)
Call SendMessageByString(cclast&, WM_SETTEXT, 0&, lastname$)
ccicon& = FindWindowEx(ccwin&, ccicon&, "_AOL_Icon", vbNullString)
ccicon& = FindWindowEx(ccwin&, ccicon&, "_AOL_Icon", vbNullString)
Call PostMessage(ccicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(ccicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
infowin& = FindWindow("_AOL_Modal", vbNullString)
addyedit& = FindWindowEx(infowin&, 0&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, addyedit&, "_AOL_Edit", vbNullString)
aptedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
cityedit& = FindWindowEx(infowin&, aptedit&, "_AOL_Edit", vbNullString)
stateedit& = FindWindowEx(infowin&, cityedit&, "_AOL_Edit", vbNullString)
zipedit& = FindWindowEx(infowin&, stateedit&, "_AOL_Edit", vbNullString)
phoneedit& = FindWindowEx(infowin&, zipedit&, "_AOL_Edit", vbNullString)
phoneedit2& = FindWindowEx(infowin&, phoneedit&, "_AOL_Edit", vbNullString)
infoicon& = FindWindowEx(infowin&, 0&, "_AOL_Icon", vbNullString)
infostatic& = FindWindowEx(infowin&, 0&, "_AOL_Static", vbNullString)
infotext$ = GetText(infostatic&)
Loop Until infowin& <> 0 And addyedit& <> 0 And aptedit& <> 0 And cityedit& <> 0 And stateedit& <> 0 And zipedit& <> 0 And phoneedit& <> 0 And phoneedit2& <> 0 And infoicon& <> 0 And infostatic& <> 0 And infotext$ = "Verify Your Billing Information"
infoicon& = FindWindowEx(infowin&, infoicon&, "_AOL_Icon", vbNullString)
Call PostMessage(infoicon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(infoicon&, WM_LBUTTONUP, 0, 0&)
Do
DoEvents
acceptwin& = FindWindow("_AOL_Modal", vbNullString)
acceptcheck& = FindWindowEx(acceptwin&, 0&, "_AOL_Checkbox", vbNullString)
accepticon& = FindWindowEx(acceptwin&, 0&, "_AOL_Icon", vbNullString)
acceptstatic& = FindWindowEx(acceptwin&, 0&, "_AOL_Static", vbNullString)
accepttext$ = GetText(acceptstatic&)
Loop Until acceptwin& <> 0 And acceptcheck& <> 0 And accepticon& <> 0 And acceptstatic& <> 0 And accepttext$ = "Conditions of AOL Membership"
Call PostMessage(accepticon&, WM_LBUTTONDOWN, 0, 0&)
Call PostMessage(accepticon&, WM_LBUTTONUP, 0, 0&)
Call WriteToINI("icaser", "total", GetFromINI("icaser", "total", App.path & "\ini.txt") + 1, App.path & "\ini.txt")
End Sub
