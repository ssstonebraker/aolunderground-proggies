Attribute VB_Name = "arena32"
'arena32.bas, aol4.o 32 bit
'by lash
'lash@tango5.com
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
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
Declare Function EnumWindows& Lib "user32" (ByVal lpenumfunc As Long, ByVal lParam As Long)
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
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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
Private Declare Function GetWindowThreadProcessID Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function GetNextWindow Lib "user32" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



'API Declarations for "Kernel32":

Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'API Declarations for "dwspy32.dll"
'In there just in case you feel like subclassing
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataByNum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)


'API Declarations for "shell32.dll" and winmm.dll

Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpsSoundName As String, ByVal uFlags As Long) As Long


'Here are the only Constants you actually need
'Otherwise..... just use Hex.

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const MOVE = &HA1

Type POINTAPI
   PointX As Long
   PointY As Long
End Type

Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type












Function alternatingcolors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
For Alternate = 1 To Len(Text$)
    ThisChr$ = Mid(Text$, Alternate, 1)
    
    If isodd(Alternate) = True Then
        If Wavy = True Then
          Let Fixed$ = Fixed$ & "<FONT COLOR=#" & rgb2hex(Red1, Green1, Blue1) & ">" & "<sup>" & ThisChr$ & "</sup>"
        ElseIf Wavy = False Then
            Let Fixed$ = Fixed$ & "<FONT COLOR=#" & rgb2hex(Red1, Green1, Blue1) & ">" & ThisChr$
    End If
    
    ElseIf isodd(Alternate) = False Then
        If Wavy = True Then
            Let Fixed$ = Fixed$ & "<FONT COLOR=#" & rgb2hex(Red2, Green2, Blue2) & ">" & "<sub>" & ThisChr$ & "</sub>"
        ElseIf Wavy = False Then
            Let Fixed$ = Fixed$ & "<FONT COLOR=#" & rgb2hex(Red2, Green2, Blue2) & ">" & ThisChr$
        End If
    End If
Next Alternate

alternatingcolors = Fixed$

End Function


Sub getcursor()
Call runmenubystring(aolwindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
ret& = SendMessage(FindWindow("_AOL_Modal", vbNullString), &H10, 0&, 0&)
End Sub

Sub doubleclick(hWnd As Integer)
ret& = SendMessageByNum(hWnd, &H203, 0&, 0&)
End Sub

Function isodd(Number) As Boolean
Number = Val(Number)
Test$ = Number / 2
If InStr(1, Test$, ".") <> 0 Then
isodd = True
Else: isodd = False
End If
End Function

Sub killannoyingglyph()

AoL40& = aolwindow
TB& = findchildbyclass(AoL40&, "AOL Toolbar")
Toolz& = findchildbyclass(TB&, "_AOL_Toolbar")
Glyph& = findchildbyclass(Toolz&, "_AOL_Glyph")
ret& = SendMessage(Glyph&, &H10, 0, 0&)

End Sub

Function chat_hyperlink(Where As String, WhatToSay As String)
chat_hyperlink = "<a href=""""><a href=""""><a href=" & Where & ">" & WhatToSay & "<FONT COLOR=""#fffeff"">" & "</a><a href="""">"
End Function
Sub moveborderlessform(TheFrm As Form)
ReleaseCapture
ret& = SendMessage(TheFrm.hWnd, &HA1, 2, 0)

End Sub

Sub playwav(File As String)
ret& = sndPlaySound(File$, 1)
End Sub

Sub writetolog(What As String, FilePath As String)
If FilePath = "" Then Exit Sub
f% = FreeFile
Open FilePath For Binary Access Write As f%
p$ = What & Chr(10)
Put #1, LOF(1) + 1, p$
Close f%
End Sub
Sub centerform(Frm As Form)
a% = (Screen.Width - Frm.Width) / 2
b% = (Screen.Height - Frm.Height) / 2
Frm.MOVE a%, b%
End Sub

Function chat_black()
chat_black = "<FONT COLOR=""#FFFFFF"">"
End Function

Function chat_blue()
chat_blue = "<FONT COLOR=""#0000FF"">"
End Function


Function chat_red()
chat_red = "<FONT COLOR=""#FF0000"">"
End Function
Function chat_green()
chat_green = "<FONT COLOR=""#008000"">"
End Function


Function chatlag(TheText As String)
g$ = TheText$
a = Len(g$)
For w = 1 To a Step 3
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html><pre><html><pre><html>" & r$ & "</html></pre></html></pre></html></pre>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html>" & s$ & "</html></pre>"
Next w
chatlag = p$
End Function

Function findchildbyclass(Parent As Long, ChildName As String)

Temp& = FindWindowEx(Parent&, 0, ChildName$, vbNullString)
findchildbyclass = Temp&

End Function


Function aolgetlist(LBHandle, Index, Buffer As String)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim Person As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
AOLThread = GetWindowThreadProcessID(LBHandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or &HF0000, False, AOLProcess)

If AOLProcessThread Then
Person$ = String$(4, vbNullChar)
ListItemHold = SendMessage(LBHandle, &H199, ByVal CLng(Index), ByVal 0&)
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
Function clicklist(hWnd As Long, Index As Integer)
ret& = SendMessage(hWnd, &H186, ByVal CLng(Index), ByVal 0&)
End Function
Public Function getchildcount(ByVal hWnd As Long) As Long
Dim hChild As Long

Dim q As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, 5)
   

While hChild
hChild = GetWindow(hChild, 2)
q = q + 1
Wend

getchildcount = q
   
Exit Function
Return_False:
getchildcount = 0
Exit Function
End Function

Sub clickbutton(Button As Long)
ret& = SendMessage(Button&, &H100, &H20, 0&)
DoEvents
ret& = SendMessage(Button&, &H100, &H20, 0&)
End Sub





Function addlisttostring(List As ListBox)
For DoList = 0 To List.ListCount - 1
addlisttostring = addlisttostring & List.List(DoList) & ", "
Next DoList
addlisttostring = Mid(addlisttostring, 1, Len(addlisttostring) - 2)

End Function
Function getmessagefromim()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = findchildbyclass(AOL&, "MDIClient")

IM& = findchildbytitle(MDI&, ">Instant Message From:")
If IM& Then GoTo Fletcher
IM& = findchildbytitle(MDI&, "  Instant Message From:")
If IM& Then GoTo Fletcher
Exit Function
Fletcher:
ImText& = findchildbyclass(IM&, "RICHCNTL")
IMmessage$ = gettext(ImText&)
IMCap$ = getcaption(IM&)
SN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNLen% = Len(SN$) + 3
Blah$ = Mid(IMmessage$, InStr(IMmessagge$, SN$) + SNLen%)
getmessagefromim = Left(Blah$, Len(Blah$) - 1)
closewindow (IM&)
End Function


Function gettext(Child As Long)
GetTrim = SendMessage(Child, &HD, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(Child, &HC, GetTrim + 1, TrimSpace$)
gettext = TrimSpace$
End Function


Sub clickicon(Icon As Long)
ret& = SendMessageByNum(Icon&, &H201, 0&, 0&)
ret& = SendMessageByNum(Icon&, &H201, 0&, 0&)
ret& = SendMessageByNum(Icon&, &H202, 0&, 0&)
ret& = SendMessageByNum(Icon&, &H202, 0&, 0&)

End Sub

Function aolisonline() As Integer
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = findchildbyclass(AOL&, "MDIClient")
Welcome& = findchildbytitle(MDI&, "Welcome, ")
If Welcome& = 0 Then
MsgBox "AOL client error: Please sign on to AOL before you resume.", 64
aolisonline = 0
Exit Function
End If
aolisonline = 1
End Function



Function aolmdi() As Long

AOL& = FindWindow("AOL Frame25", vbNullString)
aolmdi = findchildbyclass(AOL&, "MDIClient")
End Function

Function findroom()

ChildHandle& = GetWindow(aolmdi, 5)

While ChildHandle&
Glyph& = findchildbyclass(ChildHandle&, "_AOL_Glyph")
AOLStatic& = findchildbyclass(ChildHandle&, "_AOL_Static")
Rich& = findchildbyclass(ChildHandle&, "RICHCNTL")
combo& = findchildbyclass(ChildHandle&, "_AOL_Combobox")
ListBox& = findchildbyclass(ChildHandle&, "_AOL_Listbox")
Icon& = findchildbyclass(ChildHandle&, "_AOL_Icon")
If Glyph& <> 0 And AOLStatic& <> 0 And Rich& <> 0 And combo& <> 0 And ListBox& <> 0 And Icon& <> 0 Then
findroom = ChildHandle&
Exit Function
End If
ChildHandle& = GetWindow(ChildHandle&, 2)
Wend

End Function


Sub settext(Window As Long, Text$)
ret& = SendMessageByString(Window, &HC, 0&, "")
ret& = SendMessageByString(Window, &HC, 0&, Text$)
End Sub



Function aolwindow() As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
aolwindow = AOL&
End Function

Sub chatsend40(Text As String)

If findroom = 0 Then Exit Sub
AOLListz& = findchildbyclass(findroom, "_AOL_Listbox")
If AOLListz& = 0 Then Exit Sub
Room& = findroom()
R1& = findchildbyclass(Room&, "RICHCNTL")
For GetThis = 1 To 6
R1& = GetWindow(R1, 2)
Next GetThis

DoEvents

ret& = SendMessageByString(R1&, &HC, 0, Text$)
ret& = SendMessageByNum(R1&, &H102, 13, 0&)

End Sub

Sub hidewelcome()
Welc& = findchildbytitle(aolmdi, "Welcome,")
ret& = ShowWindow(Welc&, 0)
ret& = SetFocusAPI(AOL&)
End Sub

Function countmail()
AoL40& = FindWindow("AOL Frame25", vbNullString)
MDI& = aolmdi
TB& = findchildbyclass(AoL40&, "AOL Toolbar")
Toolz& = findchildbyclass(TB&, "_AOL_Toolbar")
AoLRead& = findchildbyclass(Toolz&, "_AOL_Icon")
clickicon (AoLRead&)
pause 0.1
u$ = getuser
Do:
DoEvents
MailPar& = findchildbytitle(MDI&, u$ + "'s Online Mailbox")
TabControl& = findchildbyclass(MailPar&, "_AOL_TabControl")
TabPage& = findchildbyclass(TabControl&, "_AOL_TabPage")
Tree& = findchildbyclass(TabPage&, "_AOL_Tree")
If MailPar& <> 0 And TabControl& <> 0 And TabPage& <> 0 And Tree& <> 0 Then Exit Do
Loop
pause 5
sBuffer = SendMessage(Tree&, &H18B, 0, 0&)
If sBuffer > 1 Then
MsgBox "You have " & sBuffer & " messages in your Mailbox.", vbInformation
GoTo Closer
End If
If sBuffer = 1 Then
MsgBox "You have one message in your Mailbox.", vbInformation
GoTo Closer
End If
If sBuffer < 1 Then
MsgBox "You have no messages in your Mailbox.", vbInformation
GoTo Closer
End If
Closer:
ret& = SendMessage(MailPar&, &H10, 0, 0&)
End Function



Function encrypttext(WhereToType As TextBox, Hidden As TextBox)

Dim a%, X%, Y%, z%, i%, Temp, Pharse$
Dim T As String
If WhereToType.Text = "" Then
Exit Function
Else
WhereToType.Enabled = False
Hidden.Text = ""
Hidden.Text = WhereToType.Text
WhereToType.Text = ""
Pharse$ = Hidden.Text
For i = 1 To Len(Pharse$)
      Temp = Asc(Mid$(Pharse$, i, 1))
      Mid$(Pharse$, i, 1) = Chr$(Abs(Temp - 255))
T$ = T$ + Mid$(Pharse$, i, 1)
Next i
T$ = ""
Pharse$ = Hidden.Text
Hidden.Text = ""
Hidden.Text = WhereToType.Text
WhereToType.Text = ""
For i = 1 To Len(Pharse$)
      Temp = Asc(Mid$(Pharse$, i, 1))
      Mid$(Pharse$, i, 1) = Chr$(Abs(Temp - 255))
T$ = T$ + Mid$(Pharse$, i, 1)
Next i
WhereToType.Enabled = True
WhereToType.Text = T$
End If
End Function





Function findchildbytitle(Parent As Long, Child As String) As Long
ret& = GetWindow(Parent, 5)
ret& = GetWindow(ret&, 0)
While ret&
    DoEvents
    a& = SendMessage(ret&, &HE, 0&, 0&)
    b$ = String$(a&, 0)
    g& = SendMessageByString(ret&, &HD, a& + 1, b$)
    If UCase$(b$) Like UCase$(Child$) & "*" Then
        findchildbytitle = ret&
        Exit Function
    End If
    ret& = GetWindow(ret&, 2)
Wend
End Function













Function getapitext(hWnd As Long)
Dim sBuffer As String
Dim cLen As Long
cLen = GetWindowTextLength(hWnd)
sBuffer = String(cLen + 1, Chr$(0))
ret& = GetWindowText(hWnd, sBuffer, cLen)
getapitext = sBuffer

End Function

Function getcaption(hWnd As Long)
hwndLength& = GetWindowTextLength(hWnd&)
hWndTitle$ = String$(hwndLength&, 0)
a& = GetWindowText(hWnd&, hWndTitle$, (hwndLength& + 1))
getcaption = hWndTitle$
End Function

Function getchat()

ChildS& = findroom
Child& = findchildbyclass(ChildS&, "RICHCNTL")
Tex$ = gettext(Child&)
For i = 1 To Len(Tex$)
Returns = InStr(i, Tex$, Chr(13), vbTextCompare)
Next i
getchat = Tex$

End Function

Function getclass(Child As Long)
Buffer$ = String$(250, 0)
Getclas& = GetClassName(Child, Buffer$, 250)
getclass = Buffer$
End Function

Function getlinecount(Text)
TheView$ = Text

For FindChar = 1 To Len(TheView$)
TheChar$ = Mid(TheView$, FindChar, 1)

If TheChar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
getlinecount = numline
Else
getlinecount = numline + 1
End If
End Function

Function getuser()
On Error Resume Next
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = findchildbyclass(AOL&, "MDIClient")
Welcome& = findchildbytitle(MDI&, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome&)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(Welcome&, WelcomeTitle$, (WelcomeLength% + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
getuser = User$
End Function

Function getwindowsdir()
Buffer$ = String$(255, 0)
u = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function

Sub ini_write(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)

ret& = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

Function ini_read(AppName, KeyName As String, Filename As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    ini_read = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), Filename))
End Function

Sub imrespond(Message$)

IM& = findchildbytitle(aolmdi(), ">Instant Message From:") 'Finds the title bar of your new IM
If IM& Then GoTo OhBoy
IM& = findchildbytitle(aolmdi(), "  Instant Message From:") 'Finds the title bar  where an existing IM is
If IM& Then GoTo OhBoy
Exit Sub
OhBoy:
e& = findchildbyclass(IM&, "RICHCNTL")
For i = 1 To 9
e& = GetWindow(e&, 2)
Next i
ret& = SendMessageByString(e&, &HC, 0&, Message$)
clickicon (findchildbytitle(IM&, "Send"))
pause 0.2

End Sub

Function imsoff()
Call sendim("$IM_Off", "OFF!")
pause (0.5)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = findchildbyclass(AOL&, "MDIClient")
IM& = findchildbytitle(MDI&, "Send Instant Message")
aolcl& = FindWindow("#32770", "America Online")
Closer& = SendMessage(aolcl&, &H10, 0, 0&)
Closer2& = SendMessage(IM&, &H10, 0, 0&)
Loop While aolcl& <> 0 And IM& <> 0
End Function

Function imson()
Call sendim("$IM_On", "ON!")
pause (0.5)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = findchildbyclass(AOL&, "MDIClient")
IM& = findchildbytitle(MDI&, "Send Instant Message")
aolcl& = FindWindow("#32770", "America Online")
Closer& = SendMessage(aolcl&, &H10, 0, 0&)
Closer2& = SendMessage(IM&, &H10, 0, 0&)
Loop While aolcl& <> 0 And IM& <> 0
End Function
Sub sendim(Person$, Message4)
AppActivate "America  Online"
SendKeys ("^i")
Do: DoEvents
    MDI& = aolmdi
    IMWin& = findchildbytitle(MDI&, "Send Instant Message")
    who& = findchildbyclass(IMWin&, "_AOL_Edit")
    Mess& = findchildbyclass(IMWin&, "RICHCNTL")
    Sender& = GetWindow(Mess&, 2)
 If IMWin& <> 0 And who& <> 0 And Mess& <> 0 And Sender& <> 0 Then Exit Do
Loop
pause 0.1
DoEvents
ret& = SendMessageByString(who&, &HC, 0&, Person$)
pause 0.1
ret& = SendMessageByString(Mess&, &HC, 0&, Message$)
clickicon (Sender&)
End Sub


Sub keyword40(Keyword$)

AOL& = FindWindow("AOL Frame25", vbNullString)
AOTooL& = findchildbyclass(AOL&, "AOL Toolbar")
AOTool2& = findchildbyclass(AOTooL&, "_AOL_Toolbar")
Do:
DoEvents
AOedit& = findchildbyclass(AOTool2&, "_AOL_ComboBox")
AOedit2& = findchildbyclass(AOedit&, "Edit")
If AOTool2& <> 0 And AOedit& <> 0 And AOedit2& <> 0 Then Exit Do
Loop

ret& = SendMessageByString(AOedit2&, &HC, 0&, Keyword$)
DoEvents
ret& = SendMessageByNum(AOedit2&, &H102, &H20, 0)
ret& = SendMessageByNum(AOedit&, &H102, &HD, 0)
End Sub






Function lastchatline()
GetPar& = findroom()
Child& = findchildbyclass(GetPar&, "RICHCNTL")
GetTrim& = SendMessageByNum(Child&, &HD, 0&, 0&)
TrimSpace$ = Space$(GetTrim&)
GetString& = SendMessageByString(Child&, &HC, GetTrim& + 1, TrimSpace$)

TheView$ = TrimSpace$


For FindChar = 1 To Len(TheView$)
TheChar$ = Mid(TheView$, FindChar, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = Chr(13) Then
TheChatext$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
TheChars$ = ""
End If

Next FindChar

LastLen = Val(FindChar) - Len(TheChars$)
LastLineo$ = Mid(TheView$, LastLen + 1, Len(TheChars$))
lastchatline = LastLineo
End Function

Function linefromtext(Text$, TheLine As Integer)
TheView$ = Text$
For FindChar = 1 To Len(TheView$)
TheChar$ = Mid(TheView$, FindChar, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = Chr(13) Then
c = c + 1
TheChatext$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
If TheLine = c Then GoTo Saturn
TheChars$ = ""
End If

Next FindChar
Exit Function
Saturn:
TheChatext$ = replacetext(TheChatext$, Chr(13), "")
TheChatext$ = replacetext(TheChatext$, Chr(10), "")
linefromtext = TheChatext$


End Function

Sub mailsomething(Person$, Subject$, Message$)

AoL40& = aolwindow
TB& = findchildbyclass(AoL40&, "AOL Toolbar")
Toolz& = findchildbyclass(TB&, "_AOL_Toolbar")
AoLRead& = findchildbyclass(Toolz&, "_AOL_Icon")
AoLWrite& = GetWindow(AoLRead&, 2)
clickicon (AoLWrite&)
pause (0.1)

Do:
DoEvents
MailWin& = findchildbytitle(aolmdi(), "Write Mail")
SendTo& = findchildbyclass(MailWin&, "_AOL_Edit")
StaCopyTo& = GetWindow(SendTo&, 2)
CopyTo& = GetWindow(StaCopyTo&, 2)
StaSubj& = GetWindow(CopyTo&, 2)
subj& = GetWindow(StaSubj&, 2)
Msgg& = findchildbyclass(MailWin&, "RICHCNTL")
s1& = findchildbyclass(MailWin&, "_AOL_Icon")
For i = 1 To 17
s17& = GetWindow(s1&, 2)
Next i
s18& = GetWindow(s17&, 2)
SendNow& = GetWindow(s18&, 2)
If SendTo& <> 0 And CopyTo& <> 0 And subj& <> 0 And Msgg& <> 0 And SendNow& <> 0 Then Exit Do
Loop

ret& = SendMessageByString(SendTo&, &HC, 0&, Person$)
ret& = SendMessageByString(subj&, &HC, 0&, Subject$)
ret& = SendMessageByString(Msgg&, &HC, 0&, Message$)

pause 0.1
clickicon (SendNow&)

Do:
DoEvents
Modd& = FindWindow("_AOL_Modal", vbNullString)
Icc& = findchildbyclass(Modd&, "_AOL_Icon")
If Icc& <> 0 Then Exit Do
Loop
pause 0.1
clickicon (Icc&)
If Icc& <> 0 Then
ret& = SendMessage(Modd&, &H10, 0&, 0&)
GoTo Here
ElseIf Icc& = 0 Then
Here:
ret& = SendMessage(MailWin&, &H10, 0&, 0&)
End Sub

Sub notontop(the As Form)
ret& = SetWindowPos(the.hWnd, -2, 0&, 0&, 0&, 0&, Flags)
End Sub






Sub parentchange(Child As Long, NewParent As Long)
ret& = SetParent(Child&, NewParent&)
End Sub

Sub pause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Function freeprocess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function randomnumber(Finished As Integer)
Randomize
randomnumber = Int((Val(Finished) * Rnd) + 1)
End Function

Function replacetext(Text$, CharFind$, CharChange$)
If InStr(Text$, CharFind$) = 0 Then
replacetext = Text$
Exit Function
End If

For ReplaceThis = 1 To Len(Text$)
TheChar$ = Mid(Text$, ReplaceThis, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = CharFind$ Then
TheChars$ = Mid(TheChars$, 1, Len(TheChars$) - 1) + CharChange$
End If
Next ReplaceThis

replacetext = TheChars$

End Function

Function reversetext(Text$)
For Reverse = Len(Text$) To 1 Step -1
    reversetext = reversetext & Mid(Text$, Reverse, 1)
Next Reverse
End Function

Function roomcount()
TheChild& = findroom()
List& = findchildbyclass(TheChild&, "_AOL_Listbox") 'Finds the AoL Lisbox in a chat room

GetCount& = SendMessage(List&, &H18B, 0, 0&)
roomcount = GetCount&
End Function

Sub runmenu(Menu1 As Long, Menu2 As Long)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus& = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu& = GetSubMenu(AOLMenus&, Menu1)
AOLItemID& = GetMenuItemID(AOLSubMenu&, Menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu& = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), &H111, AOLItemID, 0&)

End Sub

Sub runmenubystring(Application, StringSearch)
ToSearch& = GetMenu(Application)
MenuCount& = GetMenuItemCount(ToSearch&)

For FindString = 0 To MenuCount& - 1
ToSearchSub& = GetSubMenu(ToSearch&, FindString)
MenuItemCount& = GetMenuItemCount(ToSearchSub&)

For GetString = 0 To MenuItemCount& - 1
SubCount& = GetMenuItemID(ToSearchSub&, GetString)
MenuString$ = String$(100, " ")
GetStringMenu& = GetMenuString(ToSearchSub&, SubCount&, MenuString$, 100, 1)

If InStr(UCase(MenuString$), UCase(StringSearch)) Then
MenuItem& = SubCount&
GoTo MatchString
End If

Next GetString

Next FindString
MatchString:
RunTheMenu& = SendMessage(Application, &H111, MenuItem&, 0)
End Sub



Sub sendcharnum(Win&, Chars$)
e& = SendMessage(Win&, &H102, Chars$, 0&)
End Sub

Function setchildfocus(Child As Long)
SetThisFocus& = SetFocusAPI(Child&)
End Function

Sub signoff()
ret& = SendMessage(aolwindow, WM_CLOSE, 0&, 0&)
End Sub



Sub stayontop(TheForm As Form)

ret& = SetWindowPos(TheForm.hWnd, -1, 0&, 0&, 0&, 0&, Flags)
End Sub


Function trimcharacter(TheText, Chars)
trimcharacter = replacetext(TheText, Chars, "")
End Function

Function trimreturns(Text$)
TakeChr13 = replacetext(Text$, Chr$(13), "")
TakeChr10 = replacetext(TakeChr13, Chr$(10), "")
trimreturns = TakeChr10
End Function

Function trimspaces(Text$)
Dim TrimSpace As Integer
If InStr(Text, " ") = 0 Then
trimspaces = Text
Exit Function
End If

For TrimSpace = 1 To Len(Text$)
TheChar$ = Mid(Text$, TrimSpace, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = " " Then
TheChars$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
End If
Next TrimSpace

trimspaces = TheChars$
End Function













Function iffileexists(ByVal sFileName As String) As Boolean
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        iffileexists = False
    Else
        iffileexists = True
    End If

End Function


Sub waitforok()
Do:
DoEvents
Modal& = FindWindow("#32770", "America Online")
OKbtn& = findchildbytitle(Modal&, "OK")
If Modal& <> 0 And OKbtn& <> 0 Then Exit Do
Loop
pause 0.1
clickbutton (OKbtn&)
ret& = SendMessage(OKbtn&, &H101, &H20, 0)
End Sub

















Function rgb2hex(r, g, b)
Dim X&
Dim XX&
Dim Color&
Dim Divide
Dim Answer&
Dim Remainder&
Dim Configuring$
For X& = 1 To 3
    If X& = 1 Then Color& = b
    If X& = 2 Then Color& = g
    If X& = 3 Then Color& = r
    For XX& = 1 To 2
        Divide = Color& / 16
        Answer& = Int(Divide)
        Remainder& = (10000 * (Divide - Answer&)) / 625
        If Remainder& < 10 Then Configuring$ = Str(Remainder&) + Configuring$
        If Remainder& = 10 Then Configuring$ = "A" + Configuring$
        If Remainder& = 11 Then Configuring$ = "B" + Configuring$
        If Remainder& = 12 Then Configuring$ = "C" + Configuring$
        If Remainder& = 13 Then Configuring$ = "D" + Configuring$
        If Remainder& = 14 Then Configuring$ = "E" + Configuring$
        If Remainder& = 15 Then Configuring$ = "F" + Configuring$
        Color& = Answer&
    Next XX&
Next X&
Configuring$ = trimspaces(Configuring$)
rgb2hex = Configuring$
End Function

Function fade(YourMessage, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)

   C1BAK = C1
   C2BAK = C2
   C3BAK = C3
   C4BAK = C4
   c = 0
   o = 0
   o2 = 0
   q = 1
   Q2 = 1
   For X = 1 To Len(YourMessage)
            BVAL1 = Red2 - Red1
            BVAL2 = Green2 - Green1
            BVAL3 = Blue2 - Blue1
            val1 = (BVAL1 / Len(YourMessage) * X) + Red1
            val2 = (BVAL2 / Len(YourMessage) * X) + Green1
            VAL3 = (BVAL3 / Len(YourMessage) * X) + Blue1
            C1 = rgb2hex(val1, val2, VAL3)
            C2 = rgb2hex(val1, val2, VAL3)
            C3 = rgb2hex(val1, val2, VAL3)
            C4 = rgb2hex(val1, val2, VAL3)
            If C1 = C2 And C2 = C3 And C3 = C4 And C4 = C1 Then c = 1: Msg = Msg & "<FONT COLOR=#" + C1 + ">"
            If o2 = 1 Then o2 = 2 Else If o2 = 2 Then o2 = 3 Else If o2 = 3 Then o2 = 4 Else o2 = 1
            If c <> 1 Then
            If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
            If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
            If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
            If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
            End If
            If Wavy = True Then
                If o2 = 1 Then Msg = Msg + "<sub>"
                If o2 = 3 Then Msg = Msg + "<sup>"
                Msg = Msg + Mid$(YourMessage, X, 1)
                If o2 = 1 Then Msg = Msg + "</sub>"
                If o2 = 3 Then Msg = Msg + "</sup>"
                If Q2 = 2 Then
                    q = 1
                    Q2 = 1

                    If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
                    If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
                    If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
                    If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
                End If
            ElseIf Wavy = False Then
                Msg = Msg + Mid$(YourMessage, X, 1)
                If Q2 = 2 Then
                    q = 1
                    Q2 = 1

                    If o2 = 1 Then Msg = Msg + "<FONT COLOR=#" + C1 + ">"
                    If o2 = 2 Then Msg = Msg + "<FONT COLOR=#" + C2 + ">"
                    If o2 = 3 Then Msg = Msg + "<FONT COLOR=#" + C3 + ">"
                    If o2 = 4 Then Msg = Msg + "<FONT COLOR=#" + C4 + ">"
                End If
        
            End If
nc: Next X
C1 = C1BAK
C2 = C2BAK
C3 = C3BAK
C4 = C4BAK
fade = Msg
End Function

Function trimnull(TheString$)

NewString$ = Trim(TheString$)
Do Until Counter = Len(TheString$)
Counter% = Counter% + 1
Let ThisChr = Asc(Mid$(TheString$, Counter%, 1))
If ThisChr > 31 And ThisChr$ <> 256 Then Word$ = Word$ & Mid$(TheString$, Counter%, 1)
Loop
trimnull = Word$
End Function

Sub sendbuddyinvitation(Person$, Message$, Room$)

MDI& = aolmdi
o& = findchildbytitle(MDI&, "Buddy List Window")
If o& <> 0 Then GoTo Fletcher
keyword40 ("bv")
Do:
DoEvents
Fletcher:
BuddyBox& = findchildbytitle(MDI&, "Buddy List Window")
Ico1& = findchildbyclass(BuddyBox&, "_AOL_Icon")
For i = 1 To 6
Ico1& = GetWindow(Ico1&, 2)
Next i
pause 0.1
clickicon (Ico1&)
If BuddyBox& <> 0 And Ico1& <> 0 Then Exit Do
Loop

Do:
DoEvents
BuddyChat& = findchildbytitle(MDI&, "Buddy Chat")
WhoToInvite& = findchildbyclass(BuddyChat&, "_AOL_Edit")
MsgToSend& = GetWindow(WhoToInvite&, 2)
w1& = GetWindow(MsgToSend&, 2)
w2& = GetWindow(w1&, 2)
w3& = GetWindow(w2&, 2)
w4& = GetWindow(w3&, 2)
Roomer% = GetWindow(w4&, 2)
If BuddyChat% <> 0 And WhoToInvite& <> 0 And MsgToSend& <> 0 And Roomer& <> 0 Then Exit Do
Loop

ret& = SendMessageByString(WhoToInvite&, &HC, 0&, Person$)
ret& = SendMessageByString(MsgToSend&, &HC, 0&)
ret& = SendMessageByString(Roomer&, &HC, 0&, Room$)
pause 0.1
SendBut& = findchildbyclass(BuddyChat%, "_AOL_Icon")
clickicon (SendBut%)


End Sub

Sub closewindow(hWnd As Long)
ret& = SendMessageByNum(hWnd, &H10, 0&, 0&)
End Sub






Sub upchaton()

Modal& = FindWindow("_AOL_Modal", vbNullString)
If InStr(1, getcaption(Modal&), "File Transfer") <> 0 Then
DoEvents
ret& = ShowWindow(Modal&, 6)
ret& = SetFocusAPI(aolwindow)

End Sub

Sub upchatoff()
Modal& = FindWindow("_AOL_Modal", vbNullString)
If InStr(1, getcaption(Modal&), "File Transfer") <> 0 Then
DoEvents
ret& = ShowWindow(Modal&, 1)
ret& = SetFocusAPI(Modal&)

End Sub

Function getsnfromim()

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = findchildbyclass(AOL&, "MDIClient") '

IM& = findchildbytitle(MDI&, ">Instant Message From:")
If IM& Then GoTo Fletcher
IM& = findchildbytitle(MDI&, "  Instant Message From:")
If IM& Then GoTo Fletcher
Exit Function
Fletcher:
IMCap$ = getcaption(IM&)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
getsnfromim = TheSN$

End Function










Sub killmodals()
Dim the As Integer
Modal& = FindWindow("_AOL_Modal", vbNullString)
the = -1
Do:
If Modal& = 0 Then Exit Do
Modal& = FindWindow("_AOL_Modal", vbNullString)
ret& = SendMessage(Modal&, &H10, 0&, 0&)
the = the + 1
Loop
If the < 1 Then
MsgBox "You have no Modal Windows open", vbInformation
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




Sub addroom(lst As ListBox)
Room& = findroom()
ListX& = findchildbyclass(Room&, "_AOL_Listbox")
Count& = SendMessageByNum(ListX&, &H18B, 0, 0)
Buffer$ = Space(255)
For Counter% = 0 To Count& - 1
List& = aolgetlist(ListX&, Counter%, Buffer$)
For e = 0 To lst.ListCount
    If lst.List(e) = Buffer$ Then
    GoTo Here:
    End If
Next e
If Buffer$ = getuser Then GoTo Here
lst.AddItem (Buffer$)
Here:
Next Counter%
End Sub







Function spiraltext(sBuffer$)
k$ = sBuffer
DeLTa:
Dim MyLen As Integer
MyString$ = sBuffer$
MyLen = Len(MyString$)
MyStr$ = Mid(MyString$, 2, MyLen) + Mid(MyString$, 1, 1)
sBuffer$ = MyStr$
pause 1
If sBuffer$ = k$ Then
spiraltext = k$: Exit Function
End If
GoTo DeLTa
End Function





Sub hideaol()
If aolisonline = 0 Then Exit Sub
DeLTa& = FindWindow("AOL Frame25", vbNullString)
dkfhjl& = ShowWindow(DeLTa, 0)
End Sub

Sub showaol()
Tool& = findchildbyclass(aolwindow(), "AOL TOOLBAR")
ret& = ShowWindow(aolwindow(), 5)
ret& = ShowWindow(Tool&, 5)
End Sub

Sub showaoltoolbar()
DeLTa& = FindWindow("AOL Frame25", vbNullString)
SocK& = findchildbyclass(DeLTa&, "AOL Toolbar")
PLoP& = ShowWindow(SocK&, 5)
End Sub














Function whichaol()
If aolwindow = 0 Then Exit Function
ToolBar30& = findchildbyclass(aolwindow, "AOL Toolbar")
ToolBar40& = findchildbyclass(ToolBar30&, "_AOL_Toolbar")

If ToolBar40& <> 0 Then
    whichaol = 4
Else
    whichaol = 3
End If
End Function











Sub addstringtolist(TheString, TheList As ListBox)

If Not Mid(TheString, Len(TheString), 1) = "," Then
    TheString = TheString & ","
End If

For DoList = 1 To Len(TheString)
    TheChars$ = TheChars$ & Mid(TheString, DoList, 1)

    If Mid(TheString, DoList, 1) = "," Then
        TheList.AddItem Mid(TheChars$, 1, Len(TheChars$) - 1)
        TheChars$ = ""
        If Mid(TheString, DoList + 1, 1) = " " Then
            DoList = DoList + 1
        End If
    End If
Next DoList

End Sub









Function addmailtolist(Which As Integer, List As ListBox)

AoL40& = FindWindow("AOL Frame25", vbNullString)
MDI& = aolmdi
TB& = findchildbyclass(AoL40&, "AOL Toolbar")
Toolz& = findchildbyclass(TB&, "_AOL_Toolbar")
AoLRead& = findchildbyclass(Toolz&, "_AOL_Icon")
clickicon (AoLRead&)
pause 0.1
u$ = getuser
Do:
DoEvents
MailPar& = findchildbytitle(MDI&, u$ + "'s Online Mailbox")
TabControl& = findchildbyclass(MailPar&, "_AOL_TabControl")
TabPage& = findchildbyclass(TabControl&, "_AOL_TabPage")
If Which = 2 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
If Which = 3 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT): TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
Tree& = findchildbyclass(TabPage&, "_AOL_Tree")
If MailPar& <> 0 And TabControl& <> 0 And TabPage& <> 0 And Tree& <> 0 Then Exit Do
Loop

sBuffer& = SendMessage(Tree&, &H18B, 0, 0)

For MailNum = 0 To sBuffer
TxtLen& = SendMessageByNum(Tree&, &H18A, MailNum, 0&)
txt$ = String(TxtLen& + 1, 0&)
GTTXT& = SendMessageByString(Tree&, &H189, MailNum, txt$)
NewMail = RTrim(txt$)
List.AddItem (NewMail)
Next MailNum

End Function

Sub killconnectionlog()
LogWin& = findchildbyclass(aolmdi, "Connection Log")
ret& = SendMessage(LogWin&, &H10, 0&, 0&)
End Sub





