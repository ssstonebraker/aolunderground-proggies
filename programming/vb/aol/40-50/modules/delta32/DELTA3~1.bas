Attribute VB_Name = "DeLTa32_2"
'Here is build 2 of DeLTa32.  I have improved some bugs
'and sped up the code.... Unfortunately.... no subs
'to run AOL's PopUpMenus or Auto-Phade... but this one is
'better and faster than the first release.
' -ÐèLTá
'If you don't know me... I made the DoT Programs
'                               _________________
'                             -=API Declarations:=-
'                               ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯




'API Declarations for "User32":

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










Sub A_ReadME()
'Release as is...if you don't like the way I code... too
'bad... My Full module is much better.

'I put this first... figuring I would get you to read
'it.  I just want to share some thoughts with you.

'As I mentioned on the first build of this BAS... it is
'basically a downscaled version of the BAS file used in
'the DoT Programs.

'First Thing you will notice is that there are only about
'5 Public Const. As I am called a  "Hex Freak" I
'don't need them because 1. I love Hex and 2. I am too
'lazy to write out something long like WM_GETTEXT

'Next... you will notice a drastic change in my API.
'ALL of the % (Integers) have been changed to & (Longs)
'Welcome to the world of 32-bit where almost everything
'changes to a Long... unlike in 16 bit where it was all
'integers.  The API is now MUCH faster. The reason behind
'this is that since it's supposed to be a Long anyway...
' and you make it an integer, you are forcing VB to do
'some conversions.... so it will just slow down code.

'About AOL's PopUpMenus:  The reason why I don't share
'1 of my 3 subs that can do this is because the first sub
' Took me about 2 months to write.  That's a HELL of a
'time for AOL-Addon purposes.  If you need help... I will
'try to help you..... I wont give you any code.... so I
'hope when you come to me, you know what you are talking
'about

'Bots in AOL4..... Yeah..... they are really hard. Don't
'waste your time subclassing..... it's not like on AOL3
'where it was WM_Settext then you used GetStringFromLPSTR
'to get the text.  What I do is I scan the chat text....
'IM or mail me to find out more on this.

'Auto-Phader -  As seen on DoT Phader 2.0 or higher...
'NO... I can't share that code..... it's not impossible
'to do... it took me about 10 minutes to write a small
'auto-phader.  Since I am the only one with an Auto-Phader
' I will know who copied my idea.

'Fading in General OK... I have included two functions
'One is RGB2Hex.. that will convert RGB to a hex value.
'It pisses me off to see that it is such a long code.
'I have written a MUCH shorter one but for some reason
'When you use that hex it doesn't work... I don't know
'why. The second Function is Fade.  That fades two colors
'With just that one function, you can fade an infinite
'amount of colors.  The solution is simple.  Inquire for
'help.

'Chat-Eater:  You are reading the text of the inventor.
'I invented the chat-eater one lonlely boring day in
'April.... I never publicized it for the reason that it
'would get too lame... once someone else figured it out..
'everyone was doing it.

'Well.... that's about it....  If you need any help or
'anything... IM me..

End Sub

Function AlternatingColors(Text As String, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)
For Alternate = 1 To Len(Text$)
    ThisChr$ = Mid(Text$, Alternate, 1)
    
    If IsOdd(Alternate) = True Then
        If Wavy = True Then
          Let Fixed$ = Fixed$ & "<FONT COLOR=#" & RGB2HEX(Red1, Green1, Blue1) & ">" & "<sup>" & ThisChr$ & "</sup>"
        ElseIf Wavy = False Then
            Let Fixed$ = Fixed$ & "<FONT COLOR=#" & RGB2HEX(Red1, Green1, Blue1) & ">" & ThisChr$
    End If
    
    ElseIf IsOdd(Alternate) = False Then
        If Wavy = True Then
            Let Fixed$ = Fixed$ & "<FONT COLOR=#" & RGB2HEX(Red2, Green2, Blue2) & ">" & "<sub>" & ThisChr$ & "</sub>"
        ElseIf Wavy = False Then
            Let Fixed$ = Fixed$ & "<FONT COLOR=#" & RGB2HEX(Red2, Green2, Blue2) & ">" & ThisChr$
        End If
    End If
Next Alternate

AlternatingColors = Fixed$

End Function


Sub GetCursor()
Call RunMenuByString(AOLWindow(), "&About America Online")
Do: DoEvents
Loop Until FindWindow("_AOL_Modal", vbNullString)
ret& = SendMessage(FindWindow("_AOL_Modal", vbNullString), &H10, 0&, 0&)
End Sub

Sub DoubleClick(hWnd As Integer)
ret& = SendMessageByNum(hWnd, &H203, 0&, 0&)
End Sub

Function IsOdd(Number) As Boolean
Number = Val(Number)
Test$ = Number / 2
If InStr(1, Test$, ".") <> 0 Then
IsOdd = True
Else: IsOdd = False
End If
End Function

Sub KillAnnoyingGlyph()
'Kills that stupid spinning,glowing,blue AOL picture
AoL40& = AOLWindow
TB& = FindChildByClass(AoL40&, "AOL Toolbar")
Toolz& = FindChildByClass(TB&, "_AOL_Toolbar")
Glyph& = FindChildByClass(Toolz&, "_AOL_Glyph")
ret& = SendMessage(Glyph&, &H10, 0, 0&)

End Sub

Function Chat_Hyperlink(Where As String, WhatToSay As String)
Chat_Hyperlink = "<a href=""""><a href=""""><a href=" & Where & ">" & WhatToSay & "<FONT COLOR=""#fffeff"">" & "</a><a href="""">"
End Function
Sub MoveBorderlessForm(TheFrm As Form)
'Place this in Form_MouseDown or if you have a picture
'that is a caption bar put in in Picture1_MouseDown
'The Syntax:  MoveBorderlessForm Me
ReleaseCapture
ret& = SendMessage(TheFrm.hWnd, &HA1, 2, 0)

End Sub

Sub PlayWAV(File As String)
ret& = sndPlaySound(File$, 1)
End Sub

Sub WriteToLog(What As String, FilePath As String)
If FilePath = "" Then Exit Sub
f% = FreeFile
Open FilePath For Binary Access Write As f%
p$ = What & Chr(10)
Put #1, LOF(1) + 1, p$
Close f%
End Sub
Sub CenterForm(Frm As Form)
a% = (Screen.Width - Frm.Width) / 2
b% = (Screen.Height - Frm.Height) / 2
Frm.MOVE a%, b%
End Sub

Function Chat_Black()
Chat_Black = "<FONT COLOR=""#FFFFFF"">"
End Function

Function Chat_Blue()
Chat_Blue = "<FONT COLOR=""#0000FF"">"
End Function


Function Chat_Red()
Chat_Red = "<FONT COLOR=""#FF0000"">"
End Function
Function Chat_Green()
Chat_Green = "<FONT COLOR=""#008000"">"
End Function


Function ChatLag(TheText As String)
g$ = TheText$
a = Len(g$)
For w = 1 To a Step 3
    r$ = Mid$(g$, w, 1)
    u$ = Mid$(g$, w + 1, 1)
    s$ = Mid$(g$, w + 2, 1)
    p$ = p$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html><pre><html><pre><html>" & r$ & "</html></pre></html></pre></html></pre>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & u$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html>" & s$ & "</html></pre>"
Next w
ChatLag = p$
End Function

Function FindChildByClass(Parent As Long, ChildName As String)

Temp& = FindWindowEx(Parent&, 0, ChildName$, vbNullString)
FindChildByClass = Temp&

End Function


Function AOLGetList(LBHandle, Index, Buffer As String)
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
Function ClickList(hWnd As Long, Index As Integer)
ret& = SendMessage(hWnd, &H186, ByVal CLng(Index), ByVal 0&)
End Function
Public Function GetChildCount(ByVal hWnd As Long) As Long
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

GetChildCount = q
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Sub ClickButton(Button As Long)
ret& = SendMessage(Button&, &H100, &H20, 0&)
DoEvents
ret& = SendMessage(Button&, &H100, &H20, 0&)
End Sub





Function AddListToString(List As ListBox)
For DoList = 0 To List.ListCount - 1
AddListToString = AddListToString & List.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)

End Function
Function GetMessageFromIM()
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")

IM& = FindChildByTitle(MDI&, ">Instant Message From:")
If IM& Then GoTo Fletcher
IM& = FindChildByTitle(MDI&, "  Instant Message From:")
If IM& Then GoTo Fletcher
Exit Function
Fletcher:
ImText& = FindChildByClass(IM&, "RICHCNTL")
IMmessage$ = GetText(ImText&)
IMCap$ = GetCaption(IM&)
SN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNLen% = Len(SN$) + 3
Blah$ = Mid(IMmessage$, InStr(IMmessagge$, SN$) + SNLen%)
GetMessageFromIM = Left(Blah$, Len(Blah$) - 1)
CloseWindow (IM&)
End Function


Function GetText(Child As Long)
GetTrim = SendMessage(Child, &HD, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(Child, &HC, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function


Sub ClickIcon(Icon As Long)
ret& = SendMessageByNum(Icon&, &H201, 0&, 0&)
ret& = SendMessageByNum(Icon&, &H201, 0&, 0&)
ret& = SendMessageByNum(Icon&, &H202, 0&, 0&)
ret& = SendMessageByNum(Icon&, &H202, 0&, 0&)

End Sub

Function AOLIsOnline() As Integer
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
Welcome& = FindChildByTitle(MDI&, "Welcome, ")
If Welcome& = 0 Then
MsgBox "AOL client error: Please sign on to AOL before you resume.", 64
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function



Function AOLMDI() As Long
'MDI = Multiple Document Interface
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL&, "MDIClient")
End Function

Function FindRoom()

ChildHandle& = GetWindow(AOLMDI, 5)

While ChildHandle&
Glyph& = FindChildByClass(ChildHandle&, "_AOL_Glyph")
AOLStatic& = FindChildByClass(ChildHandle&, "_AOL_Static")
Rich& = FindChildByClass(ChildHandle&, "RICHCNTL")
combo& = FindChildByClass(ChildHandle&, "_AOL_Combobox")
ListBox& = FindChildByClass(ChildHandle&, "_AOL_Listbox")
Icon& = FindChildByClass(ChildHandle&, "_AOL_Icon")
If Glyph& <> 0 And AOLStatic& <> 0 And Rich& <> 0 And combo& <> 0 And ListBox& <> 0 And Icon& <> 0 Then
FindRoom = ChildHandle&
Exit Function
End If
ChildHandle& = GetWindow(ChildHandle&, 2)
Wend

End Function


Sub SetText(Window As Long, Text$)
'Don't really need this  I prefer SendMessageByString
ret& = SendMessageByString(Window, &HC, 0&, "")
ret& = SendMessageByString(Window, &HC, 0&, Text$)
End Sub



Function AOLWindow() As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL&
End Function

Sub ChatSend40(Text As String)

If FindRoom = 0 Then Exit Sub
AOLListz& = FindChildByClass(FindRoom, "_AOL_Listbox")
If AOLListz& = 0 Then Exit Sub
Room& = FindRoom()
R1& = FindChildByClass(Room&, "RICHCNTL")
For GetThis = 1 To 6
R1& = GetWindow(R1, 2)
Next GetThis

DoEvents

ret& = SendMessageByString(R1&, &HC, 0, Text$)
ret& = SendMessageByNum(R1&, &H102, 13, 0&)

End Sub

Sub HideWelcome()
Welc& = FindChildByTitle(AOLMDI, "Welcome,")
ret& = ShowWindow(Welc&, 0)
ret& = SetFocusAPI(AOL&)
End Sub

Function CountMail()
AoL40& = FindWindow("AOL Frame25", vbNullString)
MDI& = AOLMDI
TB& = FindChildByClass(AoL40&, "AOL Toolbar")
Toolz& = FindChildByClass(TB&, "_AOL_Toolbar")
AoLRead& = FindChildByClass(Toolz&, "_AOL_Icon")
ClickIcon (AoLRead&)
Pause 0.1
u$ = GetUser
Do:
DoEvents
MailPar& = FindChildByTitle(MDI&, u$ + "'s Online Mailbox")
TabControl& = FindChildByClass(MailPar&, "_AOL_TabControl")
TabPage& = FindChildByClass(TabControl&, "_AOL_TabPage")
Tree& = FindChildByClass(TabPage&, "_AOL_Tree")
If MailPar& <> 0 And TabControl& <> 0 And TabPage& <> 0 And Tree& <> 0 Then Exit Do
Loop
Pause 5
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



Function EncryptText(WhereToType As TextBox, Hidden As TextBox)

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





Function FindChildByTitle(Parent As Long, Child As String) As Long
ret& = GetWindow(Parent, 5)
ret& = GetWindow(ret&, 0)
While ret&
    DoEvents
    a& = SendMessage(ret&, &HE, 0&, 0&)
    b$ = String$(a&, 0)
    g& = SendMessageByString(ret&, &HD, a& + 1, b$)
    If UCase$(b$) Like UCase$(Child$) & "*" Then
        FindChildByTitle = ret&
        Exit Function
    End If
    ret& = GetWindow(ret&, 2)
Wend
End Function













Function GetAPIText(hWnd As Long)
Dim sBuffer As String
Dim cLen As Long
cLen = GetWindowTextLength(hWnd)
sBuffer = String(cLen + 1, Chr$(0))
ret& = GetWindowText(hWnd, sBuffer, cLen)
GetAPIText = sBuffer

End Function

Function GetCaption(hWnd As Long)
hwndLength& = GetWindowTextLength(hWnd&)
hWndTitle$ = String$(hwndLength&, 0)
a& = GetWindowText(hWnd&, hWndTitle$, (hwndLength& + 1))
GetCaption = hWndTitle$
End Function

Function GetChat()

ChildS& = FindRoom
Child& = FindChildByClass(ChildS&, "RICHCNTL")
Tex$ = GetText(Child&)
For i = 1 To Len(Tex$)
Returns = InStr(i, Tex$, Chr(13), vbTextCompare)
Next i
GetChat = Tex$

End Function

Function GetClass(Child As Long)
Buffer$ = String$(250, 0)
Getclas& = GetClassName(Child, Buffer$, 250)
GetClass = Buffer$
End Function

Function GetLineCount(Text)
TheView$ = Text

For FindChar = 1 To Len(TheView$)
TheChar$ = Mid(TheView$, FindChar, 1)

If TheChar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Function GetUser()
On Error Resume Next
'This nifty function tell you who is using AoL
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
Welcome& = FindChildByTitle(MDI&, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome&)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(Welcome&, WelcomeTitle$, (WelcomeLength% + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
GetUser = User$
End Function

Function GetWindowsDir()
Buffer$ = String$(255, 0)
u = GetWindowsDirectory(Buffer$, 255)
If Right$(Buffer$, 1) <> "\" Then Buffer$ = Buffer$ + "\"
GetWindowDir = Buffer$
End Function

Sub INI_Write(sAppname As String, sKeyName As String, sNewString As String, sFileName As String)

ret& = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)

End Sub

Function INI_Read(AppName, KeyName As String, Filename As String) As String
Dim sRet As String
    sRet = String(255, Chr(0))
    INI_Read = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), Filename))
End Function

Sub IMRespond(Message$)

IM& = FindChildByTitle(AOLMDI(), ">Instant Message From:") 'Finds the title bar of your new IM
If IM& Then GoTo OhBoy
IM& = FindChildByTitle(AOLMDI(), "  Instant Message From:") 'Finds the title bar  where an existing IM is
If IM& Then GoTo OhBoy
Exit Sub
OhBoy:
e& = FindChildByClass(IM&, "RICHCNTL")
For i = 1 To 9
e& = GetWindow(e&, 2)
Next i
ret& = SendMessageByString(e&, &HC, 0&, Message$)
ClickIcon (FindChildByTitle(IM&, "Send"))
Pause 0.2

End Sub

Function IMsOff()
Call SendIM("$IM_Off", "OFF!")
Pause (0.5)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
IM& = FindChildByTitle(MDI&, "Send Instant Message")
aolcl& = FindWindow("#32770", "America Online")
Closer& = SendMessage(aolcl&, &H10, 0, 0&)
Closer2& = SendMessage(IM&, &H10, 0, 0&)
Loop While aolcl& <> 0 And IM& <> 0
End Function

Function IMsOn()
Call SendIM("$IM_On", "ON!")
Pause (0.5)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
IM& = FindChildByTitle(MDI&, "Send Instant Message")
aolcl& = FindWindow("#32770", "America Online")
Closer& = SendMessage(aolcl&, &H10, 0, 0&)
Closer2& = SendMessage(IM&, &H10, 0, 0&)
Loop While aolcl& <> 0 And IM& <> 0
End Function
Sub SendIM(Person$, Message4)
AppActivate "America  Online"
SendKeys ("^i")
'Why sendkeys?
'Sendkeys are not as BAD as everyone thinks they are
'Besides.... I am not sending a keystroke to AOL's MDI
'so it doesn't matter.
'They work fine and since I didn't include my
'RunPopUpMenu Sub you must resort to sendkeys
'IF you don't like it then put this
'Keyword40 ("aol://9293:")
Do: DoEvents
    MDI& = AOLMDI
    IMWin& = FindChildByTitle(MDI&, "Send Instant Message")
    who& = FindChildByClass(IMWin&, "_AOL_Edit")
    Mess& = FindChildByClass(IMWin&, "RICHCNTL")
    Sender& = GetWindow(Mess&, 2)
    'KWin& = FindChildByTitle(AOLMDI, "Keyword")
    'X = SendMessage(KWin%, WM_CLOSE, 0, 0)
 If IMWin& <> 0 And who& <> 0 And Mess& <> 0 And Sender& <> 0 Then Exit Do
Loop
Pause 0.1
DoEvents
ret& = SendMessageByString(who&, &HC, 0&, Person$)
Pause 0.1
ret& = SendMessageByString(Mess&, &HC, 0&, Message$)
ClickIcon (Sender&)
End Sub


Sub Keyword40(Keyword$)

AOL& = FindWindow("AOL Frame25", vbNullString)
AOTooL& = FindChildByClass(AOL&, "AOL Toolbar")
AOTool2& = FindChildByClass(AOTooL&, "_AOL_Toolbar")
Do:
DoEvents
AOedit& = FindChildByClass(AOTool2&, "_AOL_ComboBox")
AOedit2& = FindChildByClass(AOedit&, "Edit")
If AOTool2& <> 0 And AOedit& <> 0 And AOedit2& <> 0 Then Exit Do
Loop

ret& = SendMessageByString(AOedit2&, &HC, 0&, Keyword$)
DoEvents
ret& = SendMessageByNum(AOedit2&, &H102, &H20, 0)
ret& = SendMessageByNum(AOedit&, &H102, &HD, 0)
End Sub






Function LastChatLine()
'Very Crappy solution to a bot....
GetPar& = FindRoom()
Child& = FindChildByClass(GetPar&, "RICHCNTL")
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
LastChatLine = LastLineo
End Function

Function LineFromText(Text$, TheLine As Integer)
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
TheChatext$ = ReplaceText(TheChatext$, Chr(13), "")
TheChatext$ = ReplaceText(TheChatext$, Chr(10), "")
LineFromText = TheChatext$


End Function

Sub MailSomething(Person$, Subject$, Message$)

AoL40& = AOLWindow
TB& = FindChildByClass(AoL40&, "AOL Toolbar")
Toolz& = FindChildByClass(TB&, "_AOL_Toolbar")
AoLRead& = FindChildByClass(Toolz&, "_AOL_Icon")
AoLWrite& = GetWindow(AoLRead&, 2)
ClickIcon (AoLWrite&)
Pause (0.1)

Do:
DoEvents
MailWin& = FindChildByTitle(AOLMDI(), "Write Mail")
SendTo& = FindChildByClass(MailWin&, "_AOL_Edit")
StaCopyTo& = GetWindow(SendTo&, 2)
CopyTo& = GetWindow(StaCopyTo&, 2)
StaSubj& = GetWindow(CopyTo&, 2)
subj& = GetWindow(StaSubj&, 2)
Msgg& = FindChildByClass(MailWin&, "RICHCNTL")
s1& = FindChildByClass(MailWin&, "_AOL_Icon")
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

Pause 0.1
ClickIcon (SendNow&)

Do:
DoEvents
Modd& = FindWindow("_AOL_Modal", vbNullString)
Icc& = FindChildByClass(Modd&, "_AOL_Icon")
If Icc& <> 0 Then Exit Do
Loop
Pause 0.1
ClickIcon (Icc&)
If Icc& <> 0 Then
ret& = SendMessage(Modd&, &H10, 0&, 0&)
GoTo Here
ElseIf Icc& = 0 Then
Here:
ret& = SendMessage(MailWin&, &H10, 0&, 0&)
End Sub

Sub NotOnTop(the As Form)
ret& = SetWindowPos(the.hWnd, -2, 0&, 0&, 0&, 0&, Flags)
End Sub






Sub ParentChange(Child As Long, NewParent As Long)
ret& = SetParent(Child&, NewParent&)
End Sub

Sub Pause(interval)
'Easiest function on the whole Module
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Function RandomNumber(Finished As Integer)
Randomize
RandomNumber = Int((Val(Finished) * Rnd) + 1)
End Function

Function ReplaceText(Text$, CharFind$, CharChange$)
If InStr(Text$, CharFind$) = 0 Then
ReplaceText = Text$
Exit Function
End If

For ReplaceThis = 1 To Len(Text$)
TheChar$ = Mid(Text$, ReplaceThis, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = CharFind$ Then
TheChars$ = Mid(TheChars$, 1, Len(TheChars$) - 1) + CharChange$
End If
Next ReplaceThis

ReplaceText = TheChars$

End Function

Function ReverseText(Text$)
For Reverse = Len(Text$) To 1 Step -1
    ReverseText = ReverseText & Mid(Text$, Reverse, 1)
Next Reverse
End Function

Function RoomCount()
TheChild& = FindRoom()
List& = FindChildByClass(TheChild&, "_AOL_Listbox") 'Finds the AoL Lisbox in a chat room

GetCount& = SendMessage(List&, &H18B, 0, 0&)
RoomCount = GetCount&
End Function

Sub RunMenu(Menu1 As Long, Menu2 As Long)

'Menu1 is Vertical and starts at 0
'Menu2 is Horizontal and starts at 0
Dim AOLWorks As Long
Static Working As Integer

AOLMenus& = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu& = GetSubMenu(AOLMenus&, Menu1)
AOLItemID& = GetMenuItemID(AOLSubMenu&, Menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu& = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), &H111, AOLItemID, 0&)

End Sub

Sub RunMenuByString(Application, StringSearch)
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



Sub SendCharNum(Win&, Chars$)
e& = SendMessage(Win&, &H102, Chars$, 0&)
End Sub

Function SetChildFocus(Child As Long)
SetThisFocus& = SetFocusAPI(Child&)
End Function

Sub SignOff()
ret& = SendMessage(AOLWindow, WM_CLOSE, 0&, 0&)
End Sub



Sub StayOnTop(TheForm As Form)

ret& = SetWindowPos(TheForm.hWnd, -1, 0&, 0&, 0&, 0&, Flags)
End Sub


Function TrimCharacter(TheText, Chars)
TrimCharacter = ReplaceText(TheText, Chars, "")
End Function

Function TrimReturns(Text$)
TakeChr13 = ReplaceText(Text$, Chr$(13), "")
TakeChr10 = ReplaceText(TakeChr13, Chr$(10), "")
TrimReturns = TakeChr10
End Function

Function TrimSpaces(Text$)
Dim TrimSpace As Integer
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For TrimSpace = 1 To Len(Text$)
TheChar$ = Mid(Text$, TrimSpace, 1)
TheChars$ = TheChars$ & TheChar$

If TheChar$ = " " Then
TheChars$ = Mid(TheChars$, 1, Len(TheChars$) - 1)
End If
Next TrimSpace

TrimSpaces = TheChars$
End Function













Function IfFileExists(ByVal sFileName As String) As Boolean
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        IfFileExists = False
    Else
        IfFileExists = True
    End If

End Function


Sub WaitForOK()
Do:
DoEvents
Modal& = FindWindow("#32770", "America Online")
OKbtn& = FindChildByTitle(Modal&, "OK")
If Modal& <> 0 And OKbtn& <> 0 Then Exit Do
Loop
Pause 0.1
ClickButton (OKbtn&)
ret& = SendMessage(OKbtn&, &H101, &H20, 0)
End Sub

















Function RGB2HEX(r, g, b)
'Hex kix ass!....
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
Configuring$ = TrimSpaces(Configuring$)
RGB2HEX = Configuring$
End Function

Function Fade(YourMessage, Red1, Green1, Blue1, Red2, Green2, Blue2, Wavy As Boolean)

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
            C1 = RGB2HEX(val1, val2, VAL3)
            C2 = RGB2HEX(val1, val2, VAL3)
            C3 = RGB2HEX(val1, val2, VAL3)
            C4 = RGB2HEX(val1, val2, VAL3)
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
Fade = Msg
End Function

Function TrimNull(TheString$)

NewString$ = Trim(TheString$)
Do Until Counter = Len(TheString$)
Counter% = Counter% + 1
Let ThisChr = Asc(Mid$(TheString$, Counter%, 1))
If ThisChr > 31 And ThisChr$ <> 256 Then Word$ = Word$ & Mid$(TheString$, Counter%, 1)
Loop
TrimNull = Word$
End Function

Sub SendBuddyInvitation(Person$, Message$, Room$)

MDI& = AOLMDI
o& = FindChildByTitle(MDI&, "Buddy List Window")
If o& <> 0 Then GoTo Fletcher
Keyword40 ("bv")
Do:
DoEvents
Fletcher:
BuddyBox& = FindChildByTitle(MDI&, "Buddy List Window")
Ico1& = FindChildByClass(BuddyBox&, "_AOL_Icon")
For i = 1 To 6
Ico1& = GetWindow(Ico1&, 2)
Next i
Pause 0.1
ClickIcon (Ico1&)
If BuddyBox& <> 0 And Ico1& <> 0 Then Exit Do
Loop

Do:
DoEvents
BuddyChat& = FindChildByTitle(MDI&, "Buddy Chat")
WhoToInvite& = FindChildByClass(BuddyChat&, "_AOL_Edit")
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
Pause 0.1
SendBut& = FindChildByClass(BuddyChat%, "_AOL_Icon")
ClickIcon (SendBut%)


End Sub

Sub CloseWindow(hWnd As Long)
ret& = SendMessageByNum(hWnd, &H10, 0&, 0&)
End Sub






Sub UpchatOn()

Modal& = FindWindow("_AOL_Modal", vbNullString)
If InStr(1, GetCaption(Modal&), "File Transfer") <> 0 Then
DoEvents
ret& = ShowWindow(Modal&, 6)
ret& = SetFocusAPI(AOLWindow)

End Sub

Sub UpChatOff()
Modal& = FindWindow("_AOL_Modal", vbNullString)
If InStr(1, GetCaption(Modal&), "File Transfer") <> 0 Then
DoEvents
ret& = ShowWindow(Modal&, 1)
ret& = SetFocusAPI(Modal&)

End Sub

Function GetSNFromIM()

AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient") '

IM& = FindChildByTitle(MDI&, ">Instant Message From:")
If IM& Then GoTo Fletcher
IM& = FindChildByTitle(MDI&, "  Instant Message From:")
If IM& Then GoTo Fletcher
Exit Function
Fletcher:
IMCap$ = GetCaption(IM&)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
GetSNFromIM = TheSN$

End Function










Sub KillModals()
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




Sub AddRoom(lst As ListBox)
Room& = FindRoom()
ListX& = FindChildByClass(Room&, "_AOL_Listbox")
Count& = SendMessageByNum(ListX&, &H18B, 0, 0)
Buffer$ = Space(255)
For Counter% = 0 To Count& - 1
List& = AOLGetList(ListX&, Counter%, Buffer$)
For e = 0 To lst.ListCount
    If lst.List(e) = Buffer$ Then
    GoTo Here:
    End If
Next e
If Buffer$ = GetUser Then GoTo Here
lst.AddItem (Buffer$)
Here:
Next Counter%
End Sub







Function SpiralText(sBuffer$)
k$ = sBuffer
DeLTa:
Dim MyLen As Integer
MyString$ = sBuffer$
MyLen = Len(MyString$)
MyStr$ = Mid(MyString$, 2, MyLen) + Mid(MyString$, 1, 1)
sBuffer$ = MyStr$
Pause 1
If sBuffer$ = k$ Then
SpiralText = k$: Exit Function
End If
GoTo DeLTa
End Function





Sub HideAoL()
If AOLIsOnline = 0 Then Exit Sub
DeLTa& = FindWindow("AOL Frame25", vbNullString)
dkfhjl& = ShowWindow(DeLTa, 0)
End Sub

Sub ShowAoL()
Tool& = FindChildByClass(AOLWindow(), "AOL TOOLBAR")
ret& = ShowWindow(AOLWindow(), 5)
ret& = ShowWindow(Tool&, 5)
End Sub

Sub ShowAoLToolBar()
DeLTa& = FindWindow("AOL Frame25", vbNullString)
SocK& = FindChildByClass(DeLTa&, "AOL Toolbar")
PLoP& = ShowWindow(SocK&, 5)
End Sub














Function WhichAoL()
If AOLWindow = 0 Then Exit Function
ToolBar30& = FindChildByClass(AOLWindow, "AOL Toolbar")
ToolBar40& = FindChildByClass(ToolBar30&, "_AOL_Toolbar")

If ToolBar40& <> 0 Then
    WhichAoL = 4
Else
    WhichAoL = 3
End If
End Function











Sub AddStringToList(TheString, TheList As ListBox)

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









Function AddMailToList(Which As Integer, List As ListBox)

AoL40& = FindWindow("AOL Frame25", vbNullString)
MDI& = AOLMDI
TB& = FindChildByClass(AoL40&, "AOL Toolbar")
Toolz& = FindChildByClass(TB&, "_AOL_Toolbar")
AoLRead& = FindChildByClass(Toolz&, "_AOL_Icon")
ClickIcon (AoLRead&)
Pause 0.1
u$ = GetUser
Do:
DoEvents
MailPar& = FindChildByTitle(MDI&, u$ + "'s Online Mailbox")
TabControl& = FindChildByClass(MailPar&, "_AOL_TabControl")
TabPage& = FindChildByClass(TabControl&, "_AOL_TabPage")
If Which = 2 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
If Which = 3 Then TabPage& = GetWindow(TabPage&, GW_HWNDNEXT): TabPage& = GetWindow(TabPage&, GW_HWNDNEXT)
Tree& = FindChildByClass(TabPage&, "_AOL_Tree")
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

Sub KillConnectionLog()
LogWin& = FindChildByClass(AOLMDI, "Connection Log")
ret& = SendMessage(LogWin&, &H10, 0&, 0&)
End Sub





