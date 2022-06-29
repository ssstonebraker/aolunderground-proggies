Attribute VB_Name = "CueSv140"
'This is a simple bas that i made to try to teach
'all you people who want to learn API.  I mean it
'Took me awile to learn but do it the right way and
'Not the wrong way and copy off of people shit.  I
'DOnt care if you copy this bas file cause it for
'your help.  So just take this and learn from it
'so you can make your own proggies and not steal
'shit from KnK or anyone else.
'Maker:ÇùéS
'Progs made:FurY,CaniBuS
'Mail Me at CueSArt@juno.com
'This Bas is for Aol40 (32bit)

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
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
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
Declare Function gettopwindow Lib "user32" Alias "GetTopWindow" (ByVal hWnd As Long) As Long
Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Declare Function getnextwindow Lib "user32" Alias "GetNextWindow" (ByVal hWnd As Integer, ByVal wFlag As Integer) As Integer
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



'API Declarations for "Kernel32":

Declare Function GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpappname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
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
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
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
Sub Aol40_OpenFile()
AppActivate "AMERICA  ONLINE"
SendKeys "^(o)"
End Sub

Sub Aol40_OpenMail()
AppActivate "AMERICA  ONLINE"
SendKeys "^(r)"
End Sub

Sub Aol40_WriteMail()
AppActivate "AMERICA  ONLINE"
SendKeys "^(m)"
End Sub

Sub Aol40_KeyWorD()
AppActivate "AMERICA  ONLINE"
SendKeys "^(k)"
End Sub

Sub Aol40_SendIM()
AppActivate "AMERICA  ONLINE"
SendKeys "^(i)"
End Sub

Sub Aol40_GetMemberProfile()
AppActivate "AMERICA  ONLINE"
SendKeys "^(g)"
End Sub

Sub Aol40_LocateMember()
AppActivate "AMERICA  ONLINE"
SendKeys "^(l)"
End Sub

Function AOLWindow() As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLWindow = AOL&
End Function

Function AOLMDI() As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(AOL&, "MDIClient")
End Function



Function AOLIsOnline() As Integer
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
welcome& = Findchildbytitle(MDI&, "Welcome, ")
If welcome& = 0 Then
MsgBox "AOL client error: Please sign on to AOL before you resume.", 64
AOLIsOnline = 0
Exit Function
End If
AOLIsOnline = 1
End Function

Function FindChildByClass(Parent As Long, ChildName As String)

Temp& = FindWindowEx(Parent&, 0, ChildName$, vbNullString)
FindChildByClass = Temp&

End Function


Function Findchildbytitle(Parent As Long, child As String) As Long
Ret& = GetWindow(Parent, 5)
Ret& = GetWindow(Ret&, 0)
While Ret&
    DoEvents
    a& = SendMessage(Ret&, &HE, 0&, 0&)
    b$ = String$(a&, 0)
    g& = SendMessageByString(Ret&, &HD, a& + 1, b$)
    If UCase$(b$) Like UCase$(child$) & "*" Then
        Findchildbytitle = Ret&
        Exit Function
    End If
    Ret& = GetWindow(Ret&, 2)
Wend
End Function



Function FindRoom()

ChildHandle& = GetWindow(AOLMDI, 5)

While ChildHandle&
Glyph& = FindChildByClass(ChildHandle&, "_AOL_Glyph")
AOLStatic& = FindChildByClass(ChildHandle&, "_AOL_Static")
Rich& = FindChildByClass(ChildHandle&, "RICHCNTL")
Combo& = FindChildByClass(ChildHandle&, "_AOL_Combobox")
ListBox& = FindChildByClass(ChildHandle&, "_AOL_Listbox")
Icon& = FindChildByClass(ChildHandle&, "_AOL_Icon")
If Glyph& <> 0 And AOLStatic& <> 0 And Rich& <> 0 And Combo& <> 0 And ListBox& <> 0 And Icon& <> 0 Then
FindRoom = ChildHandle&
Exit Function
End If
ChildHandle& = GetWindow(ChildHandle&, 2)
Wend

End Function


Function GetUser()
On Error Resume Next
'This  tells you who is using AoL
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindChildByClass(AOL&, "MDIClient")
welcome& = Findchildbytitle(MDI&, "Welcome, ")
WelcomeLength% = GetWindowTextLength(welcome&)
WelcomeTitle$ = String$(200, 0)
a& = GetWindowText(welcome&, WelcomeTitle$, (WelcomeLength% + 1))
User$ = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
GetUser = User$
End Function


Sub KillWelcome()
Welc& = Findchildbytitle(AOLMDI, "Welcome,")
Ret& = ShowWindow(Welc&, 0)
Ret& = SetFocusApi(AOL&)
End Sub


Sub CueSPause(interval)
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
'EX: CueSPause .4
End Sub


Sub AOLChangeCaption(Change$)
Call SetText(AOLWindow, Change$)
'Ex: Call AOLChangeCaption ("New Caption here")
End Sub


Sub SignOff()
modaa% = FindWindow("#32769", vbNullString)
Wel% = Findchildbytitle(AOLMDI, "Goodbye From America Online")
'EX: Call SignOff
End Sub


Sub MassIM(Lst As ListBox, whattosay$)
For X = 0 To Lst.ListCount - 1
Call IMKeyword(Lst.List(X), Text)
Next X
'For this one you gotta change whattosay to a text
'or a message
' Put Call MassIM In the button for start Mass IM
End Sub

Sub ChatSend(Text)
'this has a pause at the bottom, so u cant
'scroll off with the new tos thingy
If ChatFindRoom = 0 Then Exit Sub
R7% = ChatSendBox
FreeProcess
sBuffer = GetText(R7%)
Call SetText(R7%, "")
Call SetText(R7%, Text)
Do
Call SendCharNum(R7%, 13)
Pause 0.2
Loop Until GetText(ChatSendBox) <> Text
Call SetText(R7%, sBuffer)
Pause 0.6
End Sub

Sub IM_On()

Call IMKeyword("$IM_ON", " ")
'EX: Call IM_On
End Sub

Sub IM_Off()

Call IMKeyword("$IM_OFF", " ")
'EX: Call IM_Off
End Sub


