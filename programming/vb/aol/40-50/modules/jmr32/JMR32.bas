Attribute VB_Name = "JMR32"
' So here it is, JMR32.bas ver. 1.  I included an
' explanation for every sub or function.  Not
' everything was tested on this, so please notify
' me if something does not work.
' Email:
' lkytdy@hotmail.com
' /././././././././././././././././././././././././././././././././././././
' "A real programmer doesn't rely on a bas file, but
' instead uses the code to further greater his/her
' knowledge in VB." - JMR
' \.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\.\
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
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
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Type POINTAPI
X As Long
Y As Long
End Type
Public AOL&
Sub Chat_Send(YourText As String)
' This sends the specified text into a chatroom
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDIC& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLC& = FindWindowEx(MDIC&, 0, "AOL Child", vbNullString)
chatbox& = FindWindowEx(AOLC&, 0, "RICHCNTL", vbNullString)
FnHandle& = FindWindowEx(AOLC&, chatbox&, "RICHCNTL", vbNullString)
If MDIC& And FnHandle& <> 0 Then Exit Do
Loop
SendMessageByString FnHandle&, WM_SETTEXT, 0&, YourText
SendMessageLong FnHandle&, WM_CHAR, ENTER_KEY, 0&
End Sub
Function Chat_FindRoom()
' Simple FindChatRoom function
AOL& = FindWindow("AOL Frame25", "America  Online")
TheMDIC& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
DaShiet& = FindWindowEx(TheMDIC&, 0, "AOL Child", vbNullString)
FindChatRoom = DaShiet&
End Function
Sub Chat_Close()
' This does what it says, it closes the chat room
SendMessage FindChatRoom, WM_CLOSE, 0&, 0&
End Sub
Sub Chat_Private(ChatRoom As String)
' Opens a specified private room
Keyword "aol://2719:2-2-" & ChatRoom
End Sub
Sub Chat_Buddy(People As String, DaRoom As String)
' Sends a buddy chat to the people you list, and
' with the room you request
Keyword "BuddyView"
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDIC& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLC& = FindWindowEx(MDIC&, 0, "AOL Child", vbNullString)
Ickon1& = FindWindowEx(AOLC&, 0, "_AOL_Icon", vbNullString)
Ickon2& = FindWindowEx(AOLC&, Ickon1&, "_AOL_Icon", vbNullString)
Ickon3& = FindWindowEx(AOLC&, Ickon2&, "_AOL_Icon", vbNullString)
BCBtn& = FindWindowEx(AOLC&, Ickon3&, "_AOL_Icon", vbNullString)
If MDIC& And BCBtn& <> 0 Then Exit Do
Loop
SendMessage BCBtn&, WM_LBUTTONDOWN, 0&, 0&
SendMessage BCBtn&, WM_LBUTTONUP, 0&, 0&
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDICing& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLCing& = FindWindowEx(MDICing&, 0, "AOL Child", vbNullString)
AOEdit& = FindWindowEx(AOLCing&, 0, "_AOL_Edit", vbNullString)
If MDICing& And AOEdit& <> 0 Then Exit Do
Loop
SendMessageByString AOEdit&, WM_SETTEXT, 0&, People
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDICer& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOChild& = FindWindowEx(MDICer&, 0, "AOL Child", vbNullString)
AOEditer1& = FindWindowEx(AOChild&, 0, "_AOL_Edit", vbNullString)
AOEditer2& = FindWindowEx(AOChild&, AOEditer1&, "_AOL_Edit", vbNullString)
AOEditFin& = FindWindowEx(AOChild&, AOEditer2&, "_AOL_Edit", vbNullString)
If MDICer& And AOEditFin& <> 0 Then Exit Do
Loop
SendMessageByString AOEditFin&, WM_SETTEXT, 0&, DaRoom
SendMessageLong AOEditFin&, WM_CHAR, ENTER_KEY, 0&
End Sub
Sub AOLPrint()
' This sub clicks the AOL Printer button so it opens the
' printer dialog.
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
Toolbar1& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar1&, 0, "_AOL_Toolbar", vbNullString)
Ikon1& = FindWindowEx(toolbar2&, 0, "_AOL_Icon", vbNullString)
Ikon2& = FindWindowEx(toolbar2&, Ikon1&, "_AOL_Icon", vbNullString)
Ikon3& = FindWindowEx(toolbar2&, Ikon2&, "_AOL_Icon", vbNullString)
DaIkon& = FindWindowEx(toolbar2&, Ikon3&, "_AOL_Icon", vbNullString)
If toolbar2& And DaIkon& <> 0 Then Exit Do
Loop
SendMessage DaIkon&, WM_LBUTTONDOWN, 0&, 0&
SendMessage DaIkon&, WM_LBUTTONUP, 0&, 0&
End Sub
Sub Keyword(DaKeyword As String)
' Opens up a specified keyword
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
Toolbar1& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
toolbar2& = FindWindowEx(Toolbar1&, 0, "_AOL_Toolbar", vbNullString)
Combo& = FindWindowEx(toolbar2&, 0, "_AOL_Combobox", vbNullString)
TheHandle& = FindWindowEx(Combo&, 0, "Edit", vbNullString)
If toolbar2& And TheHandle& <> 0 Then Exit Do
Loop
SendMessageByString TheHandle&, WM_SETTEXT, 0&, DaKeyword
SendMessageLong TheHandle&, WM_CHAR, ENTER_KEY, 0&
End Sub
Sub IM_Send(Personn As String, Message As String)
' This simply sends an IM
Keyword "IM"
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDIC& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLC& = FindWindowEx(MDIC&, 0, "AOL Child", vbNullString)
IMHando1& = FindWindowEx(AOLC&, 0, "_AOL_Edit", vbNullString)
If MDIC& And IMHando1& <> 0 Then Exit Do
Loop
SendMessageByString IMHando1&, WM_SETTEXT, 0&, Personn
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDIC_Again& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLC_Again& = FindWindowEx(MDIC_Again&, 0, "AOL Child", vbNullString)
IMHando2& = FindWindowEx(AOLC_Again&, 0, "RICHCNTL", vbNullString)
If MDIC_Again& And IMHando2& <> 0 Then Exit Do
Loop
SendMessageByString IMHando2&, WM_SETTEXT, 0&, Message
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDICer& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLcer& = FindWindowEx(MDICer&, 0, "AOL Child", vbNullString)
Iko1& = FindWindowEx(AOLcer&, 0, "_AOL_Icon", vbNullString)
Iko2& = FindWindowEx(AOLcer&, Iko1&, "_AOL_Icon", vbNullString)
Iko3& = FindWindowEx(AOLcer&, Iko2&, "_AOL_Icon", vbNullString)
Iko4& = FindWindowEx(AOLcer&, Iko3&, "_AOL_Icon", vbNullString)
Iko5& = FindWindowEx(AOLcer&, Iko4&, "_AOL_Icon", vbNullString)
Iko6& = FindWindowEx(AOLcer&, Iko5&, "_AOL_Icon", vbNullString)
Iko7& = FindWindowEx(AOLcer&, Iko6&, "_AOL_Icon", vbNullString)
Iko8& = FindWindowEx(AOLcer&, Iko7&, "_AOL_Icon", vbNullString)
DaButton& = FindWindowEx(AOLcer&, Iko8&, "_AOL_Icon", vbNullString)
If MDICer& And DaButton& <> 0 Then Exit Do
Loop
SendMessage DaButton&, WM_LBUTTONDOWN, 0&, 0&
SendMessage DaButton&, WM_LBUTTONUP, 0&, 0&
End Sub
Sub IM_Close()
' Does what it says, it closes an IM window
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
ShitImBored& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
IMWindo& = FindWindowEx(ShitImBored&, 0, "AOL Child", vbNullString)
If IMWindo& <> 0 Then Exit Do
Loop
SendMessage IMWindo&, WM_CLOSE, 0&, 0&
End Sub
Sub IM_Off()
' Turns your IMs Off
IM_Send "$IM_Off", " "
End Sub
Sub IM_On()
' Turns your IMs On
IM_Send "$IM_On", " "
End Sub
Sub Kill45Minute()
' This kills that anoying 45 minute timer
' You might want to put this in a timer of some sort
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
Palett& = FindWindowEx(AOL&, 0, "_AOL_Palette", vbNullString)
IkonHandle& = FindWindowEx(Palett&, 0, "_AOL_Icon", vbNullString)
If IkonHandle& <> 0 Then Exit Do
Loop
End Sub
Public Sub Window_OnTop(Windo)
' Sets the window on top
' Example: Window_OnTop Form1.hwnd
SetWindowPos Windo, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS
End Sub
Public Sub Window_NotOnTop(Windo)
' Makes it so the window doesn't stay on top
' Example: Window_NotOnTop Form1.hwnd
SetWindowPos Windo, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS
End Sub
Public Sub Pause(HowLong As Long)
' Pauses between code
Counter = Timer
Do Until Timer - Counter >= HowLong
DoEvents
Loop
End Sub
Public Sub Play_Wav(wav As String)
' Plays a sound
sndPlaySound wav, SND_FLAG
End Sub
Function Child_FindByTitle(Windo, ChildYouWant As String) As Integer
' Allows you to find a child window no matter what.
' All you need is the windo hwnd, and the child's
' caption, for the window you want to find.
TheChild% = GetWindow(Windo, 5)
While TheChild%
WindoLength% = GetWindowTextLength(TheChild%)
DaB$ = String$(WindoLength%, 0)
WindowText% = GetWindowText(TheChild%, DaB$, WindoLength% + 1)
If InStr(UCase(DaB$), UCase(ChildYouWant)) Then Child_FindByTitle = TheChild%: Exit Function
TheChild% = GetWindow(TheChild%, 2)
Wend
End Function
Function User_SN() As String
' Gets the users Screen Name
On Error GoTo ErEx
Dim DaString$
Dim TL&
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
AOMDI& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
DaWindo& = FindWindowEx(AOMDI&, 0, "AOL Child", vbNullString)
WWindo = Child_FindByTitle(AOMDI&, "Welcome, ")
If DaWindo& <> 0 Then Exit Do
Loop
TL& = GetWindowTextLength(WWindo)
DaString$ = String(TL&, 0&)
GetWindowText WWindo, DaString$, TL& + 1
TheSN = Left(DaString, Len(DaString) - 1)
TheSN = Right(TheSN, Len(DaString$) - 9)
User_SN = TheSN
ErEx:
End Function
Sub Chat_RollDice()
' This rolls dice in a chatroom
Chat_Send "//Roll"
End Sub
Sub VBPrint(WhatToPrint As String, FonSize As Integer, FonName As String, Boldd As Boolean, Italicc As Boolean, Underlinee As Boolean, StrikeThruu As Boolean)
' This prints within VB
Printer.FontSize = FonSize
Printer.FontName = FonName
Printer.FontBold = Boldd
Printer.FontItalic = Italicc
Printer.FontStrikethru = StrikeThruu
Printer.FontUnderline = Underlinee
Printer = WhatToPrint
Printer.Print
End Sub
Sub Mail_Send(Ppl As String, TheSbjct As String, DaMsg As String)
' Allows you to send an Email
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
AOLTool& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
AOTool& = FindWindowEx(AOLTool&, 0, "_AOL_Toolbar", vbNullString)
AOIko& = FindWindowEx(AOTool&, 0, "_AOL_Icon", vbNullString)
DaBtn& = FindWindowEx(AOTool&, AOIko&, "_AOL_Icon", vbNullString)
If AOLTool& And DaBtn& <> 0 Then Exit Do
Loop
SendMessage DaBtn&, WM_LBUTTONDOWN, 0&, 0&
SendMessage DaBtn&, WM_LBUTTONUP, 0&, 0&
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDICn& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
TheAOC& = FindWindowEx(MDICn&, 0, "AOL Child", vbNullString)
AOEditer& = FindWindowEx(TheAOC&, 0, "_AOL_Edit", vbNullString)
If MDICn& And AOEditer& <> 0 Then Exit Do
Loop
SendMessageByString AOEditer&, WM_SETTEXT, 0&, Ppl
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDICr& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLCn& = FindWindowEx(MDICr&, 0, "AOL Child", vbNullString)
AOE1& = FindWindowEx(AOLCn&, 0, "_AOL_Edit", vbNullString)
AOE2& = FindWindowEx(AOLCn&, AOE1&, "_AOL_Edit", vbNullString)
AOLEditn& = FindWindowEx(AOLCn&, AOE2&, "_AOL_Edit", vbNullString)
If MDICr& And AOLEditn& <> 0 Then Exit Do
Loop
SendMessageByString AOLEditn&, WM_SETTEXT, 0&, TheSbjct
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDIClient& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0, "AOL Child", vbNullString)
RichText& = FindWindowEx(AOLChild&, 0, "RICHCNTL", vbNullString)
If MDIClient& And RichText& <> 0 Then Exit Do
Loop
SendMessageByString RichText&, WM_SETTEXT, 0&, DaMsg
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
MDICFinal& = FindWindowEx(AOL&, 0, "MDIClient", vbNullString)
AOLCFinal& = FindWindowEx(MDICFinal&, 0, "AOL Child", vbNullString)
AIkon1& = FindWindowEx(AOLCFinal&, 0, "_AOL_Icon", vbNullString)
AIkon2& = FindWindowEx(AOLCFinal&, AIkon1&, "_AOL_Icon", vbNullString)
AIkon3& = FindWindowEx(AOLCFinal&, AIkon2&, "_AOL_Icon", vbNullString)
AIkon4& = FindWindowEx(AOLCFinal&, AIkon3&, "_AOL_Icon", vbNullString)
AIkon5& = FindWindowEx(AOLCFinal&, AIkon4&, "_AOL_Icon", vbNullString)
AIkon6& = FindWindowEx(AOLCFinal&, AIkon5&, "_AOL_Icon", vbNullString)
AIkon7& = FindWindowEx(AOLCFinal&, AIkon6&, "_AOL_Icon", vbNullString)
AIkon8& = FindWindowEx(AOLCFinal&, AIkon7&, "_AOL_Icon", vbNullString)
AIkon9& = FindWindowEx(AOLCFinal&, AIkon8&, "_AOL_Icon", vbNullString)
AIkon10& = FindWindowEx(AOLCFinal&, AIkon9&, "_AOL_Icon", vbNullString)
AIkon11& = FindWindowEx(AOLCFinal&, AIkon10&, "_AOL_Icon", vbNullString)
AIkon12& = FindWindowEx(AOLCFinal&, AIkon11&, "_AOL_Icon", vbNullString)
AIkon13& = FindWindowEx(AOLCFinal&, AIkon12&, "_AOL_Icon", vbNullString)
BtnSend& = FindWindowEx(AOLCFinal&, AIkon13&, "_AOL_Icon", vbNullString)
If MDICFinal& And BtnSend& <> 0 Then Exit Do
Loop
SendMessage BtnSend&, WM_LBUTTONDOWN, 0&, 0&
SendMessage BtnSend&, WM_LBUTTONUP, 0&, 0&
End Sub
Sub Mail_ReadNew()
' Opens the "New mail" mail box
Do: DoEvents
AOL& = FindWindow("AOL Frame25", "America  Online")
ATool& = FindWindowEx(AOL&, 0, "AOL Toolbar", vbNullString)
AToolAgain& = FindWindowEx(ATool&, 0, "_AOL_Toolbar", vbNullString)
BtnRead& = FindWindowEx(AToolAgain&, 0, "_AOL_Icon", vbNullString)
If ATool& And BtnRead& <> 0 Then Exit Do
Loop
SendMessage BtnRead&, WM_LBUTTONDOWN, 0&, 0&
SendMessage BtnRead&, WM_LBUTTONUP, 0&, 0&
End Sub
