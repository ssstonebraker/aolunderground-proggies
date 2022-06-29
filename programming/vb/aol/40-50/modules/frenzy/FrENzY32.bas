Attribute VB_Name = "FrENzY32"
'FrENzY BAS By izekial83, 1998
'Sup, this is the best easyiest, and biggest BAS
'Out There. PLZ Dont Just Copy My Shyt.
'PLZ -=> Mess Around With The Subs And Try To Learn It
'Ok This is For AOL4 ONLY And VB4 Im Sure It WorX
'In VB5, and VB6 But I Made It Using VB4
'Email: Funkdemon@yahoo.com
'SN: izekial83 (AIM)
Public Const conMCIAppTitle = "MCI Control Application"
Public Const conMCIErrInvalidDeviceID = 30257
Public Const conMCIErrDeviceOpen = 30263
Public Const conMCIErrCannotLoadDriver = 30266
Public Const conMCIErrUnsupportedFunction = 30274
Public Const conMCIErrInvalidFile = 30304

#If Win32 Then
    Declare Function GetFocus Lib "user32" () As Long
#Else
    Declare Function GetFocus Lib "User" () As Integer
#End If
Public Const FADE_RED = &HFF&
Public Const FADE_GREEN = &HFF00&
Public Const FADE_BLUE = &HFF0000
Public Const FADE_YELLOW = &HFFFF&
Public Const FADE_WHITE = &HFFFFFF
Public Const FADE_BLACK = &H0&
Public Const FADE_PURPLE = &HFF00FF
Public Const FADE_GREY = &HC0C0C0
Public Const FADE_PINK = &HFF80FF
Public Const FADE_TURQUOISE = &HC0C000

Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Public DialogCaption As String
Global ScrollStop As Boolean
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long
Type tSavedInfo
    sSentFrom(100) As String
    sMess(100) As String
End Type
Global SavedInfo As tSavedInfo
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Global File
Global Appname
Global KeyName
Global value
Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Declare Function fCreateShellLink Lib "STKIT432.DLL" _
(ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, _
ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Global IMAnswerMsgNum As Boolean
Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Dim ArrayNum As Integer     ' Index value for the menu control array mnuFileArray.
Public FileName As String   ' This variable keeps track of the filename information for opening and closing files.
Global AvailibleYes As Boolean
Global CATWatchStop As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function DiskSpaceFree Lib "SETUPKIT.DLL" () As Long
Global GuideStop As Boolean
Global GreetStop As Boolean
Global MMBotStop As Boolean
Global ScramblerStop As Boolean
Global ScramblerTimeStop As Boolean
Global ScrambleTime As Long
Global ScrambledAnswer As String
Global ScramblerAnswer As String
Global PM1 As String
Global PM2 As String
Global PM3 As String
Global PM4 As String
Global PM5 As String
Global PM6 As String
Global PM7 As String
Global PM8 As String
Global PM9 As String
Global PM0 As String
Global sTxt As String
Global IMAnswerMsg As String
Global VotedYes As Integer
Global VotedNo As Integer
Global i As Integer
Global XMassIM As Integer
Global O As Integer
Global CustChatString As String
Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long

Global bUseSendChatsub As Boolean

Global Const GFSR_SYSTEMRESOURCES = 0
Global Const GFSR_GDIRESOURCES = 1
Global Const GFSR_USERRESOURCES = 2

Global Const WM_MDICREATE = &H220
Global Const WM_MDIDESTROY = &H221
Global Const WM_MDIACTIVATE = &H222
Global Const WM_MDIRESTORE = &H223
Global Const WM_MDINEXT = &H224
Global Const WM_MDIMAXIMIZE = &H225
Global Const WM_MDITILE = &H226
Global Const WM_MDICASCADE = &H227
Global Const WM_MDIICONARRANGE = &H228
Global Const WM_MDIGETACTIVE = &H229
Global Const WM_MDISETMENU = &H230


Global Const WM_CUT = &H300
Global Const WM_COPY = &H301
Global Const WM_PASTE = &H302
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
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

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
   x As Long
   Y As Long
End Type

Function AOLWindow()
aol% = FindWindow("AOL Frame25", vbNullString)
End Function

Sub CaptionScrolls()
'+~-> This Is Just An AOLCaption Thing I Made I Puts Thats Text In The AOL Window Like Typing
'+~-> Dont Use My Scrolls
aol% = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online -")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online --")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - -")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ -")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Ê-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êt-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñ-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål -")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ß-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý -")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý î-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îz-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzè-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèk-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkî-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîå-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål8-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83-")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 ->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 -->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 --->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 ---->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 ----->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 ------>")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 ------->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 -------->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 --------->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 ---------->")
TimeOut 0.01
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online - @Ø£ Êtè®ñål ßý îzèkîål83 ----------->")

End Sub
Sub ChangeAOLCaption(Wha As TextBox)
'+~-> This Changes The AOL Window Caption
'+~-> Call ChangeAOLCaption ((Text1))
aol% = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(aol%, WM_SETTEXT, 0, "America  Online -" & " " & Wha)
TimeOut 0.01
End Sub
Sub ChangeSNInWelcomeWindow(Wha As TextBox)
'+~-> Changes The SN In The Welcome Winddow
'+~-> Call ChangeSNInWelcomeWindow((Text1))
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Welcome% = FindChildByTitle(mdi%, "Welcome, ")
Call SendMessageByString(Welcome%, WM_SETTEXT, 0, "Welcome," & " " & Wha)
End Sub


Sub FadeColors()
'FADE_RED
'FADE_GREEN
'FADE_BLUE
'FADE_YELLOW
'FADE_WHITE
'FADE_BLACK
'FADE_PURPLE
'FADE_GREY
'FADE_PINK
'FADE_TURQUOISE
'Or Make Your Own By RGB Values
End Sub
Sub CopyFile(FileName$, CopyTo$)
'+~-> This Will Copy One File To Somwhere Else
'+~-> Example Call CopyFile("C:\autoexec.bat", "C:\Temp")
If FileName$ = "" Then Exit Sub
If CopyTo$ = "" Then Exit Sub
If Not DoesFileExist(FileName$) Then Exit Sub
On Error GoTo AnErrOccured
If InStr(Right$(FileName$, 4), ".") = 0 Then Exit Sub
If InStr(Right$(CopyTo$, 4), ".") = 0 Then Exit Sub
FileCopy FileName$, CopyTo$
Exit Sub
AnErrOccured:
MsgBox "An Unexpected Error Occured!", 16, "Error"
End Sub
Sub ClickIconDouble(Button%)
'+~-> This Will Double Cllick The Button Of Your Choice
'+~-> Not The Same As Putting ClickIcon Twice
'+~-> Call ClickIconDouble(IM%)
Dim DoubleClickNow%
DoubleClickNow% = SendMessageByNum(Button%, WM_LBUTTONDBLCLK, &HD, 0)
End Sub
Sub GetFonts(thelist As Control)
'+~-> This Will Load All The Fonts Into A ListBoX
'+~-> Ex. Call GetFonts(List1)
   thelist.Clear
   For BuildIt = 1 To Screen.FontCount
     thelist.AddItem Screen.Fonts(BuildIt - 1)
   Next
End Sub
Sub HowToMakeScrollBars()
'OK Per Label (Which is used for the Color To Fade), You Need 3 ScrollBars
'We'll Do 3 Labels (9Bars) Used In Fade3Colors
'Make 3 Labels And Name Them Color1, Color2, And Color3
'Make 9 ScrollBars
'Set The Property "Max" to 255
'Then Put Them Next To Each Other With A Space Every 3 Bars
'The First 3 Bars Name Red1, Green1, Blue1
'Then Double Click On Red1
'GoTo The The Drop Menu Next To Proc:
'Go Down To Scroll
'And Put This In There
'Color1.BackColor = rgB(Red1.value, Green1.value, Blue1.value)
'Do The Same To Blue1 And Green1
'Now Do The Same To Red2, Green2, Blue2(Which Are The Next 3 Scroll Bars)
'Then Put That Code In The Scroll Statement Except Change It To
'Color2.BackColor = rgB(Red2.value, Green2.value, Blue2.value)
'And Repeat The Same With Red3, Green3, Blue3
'With This Lesson You Can Make A 10 Color Scroller Also
End Sub

Sub RunWithDefaultProgram(txtFile As TextBox)
'+~-> This Will Open Something With Default Program
'+~-> Call RunWithDefaultProgram("E:\Downloads\Images\TupacAtMGM.jpg")
'+~-> Thats An Example
If Dir(txtFile) = "" Then
    Call MsgBox("The file in the text box does not exist.", vbExclamation)
    Exit Sub
End If


Call ShellExecute(hwnd, "Open", txtFile, "", App.Path, 1)
End Sub
Sub MailIzekial83(message)
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, "Funkdemon@yahoo.com")

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, "Im Using Your BAS File")
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(mdi%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub

Sub Scroller(What)
'+~-> This Is A Plain Scroller
'+~-> Call Scroller(Text1) or Call Scroller("FrENzY32 BAS Is The BEST!")
'+~-> To Stop Put ScrollStop = True
Do
Call SendChat(What)
TimeOut 0.5
Loop Until ScrollStop = True
End Sub

Sub ServerFind(Who, What)
'+~-> This Will Find Something In A Server
'+~-> Call ServerFind ("izekial83", "Adobe")
Call SendChat("/" & Who & " Find " & What)
TimeOut 1
End Sub
Sub ServerSend(Who, What)
'+~-> This Will Send Something In A Server
'+~-> Call ServerSend ("izekial83", "69")
Call SendChat("/" & Who & " Send " & What)
TimeOut 1
End Sub
Sub ServerStatus(Who)
'+~-> This Will Send The Server Status
'+~-> Call ServerStatus("izekial83")
Call SendChat("/" & Who & " Send " & "Status")
TimeOut 1
End Sub
Public Sub ShutDownWindows()
'+~-> ShutsDwon Windows
'+~-> Call ShutDownWindows
Dim MsgRes As Long
MsgRes = MsgBox("Do you really want to Shut Down Windows 95/98?", vbYesNo Or vbQuestion)
If MsgRes = vbNo Then Exit Sub
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub
Sub SpiralScroll(Wha As TextBox)
'+~-> This Will Scroll In A Spiral
'+~-> Call SpiralScroll(WhatTextBoXHere)
SendChat (Wha)
TimeOut (0.75)
Dim MYLEN As Integer
MYSTRING = Wha
MYLEN = Len(MYSTRING)
MYSTR = Mid(MYSTRING, 2, MYLEN) + Mid(MYSTRING, 1, 1)
Wha = MYSTR
End Sub

Function Text_backwards(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let Newsent$ = NextChr$ & Newsent$
Loop
Text_backwards = Newsent$
End Function
Function Text_Big(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then Leet$ = "A"
    If letter$ = "b" Then Leet$ = "B"
    If letter$ = "c" Then Leet$ = "C"
    If letter$ = "d" Then Leet$ = "D"
    If letter$ = "e" Then Leet$ = "E"
    If letter$ = "f" Then Leet$ = "F"
    If letter$ = "g" Then Leet$ = "G"
    If letter$ = "h" Then Leet$ = "H"
    If letter$ = "i" Then Leet$ = "I"
    If letter$ = "j" Then Leet$ = "J"
    If letter$ = "k" Then Leet$ = "K"
    If letter$ = "l" Then Leet$ = "L"
    If letter$ = "m" Then Leet$ = "M"
    If letter$ = "n" Then Leet$ = "N"
    If letter$ = "o" Then Leet$ = "O"
    If letter$ = "p" Then Leet$ = "P"
    If letter$ = "q" Then Leet$ = "Q"
    If letter$ = "r" Then Leet$ = "R"
    If letter$ = "s" Then Leet$ = "S"
    If letter$ = "t" Then Leet$ = "T"
    If letter$ = "u" Then Leet$ = "U"
    If letter$ = "v" Then Leet$ = "V"
    If letter$ = "w" Then Leet$ = "W"
    If letter$ = "x" Then Leet$ = "X"
    If letter$ = "y" Then Leet$ = "Y"
    If letter$ = "z" Then Leet$ = "Z"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
Text_Big = Made$
End Function
Function Text_Elite(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then Leet$ = "â"
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "e" Then Leet$ = "ë"
    If letter$ = "i" Then Leet$ = "î"
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then Leet$ = "õ"
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then Leet$ = "ü"
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then Leet$ = "Ã"
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then Leet$ = "Í"
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
Text_Elite = Made$
End Function

Function Text_Hacker(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    If letter$ = "a" Then Leet$ = "a"
    If letter$ = "b" Then Leet$ = "B"
    If letter$ = "c" Then Leet$ = "C"
    If letter$ = "d" Then Leet$ = "D"
    If letter$ = "e" Then Leet$ = "e"
    If letter$ = "f" Then Leet$ = "F"
    If letter$ = "g" Then Leet$ = "G"
    If letter$ = "h" Then Leet$ = "H"
    If letter$ = "i" Then Leet$ = "i"
    If letter$ = "j" Then Leet$ = "J"
    If letter$ = "k" Then Leet$ = "K"
    If letter$ = "l" Then Leet$ = "L"
    If letter$ = "m" Then Leet$ = "M"
    If letter$ = "n" Then Leet$ = "N"
    If letter$ = "o" Then Leet$ = "o"
    If letter$ = "p" Then Leet$ = "P"
    If letter$ = "q" Then Leet$ = "Q"
    If letter$ = "r" Then Leet$ = "R"
    If letter$ = "s" Then Leet$ = "S"
    If letter$ = "t" Then Leet$ = "T"
    If letter$ = "u" Then Leet$ = "u"
    If letter$ = "v" Then Leet$ = "V"
    If letter$ = "w" Then Leet$ = "W"
    If letter$ = "x" Then Leet$ = "X"
    If letter$ = "y" Then Leet$ = "y"
    If letter$ = "z" Then Leet$ = "Z"
    If letter$ = "A" Then Leet$ = "a"
    If letter$ = "B" Then Leet$ = "B"
    If letter$ = "C" Then Leet$ = "C"
    If letter$ = "D" Then Leet$ = "D"
    If letter$ = "E" Then Leet$ = "e"
    If letter$ = "F" Then Leet$ = "F"
    If letter$ = "G" Then Leet$ = "G"
    If letter$ = "H" Then Leet$ = "H"
    If letter$ = "I" Then Leet$ = "i"
    If letter$ = "J" Then Leet$ = "J"
    If letter$ = "K" Then Leet$ = "K"
    If letter$ = "L" Then Leet$ = "L"
    If letter$ = "M" Then Leet$ = "M"
    If letter$ = "N" Then Leet$ = "N"
    If letter$ = "O" Then Leet$ = "o"
    If letter$ = "P" Then Leet$ = "P"
    If letter$ = "Q" Then Leet$ = "Q"
    If letter$ = "R" Then Leet$ = "R"
    If letter$ = "S" Then Leet$ = "S"
    If letter$ = "T" Then Leet$ = "T"
    If letter$ = "U" Then Leet$ = "u"
    If letter$ = "V" Then Leet$ = "V"
    If letter$ = "W" Then Leet$ = "W"
    If letter$ = "X" Then Leet$ = "X"
    If letter$ = "Y" Then Leet$ = "y"
    If letter$ = "Z" Then Leet$ = "Z"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
Text_Hacker = Made$
End Function


Function Text_Small(word$)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "A" Then Leet$ = "a"
    If letter$ = "B" Then Leet$ = "b"
    If letter$ = "C" Then Leet$ = "c"
    If letter$ = "D" Then Leet$ = "d"
    If letter$ = "E" Then Leet$ = "e"
    If letter$ = "F" Then Leet$ = "f"
    If letter$ = "G" Then Leet$ = "g"
    If letter$ = "H" Then Leet$ = "h"
    If letter$ = "I" Then Leet$ = "i"
    If letter$ = "J" Then Leet$ = "j"
    If letter$ = "K" Then Leet$ = "k"
    If letter$ = "L" Then Leet$ = "l"
    If letter$ = "M" Then Leet$ = "m"
    If letter$ = "N" Then Leet$ = "n"
    If letter$ = "O" Then Leet$ = "o"
    If letter$ = "P" Then Leet$ = "p"
    If letter$ = "Q" Then Leet$ = "q"
    If letter$ = "R" Then Leet$ = "r"
    If letter$ = "S" Then Leet$ = "s"
    If letter$ = "T" Then Leet$ = "t"
    If letter$ = "U" Then Leet$ = "u"
    If letter$ = "V" Then Leet$ = "v"
    If letter$ = "W" Then Leet$ = "w"
    If letter$ = "X" Then Leet$ = "x"
    If letter$ = "Y" Then Leet$ = "y"
    If letter$ = "Z" Then Leet$ = "z"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
Text_Small = Made$
End Function
Function Text_Spaced(strin As String)
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + " "
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Spaced = Newsent$
End Function


Sub StartGhosting()
'+~-> This Will Make You Start Ghosting
'+~-> Ex. Call StartGhosting
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
buddy% = FindChildByTitle(mdi%, "Buddy List Window")
If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If
SetupButton% = FindChildByClass(buddy%, "_AOL_Icon")
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
ClickIcon (SetupButton%)
TimeOut 3
PPButton% = FindChildByClass(buddysetup%, "_AOL_Icon")
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
ClickIcon (PPButton%)
TimeOut 3
PPWin% = FindChildByTitle(mdi%, "Privacy Preferences")
BlockAll% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
ClickIcon (BlockAll%)
BlockAll2% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
ClickIcon (BlockAll2%)
TimeOut 0.5
SaveButton% = FindChildByClass(PPWin%, "_AOL_Icon")
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
ClickIcon (SaveButton%)
waitforok
Call SendMessage(buddysetup%, WM_CLOSE, 0, 0)
End Sub
Sub StopGhosting()
'+~-> This Will Stop You From Ghosting
'+~-> Ex. Call Stop Ghosting
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
buddy% = FindChildByTitle(mdi%, "Buddy List Window")
If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If
SetupButton% = FindChildByClass(buddy%, "_AOL_Icon")
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
SetupButton% = GetWindow(SetupButton%, GW_HWNDNEXT)
ClickIcon (SetupButton%)
TimeOut 3
buddysetup% = FindChildByTitle(mdi%, UserSN & "'s Buddy Lists")
PPButton% = FindChildByClass(buddysetup%, "_AOL_Icon")
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
PPButton% = GetWindow(PPButton%, GW_HWNDNEXT)
ClickIcon (PPButton%)
TimeOut 3
PPWin% = FindChildByTitle(mdi%, "Privacy Preferences")
BlockAll% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
BlockAll% = GetWindow(BlockAll%, GW_HWNDNEXT)
ClickIcon (BlockAll%)
BlockAll2% = FindChildByClass(PPWin%, "_AOL_Checkbox")
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
BlockAll2% = GetWindow(BlockAll2%, GW_HWNDNEXT)
ClickIcon (BlockAll2%)
TimeOut 0.5
SaveButton% = FindChildByClass(PPWin%, "_AOL_Icon")
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
SaveButton% = GetWindow(SaveButton%, GW_HWNDNEXT)
ClickIcon (SaveButton%)
waitforok
Call SendMessage(buddysetup%, WM_CLOSE, 0, 0)
End Sub
Sub GreetBot()
'+~-> This Is A Greet Bot It Says Welcome to Anyone Who Comes In The Room
'+~-> Ex. Call GreetBot - To Stop Put GreetStop = true
GreetStop = False
sstart:
If GreetStop = True Then Exit Sub
Dim sTheSN As String
Dim iDiff As Integer
Dim sChatLine As String
If SNFromLastChatLine() = "OnlineHost" Then
    sChatLine = LastChatLine()
    If Right$(sChatLine, 21) = "has entered the room." Then
        iDiff = Len(sChatLine) - 22
        sTheSN = Mid$(sChatLine, 1, iDiff)
        Call SendChat("+~-> Welcome " & sTheSN)
        GoTo sstart
    End If
Else
    GoTo sstart
End If
End Sub
Sub MMBot(Lst1 As ListBox, Trigger As TextBox)
'+~-> This Will Start A MMBot
'+~-> Ex. Call MMBot ((List1), (Text1))
MMBotStop = False
sstart:
If MMBotStop = False Then Exit Sub
If LastChatLine() = Trigger Then
    For x = 0 To x = x - 1
        If x = SNFromLastChatLine() Then GoTo sstart
        Lst1.AddItem SNFromLastChatLine()
    Next x
End If
GoTo sstart
End Sub
Sub ScramblerBot(sAnswer As TextBox, sHint As TextBox, Tmr1 As Timer, Tmr2 As Timer, Tmr3 As Timer)
'+~-> This Is A Scrambler Bot
'+~-> Ex. Call ScramblerBot(text1, text2, timer1, timer2, Timer3)
ScramblerStop = False
Tmr2.Enabled = True
ASCii = ("<FONT FACE=""Arial"">" & FadeByColor3(FADE_RED, FADE_BLACK, FADE_BLUE, "+~->", False) & "<FONT COLOR=""#0000FF""> ")
ScramblerAnswer = sAnswer.text
ScrambledAnswer = Text_Scrambled(sAnswer.text)
Call SendChat(ASCii & "§çràmß£è® ßø†")
Call SendChat(ASCii & "Word: " & ScrambledAnswer)
Call SendChat(ASCii & "Hint: " & sHint)
Tmr3.Enabled = True
Tmr1.Enabled = True
End Sub
Sub GuideBot()
'+~-> This Will Check If The Chat Room Has More Than 21 Peeps In It
'+~-> Ex. Call GuideBot - To Stop Put GuideStop = True
GuideStop = False
sstart:
If GuideStop = True Then Exit Sub
Dim sTheSN As String
Dim iDiff As Integer
Dim sChatLine As String
If SNFromLastChatLine() = "OnlineHost" Then
    sChatLine = LastChatLine()
    If Right$(sChatLine, 21) = "has entered the room." Then
        sRoom% = FindChatRoom()
        iDiff = Len(sChatLine) - 22
        sTheSN = Mid$(sChatLine, 1, iDiff)
        If sTheSN = LCase("Guide") Then
            Call SendMessage(sRoom%, WM_CLOSE, 0, 0)
        End If
        GoTo sstart
    End If
Else
    GoTo sstart
End If
End Sub
Sub CircleForm()
'+~-> This Will make The Form Circle Shaped
'+~-> Put The Stuff Below In The Load Part Of The Form W/O The "'"
'SetWindowRgn hwnd, _
'CreateEllipticRgn(0, 0, 300, 200), True
End Sub



Private Sub ListBox2Clipboard(Plop As ListBox)
'+~-> This Will Put A ListBoX Content Into The ClipBoard
'+~-> Need A ListBoX
'+~-> Example Call ListBox2Clipboard(List1)
Dim lSN As Long
Dim sTheList As String
For lSN = 0 To Plop.ListCount - 1
    If lSN = 0 Then
        sTheList = Plop.List(lSN)
    Else
        sTheList = sTheList & "," & Plop.List(lSN)
    End If
Next
Clipboard.Clear
Clipboard.SetText sTheList
End Sub

Sub LoadSNList(Plop As ListBox)
'+~-> This Will Load A SN List
'+~-> Need The CommonDialogControl, A ListBox
'+~-> Example Call LoadSNList(List1)
On Error GoTo CanErr
With CommonDialog1
    .DialogTitle = "Load SN List"
    .CancelError = True
    .Filter = "Text File (*.txt)|*.txt"
    .FilterIndex = 0
    .ShowOpen
End With
Dim sSNList As String
Open CommonDialog1.FileName For Input As #1
sSNList = Input(LOF(1), 1)
Close #1

Dim sChar As String
Dim sSN As String
Dim lPos As Long
sSN = ""
For lPos = 1 To Len(sSNList)
    sChar = Mid$(sSNList, lPos, 1)
    If sChar = "," Then
        Plop.AddItem sSN
        sSN = ""
    Else
         sSN = sSN & sChar
    End If
Next
Exit Sub

CanErr:
   On Error GoTo 0
   Exit Sub
End Sub

Private Sub Macro()
'open macro's path for input as #1
Text1.text = Input(LOF(1), 1)
Close #1
End Sub

Sub SaveSNList(lst As ListBox)
'+~-> This Will Save A SN List
'+~-> Call SaveSNList(List1)
On Error GoTo CancelErr
With CommonDialog1
    .CancelError = True
    .DialogTitle = "Save SNs"
    .Filter = "Text Files (*.txt)|*.txt"
    .FilterIndex = 0
    .ShowSave
End With

Dim sList As String
Dim lSN As Long
sList = ""
For lSN = 0 To lst.ListCount - 1
    If lSN = 0 Then
        sList = lst.List(lSN)
    Else
        sList = sList & "," & lst.List(lSN)
    End If
Next
Open CommonDialog1.FileName For Output As #1
Print #1, sList
Close #1

Exit Sub

CancelErr:
  On Error GoTo 0
  Exit Sub
End Sub



Function StringInList(thelist As ListBox, FindMe As String)
'+~-> This Will Find A String(Text or Word) In A List
'+~-> Call StringInList((List1), "Adobe")
If thelist.ListCount = 0 Then GoTo nope
For a = 0 To thelist.ListCount - 1
thelist.ListIndex = a
If UCase(thelist.text) = UCase(FindMe) Then
StringInList = a
Exit Function
End If
Next a
nope:
StringInList = -1
End Function

Function FindChildByClass(parentw, childhand)
'+~-> This Will Find A Child Window By Its Class Name
'+~-> Example MDI% = FindChildByClass(AOL%, "MDIClient")
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
FindChildByClass = 0

bone:
Room% = firs%
FindChildByClass = Room%

End Function

Function FindChildByTitle(parentw, childhand)
'+~-> This Will Find A Child Window By Its Title/Caption
'+~-> Example MDI% = FindChildByTitle(AOL%, "Welcome, *")
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitle = 0

bone:
Room% = firs%
FindChildByTitle = Room%
End Function

Function GetClass(child)
'+~-> This Will Get The Class Of A Window
buffer$ = String$(250, 0)
getclas% = GetClassName(child, buffer$, 250)

GetClass = buffer$
End Function

Function FindChatRoom()
'+~-> This Will Find The Chat Room
'+~-> Example If FindChatRoom Then MsgBox "Your In A Chat Room"
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Room% = FindChildByClass(mdi%, "AOL Child")
Stuff% = FindChildByClass(Room%, "_AOL_Listbox")
MoreStuff% = FindChildByClass(Room%, "RICHCNTL")
If Stuff% <> 0 And MoreStuff% <> 0 Then
   FindChatRoom = Room%
Else:
   FindChatRoom = 0
End If
End Function


Function UserSN()
'+~-> This Will Get The Users SN
'+~-> Example Label1.caption = UserSN
On Error Resume Next
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Welcome% = FindChildByTitle(mdi%, "Welcome, ")
WelcomeLength% = GetWindowTextLength(Welcome%)
WelcomeTitle$ = String$(200, 0)
a% = GetWindowText(Welcome%, WelcomeTitle$, (WelcomeLength% + 1))
User = Mid$(WelcomeTitle$, 10, (InStr(WelcomeTitle$, "!") - 10))
UserSN = User
End Function

Sub KillWait()
'+~-> This Gets Rid Of The HourGlass
'+~-> Call KillWait

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 19
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessage(KeyWordWin%, WM_CLOSE, 0, 0)
End Sub
Function IsUserOnline()
'+~-> This Will Look For The Welcome Window And If
'Its Not There Then It Thinks Your Not Online
'+~-> If IsUserOnline = 1 Then MsgBoX "Your Online"
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Welcome% = FindChildByTitle(mdi%, "Welcome,")
If Welcome% <> 0 Then
   IsUserOnline = 1
Else:
   IsUserOnline = 0
End If
End Function
Function GetCaption(hwnd)
'+~-> This Will Get The Caption Of A Window
'+~-> Example GetCaption(AOLWindow)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))

GetCaption = hwndTitle$
End Function

Sub SendChat(Chat)
'+~-> This Will Send A Message To The Chat Room
'+~-> Call SendChat("Sup Dawg") or Call SendChat(Text1)
Room% = FindChatRoom
AORich% = FindChildByClass(Room%, "RICHCNTL")

AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)
AORich% = GetWindow(AORich%, 2)

Call SetFocusAPI(AORich%)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, Chat)
DoEvents
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
Call SendMessageByNum(AORich%, WM_CHAR, 13, 0)
End Sub

Sub TimeOut(Duration)
'+~-> This Will Make A Pause
'+~-> timeout o.5
starttime = Timer
Do While Timer - starttime < Duration
DoEvents
Loop

End Sub



Sub StayOnTop(theform As Form)
'+~-> This Will Make The Form Stay On Top Of EveryThing
'+~-> Example StayOnTop Form69
SetWinOnTop = SetWindowPos(theform.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub Anti45MinTimer()
'+~-> This Will Click That Thing Thats Says Youve Been On For 45min
'+~-> Example: call Anti45MinTimer
AOTimer% = FindWindow("_AOL_Palette", vbNullString)
AOIcon% = FindChildByClass(AOTimer%, "_AOL_Icon")
ClickIcon (AOIcon%)
End Sub





Sub WAVPlay(File)
'+~-> Plays A Wav File
'+~-> Call WavPlay("Exit.wav")
SoundName$ = File
wFlags% = SND_ASYNC Or SND_NODEFAULT
x = sndPlaySound(SoundName$, wFlags%)
End Sub
Function WAVStop()
'+~-> This Will Stop The Wav Playing
'+~-> Call WavStop
Call WAVPlay(" ")
End Function

Sub RenameFile(File$, NewName$)
'+~-> This Will Rename A File
'+~-> Example Call RenameFile((Text1.text), ("BestBasFile"))
Name File$ As NewName$
NoFreeze% = DoEvents()
End Sub
Sub DeleteFile(File$)
'+~-> This Checks If The Specified File Exists Then Deletes It
'+~-> Example Call DeleteFile("C:\Windows\System\user32.exe")
If Not DoesFileExist(File$) Then Exit Sub
Kill File$
NoFreeze% = DoEvents()
End Sub
Function DoesFileExist(ByVal sFileName As String) As Integer
'+~-> This Will Check To See If A File Exists
'+~-> Example If Not DoesFileExist(Text) Then MsgBoX "Doesnt Exist"
Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        DoesFileExist = False
        Else
            DoesFileExist = True
    End If

End Function

Function DoesDirExist(TheDirectory)
'+~-> This Will Check If A Directory Exists
'+~-> Example If Not DoesDirExist (Text1) Then MsgBoX "File Don't Exist"
Dim Check As Integer
On Error Resume Next
If Right(TheDirectory, 1) <> "/" Then TheDirectory = TheDirectory + "/"
Check = Len(Dir$(TheDirectory))
If Err Or Check = 0 Then
    DoesDirExist = False
Else
    DoesDirExist = True
End If
End Function
Sub FormFlash(Frm As Form)
'+~-> This Will Make The Form Change Colors
'+~-> Call FormFlash(Form69)
Frm.Show
Frm.BackColor = &H0&
TimeOut (".1")
Frm.BackColor = &HFF&
TimeOut (".1")
Frm.BackColor = &HFF0000
TimeOut (".1")
Frm.BackColor = &HFF00&
TimeOut (".1")
Frm.BackColor = &H8080FF
TimeOut (".1")
Frm.BackColor = &HFFFF00
TimeOut (".1")
Frm.BackColor = &H80FF&
TimeOut (".1")
Frm.BackColor = &HC0C0C0
End Sub
Function MakeListOMails(lst As ListBox)
'+~-> This Will Make A List Of Your Mails
'+~-> Need A ListBoX And A Command Button
'+~-> Example Call MakeListOMails(List1)
AOLMDI
MailWin = FindChildByTitle(AOLMDI, "New Mail")

MailWin = FindChildByTitle(AOLMDI, "New Mail")
CountMail2
start:
If counter = CountMail2 Then GoTo last
MailTree = FindChildByClass(MailWin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, counter, 0)
    buffer$ = String$(namelen, 0)
    x = SendMessageByString(MailTree, LB_GETTEXT, counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    lst.AddItem buffer$
 TimeOut (0.001)
counter = counter + 1
GoTo start
last:

End Function

Function FindSendButton(dosloop)
'+~-> This Will Find The Send Button
'+~-> Example ClickIcon(FindSendButton)
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOLWindow(), "MDIClient"), 5)
forw% = FindChildByTitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = FindChildByTitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function


Function ForwardMail2(SN As String, message As String)
'+~-> This Will Forward The Currently Selected Mail
'+~-> Call ForwardMail("izekial83", "Heres Your Phat BAS")
FindForwardButton
PERSON = SN
Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Fwd: ")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop
a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
ClickIcon (icone%)
End Function
Function FindForwardButton()
'+~-> This Will Find The Forward Window
'+~-> Example ClickIcon(FindForwardWindow)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
childfocus% = GetWindow(mdi%, 5)

While childfocus%
listers% = FindChildByTitle(childfocus%, "Send Now")
listere% = FindChildByClass(childfocus%, "_AOL_Icon")
listerb% = FindChildByClass(childfocus%, "_AOL_Button")

If listers% <> 0 And listere% <> 0 And listerb% <> 0 Then FindForwardWindow = childfocus%: Exit Function
childfocus% = GetWindow(childfocus%, 2)
Wend
End Function




Function KeepAsNew()
'+~-> This Will Keep The Current Mail As NEW
'+~-> Call KeepAsNew
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
KeepAsNewButton = FindChildByTitle(A3000%, "Keep As New")
ClickIcon (KeepAsNewButon)
End Function
Function ListMail(Box As ListBox)
'+~-> This Will Put Your Mail In A ListBoX
'+~-> CallListMail(List1)
Box.Clear
AOLMDI
MailWin = FindChildByTitle(AOLMDI, "New Mail")
MailWin = FindChildByTitle(AOLMDI, "New Mail")
CountMail2
start:
If counter = CountMail2 Then GoTo last
MailTree = FindChildByClass(MailWin, "_AOL_TREE")
   namelen = SendMessage(MailTree, LB_GETTEXTLEN, counter, 0)
    buffer$ = String$(namelen, 0)
    x = SendMessageByString(MailTree, LB_GETTEXT, counter, buffer$)
    TabPos = InStr(buffer$, Chr$(9))
    buffer$ = Right$(buffer$, (Len(buffer$) - (TabPos)))
    Box.AddItem buffer$
 TimeOut (0.001)
counter = counter + 1
GoTo start
last:
End Function
Function OutGoingMailCount()
'+~-> Counts Your OutGoing Mail
theMail% = FindChildByClass(AOLMDI(), "AOL Child")
thetree% = FindChildByClass(theMail%, "_AOL_Tree")
Mail_Out_MailCount = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function

Function DeleteSingle()
'+~-> This Will Delete The Currently Selected File
'+~-> Call DeleteSingle
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
Delete% = FindChildByTitle(A3000%, "Delete")
ClickIcon (Delete%)
End Function

Function OpenCurrentMail()
'+~-> This Will Open The Current Mail Selected
'+~-> Call OpenCurrentMail
aol% = FindWindow("AOL Frame25", vbNullString)
A2000% = FindChildByClass(aol%, "MDIClient")
A3000% = FindChildByTitle(A2000%, "New Mail")
MailTree% = FindChildByClass(A3000%, "_AOL_Tree")
x = SendMessage(MailTree%, WM_KEYDOWN, VK_RETURN, 0)
x = SendMessage(MailTree%, WM_KEYUP, VK_RETURN, 0)
End Function
Sub SendNewMail(PERSON, SUBJECT, message)
'+~-> Sends A Mail
'+~-> Call SendNewMail ("izekial83", "Your BAS", "Youre Bas Is The Best")
Call RunMenuByString(AOLWindow(), "Compose Mail")

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
icone% = FindChildByClass(MailWin%, "_AOL_Icon")
peepz% = FindChildByClass(MailWin%, "_AOL_Edit")
subjt% = FindChildByTitle(MailWin%, "Subject:")
subjec% = GetWindow(subjt%, 2)
mess% = FindChildByClass(MailWin%, "RICHCNTL")
If icone% <> 0 And peepz% <> 0 And subjec% <> 0 And mess% <> 0 Then Exit Do
Loop

a = SendMessageByString(peepz%, WM_SETTEXT, 0, PERSON)
a = SendMessageByString(subjec%, WM_SETTEXT, 0, SUBJECT)
a = SendMessageByString(mess%, WM_SETTEXT, 0, message)
ClickIcon (icone%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
MailWin% = FindChildByTitle(mdi%, "Compose Mail")
erro% = FindChildByTitle(mdi%, "Error")
aolw% = FindWindow("_AOL_Modal", vbNullString)
If MailWin% = 0 Then Exit Do
If aolw% <> 0 Then
a = SendMessage(aolw%, WM_CLOSE, 0, 0)
ClickIcon (FindChildByTitle(aolw%, "OK"))
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Sub
End If
If erro% <> 0 Then
a = SendMessage(erro%, WM_CLOSE, 0, 0)
a = SendMessage(MailWin%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
last:
End Sub
Sub Scroll8Line(STR As TextBox)
'+~-> This Is An 8 Line Scroller
'+~-> Call Scroll8Line(Text1)
lonh = String(116, Chr(32))
d = 116 - Len(Text1)
c$ = Left(lonh, d)
SendChat ("" & STR & c$ & Text1)
SendChat ("" & STR & c$ & Text1)
lonh = String(116, Chr(32))
d = 116 - Len(Text1)
c$ = Left(lonh, d)
SendChat ("" & STR & c$ & Text1)
SendChat ("" & STR & c$ & Text1)
End Sub
Sub MailPreference()
'+~-> Stops It From Saying Youre Mail Has Been Sent
Call RunMenuByString(AOLWindow(), "Preferences")

Do: DoEvents
prefer% = FindChildByTitle(AOLMDI(), "Preferences")
maillab% = FindChildByTitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_HWNDNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

TimeOut (0.2)
ClickIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
aolcloses% = FindChildByTitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = FindChildByTitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = FindChildByTitle(aolmod%, "OK")
If aolOK% <> 0 And aolcloses% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(aolcloses%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

ClickIcon (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub
Sub BuddyList2ListBox(lst As ListBox)
'+~-> This adds the AOL Buddy List to a VB listbox
'+~-> Call BuddyList2ListBox(List1)
On Error Resume Next
Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
Room = FindChildByTitle(AOLMDI(), "Buddy List Window")
aolhandle = FindChildByClass(Room, "_AOL_Listbox")
AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)
If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6
PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)
PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
lst.AddItem PERSON$
Next Index
Call CloseHandle(AOLProcessThread)
End If
End Sub

Sub Scroll15Line(txt As String)
'+~-> 15 Line Scroll ~~~> Max Of 14 Characters Or It Says Too Much Text
'+~-> Call Scroll15Line("SteveCase's GAY")
Call ActivateAOL
a = String(116, Chr(32))
d = 116 - Len(txt)
c$ = Left(a, d)
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.8
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.8
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.8
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.8
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.3
SendChat "" + txt + "" & c$ & "" + txt + ""
TimeOut 0.8
End Sub
Sub ActivateAOL()
'+~-> Activates AOL No Matter What The Caption
'+~-> Example Call ActivateAOL
x = GetCaption(AOLWindow)
AppActivate x
End Sub
Sub ClickStartButton()
windows% = FindWindow("Shell_TrayWnd", vbNullString)
StartButton% = FindChildByClass(windows%, "Button")
ClickIcon (StartButton%)
End Sub
Sub CreateProgManItem(xForm As Form, CmdLine As String, IconTitle As String)
'+~-> This Is For The File ProgMan.exe its like Your start Menu
'+~-> This Will Add An Object to it
'+~-> Example Call CreateProgManItem(Form#Here, WhereTheFileIsHere, AndTheItemName)
    Dim i%
    Dim Z%
    Screen.MousePointer = 11
    
    
    On Error Resume Next


    
    xForm.DDELabel.LinkTopic = "ProgMan|Progman"
    xForm.DDELabel.LinkMode = 2
    For i% = 1 To 10
      Z% = DoEvents()
    Next
    xForm.DDELabel.LinkTimeout = 100
    
    xForm.DDELabel.LinkExecute "[AddItem(" + CmdLine + Chr$(44) + IconTitle + Chr$(44) + ",,)]"
    xForm.DDELabel.LinkExecute "[ShowGroup(groupname, 1)]"
    
    xForm.DDELabel.LinkTimeout = 50
    xForm.DDELabel.LinkMode = 0
    Screen.MousePointer = 0
End Sub
Function GetFreeDiskSpace(Drive As String) As Long
        
    On Error GoTo FErrorHandler
    Dim TempDrive As String
    Dim XValue As Long
    Dim DirTest As String
    TempDrive = Left$(CurDir$, 2)
    ChDrive Drive
    DirTest = Dir$(Drive & "\*.*")
    XValue = DiskSpaceFree&()
    ChDrive TempDrive
    GetFreeDiskSpace = XValue
    MsgBox "You Have " & XValue & "Free Space On " & (Text1.text)
    Exit Function

FErrorHandler:
    
    GetFreeDiskSpace = -Err
    On Error GoTo 0
    ChDrive TempDrive
    Err = 0
MsgBox "You Have " & -Err & "Free Space On "
    Exit Function

End Function
Sub Availible(PERSON)
'+~-> This Will Check If Someones Available
'+~-> Example Call Availible("izekial83")
Call KeyWord("aol://9293:")
TimeOut 1.7
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IM% = FindChildByTitle(mdi%, "Send Instant Message")
IMSendTo% = FindChildByClass(IM%, "_AOL_Edit")
Call SendMessageByString(IMSendTo%, WM_SETTEXT, 0, PERSON)
E = FindChildByClass(IM%, "RICHCNTL")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
ClickIcon (E)
TimeOut 1
OkWin% = FindWindow("#32770", "America Online")
e2 = FindChildByClass(OkWin%, "Static")
e2 = GetWindow(e2, GW_HWNDNEXT)
OkWinMsgMsg$ = GetText(e2)
If OkWinMsgMsg$ = PERSON & " is online and able to receive Instant Messages." Then
    Msg$ = " Can Be Punted"
    AvailibleYes = True
    GoTo Ending
ElseIf OkWinMsgMsg$ = PERSON & " is not currently signed on." Then
    Msg$ = " Ain't Online!"
    AvailibleYes = False
Else
    Msg$ = " Has IMs Off"
    AvailibleYes = False
End If
Ending:
    If OkWin% <> 0 Then
        Call SendMessage(OkWin%, WM_CLOSE, 0, 0)
        Call SendMessage(IM%, WM_CLOSE, 0, 0)
    End If
    Call MsgBox(PERSON & Msg$)
End Sub
Sub AIMSendChatInvite(Who$, message$, Chatname$)
'+~-> This Will Send A Chat Invitation
'+~-> Example Call AIMSendChatInvite("izekial83", "Come Here", "PhatBasMakers")
Call AIMOpenNewChatInvite
ChatInvite1% = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
ChatInvite2% = FindChildByClass(ChatInvite1%, "Edit")
ChatInvite3% = SendMessageByString(ChatInvite2%, WM_SETTEXT, 0, Who$)
ChatInvite4% = FindChildByTitle(ChatInvite1%, "You are invited to the following Buddy Chat: ")
ChatInvite5% = SendMessageByString(ChatInvite4%, WM_SETTEXT, 0, message$)
ChatInvite6% = FindChildByClass(ChatInvite1%, "_Oscar_Static")
ChatInvite7% = GetWindow(ChatInvite6%, GW_HWNDNEXT)
ChatInvite8% = GetWindow(ChatInvite7%, GW_HWNDNEXT)
ChatInvite9% = GetWindow(ChatInvite8%, GW_HWNDNEXT)
ChatInvite10% = GetWindow(ChatInvite9%, GW_HWNDNEXT)
ChatInvite11% = GetWindow(ChatInvite10%, GW_HWNDNEXT)
ChatInvite12% = GetWindow(ChatInvite11%, GW_HWNDNEXT)
ChatInvite13% = GetWindow(ChatInvite12%, GW_HWNDNEXT)
ChatInvite14% = SendMessageByString(ChatInvite13%, WM_SETTEXT, 0, Chatname$)
ChatInvite15% = FindChildByClass(ChatInvite1%, "_Oscar_IconBtn")
ChatInvite16% = GetWindow(ChatInvite15%, GW_HWNDNEXT)
ChatInvite17% = GetWindow(ChatInvite16%, GW_HWNDNEXT)
ClickIcon (ChatInvite17%)
End Sub
Sub AIMSendIM(Who$, What$)
'+~-> Sends An IM
'+~-> Example Call AIMSendIM("izekial83", "Hey Man Phat BAS")
Call AIMOpenNewIM
SendIM1% = FindWindow("AIM_IMessage", vbNullString)
SendIM2% = FindChildByClass(SendIM1%, "_Oscar_PersistantComb")
SendIM3% = FindChildByClass(SendIM2%, "Edit")
SendIM4% = SendMessageByString(SendIM3%, WM_SETTEXT, 0, Who$)
SendIM5% = FindChildByClass(SendIM1%, "Ate32class")
SendIM6% = GetWindow(SendIM5%, GW_HWNDNEXT)
SendIM7% = SendMessageByString(SendIM6%, WM_SETTEXT, 0, What$)
SendIM8% = FindChildByClass(SendIM1%, "_Oscar_IconBtn")
ClickIcon (SendIM8%)
End Sub
Sub AIMOpenNewIM()
'+~-> Opens An IM
'+~-> Example Call AIMOpenNewIM
OpenNewIM1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenNewIM2% = FindChildByClass(OpenNewIM1%, "_Oscar_TabGroup")
OpenNewIM3% = FindChildByClass(OpenNewIM2%, "_Oscar_IconBtn")
OpenNewIM4% = GetWindow(OpenNewIM3%, GW_HWNDNEXT)
ClickIcon (OpenNewIM4%)
End Sub
Sub AIMOpenNewChatInvite()
'+~-> Opens A New Chat Invitation
'+~-> Call AIMOpenNewChatInvite
OpenNewChatInvite1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
OpenNewChatInvite2% = FindChildByClass(OpenNewChatInvite1%, "_Oscar_TabGroup")
OpenNewChatInvite3% = FindChildByClass(OpenNewChatInvite2%, "_Oscar_IconBtn")
ClickIcon (OpenNewChatInvite3%)
End Sub
Sub AIMHideBuddyList()
'+~-> Will Hide Only The Buddys On The Buddy List Not The Whole Window
'+~-> Call AIMHideBuddyList
On Error Resume Next
FuckBuddyList1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
FuckBuddyList2% = FindChildByClass(FuckBuddyList1%, "_Oscar_IconBtn")
FuckBuddyList3% = FindChildByClass(FuckBuddyList1%, "_Oscar_TabGroup")
FuckBuddyList4% = FindChildByClass(FuckBuddyList1%, "Ate32Class")
Call SendMessage(FuckBuddyList2%, WM_CLOSE, 0, 0)
Call SendMessage(FuckBuddyList3%, WM_CLOSE, 0, 0)
Call SendMessage(FuckBuddyList4%, WM_CLOSE, 0, 0)
FuckBuddyList5% = FindChildByClass(FuckBuddyList1%, "Ate32Class")
Call SendMessage(FuckBuddyList5%, WM_CLOSE, 0, 0)
Call AIMChangeBuddyCaption("0wned")
End Sub
Sub AIMSendChat(text$)
'+~-> Sends Text To A Chat Room
'+~-> Ex. Call AIMSendChat("Yo AIM2 Is Phat")
Chatsend1% = FindWindow("AIM_ChatWnd", vbNullString)
    If Chatsend1% = 0 Then Info% = MsgBox("There is no chat room open, so please open one", vbInformation + vbOKOnly, "Error!")
Chatsend2% = FindChildByClass(Chatsend1%, "Ate32Class")
Chatsend3% = GetWindow(Chatsend2%, GW_HWNDNEXT)
ChatSend4% = SendMessageByString(Chatsend3%, WM_SETTEXT, 0, text$)
ChatSend5% = FindChildByClass(Chatsend1%, "_Oscar_IconBtn")
    ClickIcon (ChatSend5%)
End Sub
Sub AIMChangeBuddyCaption(newcaption$)
BuddyCaption1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
BuddyCaption2% = SendMessageByString(BuddyCaption1%, WM_SETTEXT, 0, newcaption$)
End Sub
Sub AIMChangeChatCaption(newcaption$)

ChatCaption1% = FindWindow("AIM_IMessage", vbNullString)
ChatCaption2% = SendMessageByString(ChatCaption1%, WM_SETTEXT, 0, newcaption$)
End Sub
Sub AIMChangeIMCaption(newcaption$)

IMCaption1% = FindWindow("AIM_IMessage", vbNullString)
IMCaption2% = SendMessageByString(IMCaption1%, WM_SETTEXT, 0, newcaption$)
End Sub
Sub AIMMassIM(List As ListBox, text As TextBox)

List.Enabled = False
People = List.ListCount - 1
List.ListIndex = 0
For MassIM1 = 0 To People
List.ListIndex = MassIM1
Call AIMSendIM(List.text, text.text)
TimeOut 1.5
Next MassIM1
List.Enabled = True
End Sub
Function AIMGetUserSn()
'+~-> This Will Get The Users AIM Name
'+~-> Example Label1.caption = AIMGetUserSN
On Error Resume Next
UserSn1% = FindWindow("_Oscar_BuddyListWin", vbNullString)
UserSn2% = GetWindowTextLength(UserSn1%)
UserSn3$ = String$(UserSn2%, 0)
UserSn4% = GetWindowText(UserSn1%, UserSn3$, (UserSn2% + 1))
If Not Right(UserSn3$, 13) = "'s Buddy List" Then Exit Function
UserSn5$ = Mid$(UserSn3$, 1, (UserSn2% - 13))
AIMGetUserSn = UserSn5$

End Function
Function AIMGetSnFromIM()
'+~-> This Will Get The SN From An IM In AIM
'+~-> Example Call AimGetSNFromIM
On Error Resume Next
SnFromIM1% = FindWindow("AIM_IMessage", vbNullString)
SnFromIM2% = GetWindowTextLength(SnFromIM1%)
SnFromIM3$ = String$(SnFromIM2%, 0)
SnFromIM4% = GetWindowText(SnFromIM1%, SnFromIM3$, (SnFromIM2% + 1))
If Not Right(SnFromIM3$, 18) = " - Instant Message" Then Exit Function
SnFromIM5$ = Mid$(SnFromIM3$, 1, (SnFromIM2% - 18))
AIMGetSnFromIM = SnFromIM5$

End Function
Sub WAVLoop(File)
'+~-> This Will Open A Wav File And Loop It
'+~-> Call WavLoop("Exit.wav")
SoundName$ = File
wFlags% = SND_ASYNC Or SND_LOOP
x = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub ClearChat2()
'+~-> Clears The Chat Window For Only You
'+~-> Call ClearChat2
childs% = FindChatRoom()
child = FindChildByClass(childs%, "RICHCNTL")
GetTrim = SendMessageByNum(child, 13, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 12, GetTrim + 1, TrimSpace$)
theview$ = TrimSpace$
End Sub
Sub KillWin(Windo)
'+~-> This Closes A Window
'+~-> Example Call KillWin(IM%)
x = SendMessageByNum(Windo, WM_CLOSE, 0, 0)
End Sub



Sub Text_Manipulation(Who$, wut$)
'+~-> This Will Make It Look Like SomeOne Else Said Something They Didnt
'+~-> Call Text_Manipulation("SteveCase", "Im A HoMO")
aol% = FindWindow("AOL Frame25", 0&)
chatwin% = FindChildByClass(aol%, "MDIClient")
blah% = FindChildByClass(chatwin%, "AOL Child")
ChatVew% = FindChildByClass(blah%, "_AOL_View")
sndtext% = SendMessageByString(ChatVew%, WM_SETTEXT, 0, Chr(13) & Chr(9) & (Who$) & ":" & Chr(9) & (wut$))
SendNow% = SendMessageByNum(ChatVew%, WM_CHAR, &HD, 0)
End Sub

Public Sub ClearChatWin()
'+~-> This Clears The ChatRoom Text
'+~-> Need A Command Button
'+~-> Example Call ClearChatWin
getpar% = FindChatRoom()
child = FindChildByClass(getpar%, "RICHCNTL")
child.text = ""
End Sub
Sub MacroScroll(text$)
'+~-> This Will Scroll A TextBox That Has A Macro In It
'+~-> Need A TextBoX, And A Command Button
'+~-> Example: Call MacroScroll (Text1)
If Mid(text$, Len(text$), 1) <> Chr$(10) Then
    text$ = text$ + Chr$(13) + Chr$(10)
End If
Do While (InStr(text$, Chr$(13)) <> 0)
    counter = counter + 1
    SendChat Mid(text$, 1, InStr(text$, Chr(13)) - 1)
    If counter = 4 Then
        TimeOut (2.9)
        counter = 0
    End If
    text$ = Mid(text$, InStr(text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub
Sub AddRoomToTextBox(thelist As ListBox, text As TextBox)
'+~-> This Will Add A ChatRoom To A TextBox
'+~-> Need A TextBox, ListBox
'+~-> Example Call AddRoomToTextBox (List1, Text1)
Dim Y
Call AddRoomToListBox(thelist)
For Y = 0 To thelist.ListCount - 1
tt$ = tt$ + thelist.List(Y) + ","
Next Y
TimeOut (0.01)
text.text = tt$

End Sub

Sub FormFade(FormX As Form, Colr1, Colr2)
    B1 = GetRGB(Colr1).Blue
    G1 = GetRGB(Colr1).Green
    R1 = GetRGB(Colr1).Red
    B2 = GetRGB(Colr2).Blue
    G2 = GetRGB(Colr2).Green
    R2 = GetRGB(Colr2).Red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub

Sub FadeForm(FormX As Form, Colr1, Colr2)
'+~-> This Will Fade The Form By 2 Colors (Like a Setup File (Blue - Black ))
'+~-> Call FadeForm(WhatFormHere, Fade_Blue, Fade_Black) That'd Fade Blue To Black
    B1 = GetRGB(Colr1).Blue
    G1 = GetRGB(Colr1).Green
    R1 = GetRGB(Colr1).Red
    B2 = GetRGB(Colr2).Blue
    G2 = GetRGB(Colr2).Green
    R2 = GetRGB(Colr2).Red
    
    On Error Resume Next
    Dim intLoop As Integer
    FormX.DrawStyle = vbInsideSolid
    FormX.DrawMode = vbCopyPen
    FormX.ScaleMode = vbPixels
    FormX.DrawWidth = 2
    FormX.ScaleHeight = 256
    For intLoop = 0 To 255
        FormX.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(((R2 - R1) / 255 * intLoop) + R1, ((G2 - G1) / 255 * intLoop) + G1, ((B2 - B1) / 255 * intLoop) + B1), B
    Next intLoop
End Sub
Sub FadePreview(PicB As PictureBox, ByVal FadedText As String)
'+~-> This Will Give You Of A Preview Of What Your Gonna Fade in a Picture BoX
'+~-> Example: X = FadeByColor3(Color1.BackColor, Color2.BackColor, Color3.BackColor, WhatYaWannaFade, FalseIsNotWavY)
'              Text2 = X
'              Call FadePreview(Picture1, text2.Text)
'+~-> That Will Preview A 3 Color Fade - You Can it Change to 10 Or Whatever Ya Want
FadedText$ = Replacer(FadedText$, Chr(13), "+chr13+")
OSM = PicB.ScaleMode
PicB.ScaleMode = 3
TextOffX = 0: TextOffY = 0
StartX = 2: StartY = 0
PicB.Font = "Arial": PicB.FontSize = 10
PicB.FontBold = False: PicB.FontItalic = False: PicB.FontUnderline = False: PicB.FontStrikethru = False
PicB.AutoRedraw = True: PicB.ForeColor = 0&: PicB.Cls
For x = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, x, 1)
  If c$ = "<" Then
    TagStart = x + 1
    TagEnd = InStr(x + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    x = TagEnd + 1
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
      Case "/sub" 'end subscript
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
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If c$ = "+" And Mid(FadedText$, x, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        x = x + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print c$
        TextOffX = TextOffX + PicB.TextWidth(c$)
    End If
  End If
Next x
PicB.ScaleMode = OSM
End Sub
Function GETVAL%(ByVal strLetter$)
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
Function Replacer(TheStr As String, This As String, WithThis As String)
Dim STRwo13s As String
STRwo13s = TheStr
Do While InStr(1, STRwo13s, This)
DoEvents
thepos% = InStr(1, STRwo13s, This)
STRwo13s = Left(STRwo13s, (thepos% - 1)) + WithThis + Right(STRwo13s, Len(STRwo13s) - (thepos% + Len(This) - 1))
Loop

Replacer = STRwo13s
End Function
Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function
Sub FadePreview2(RichTB As Control, ByVal FadedText As String)
'+~-> This Will Give You Of A Preview Of What Your Gonna Fade In A RichText BoX
'+~-> Example: X = FadeByColor3(Color1.BackColor, Color2.BackColor, Color3.BackColor, WhatYaWannaFade, FalseIsNotWavY)
'              Text2 = X
'              Call FadePreview(RichTextBox1, text2.Text)
'+~-> That Will Preview A 3 Color Fade - You Can it Change to 10 Or Whatever Ya Want
Dim StartPlace%
StartPlace% = 0
RichTB.SelStart = StartPlace%
RichTB.Font = "Arial": RichTB.SelFontSize = 10
RichTB.SelBold = False: RichTB.SelItalic = False: RichTB.SelUnderline = False: RichTB.SelStrikeThru = False
RichTB.SelColor = 0&: RichTB.text = ""
For x = 1 To Len(FadedText$)
  c$ = Mid$(FadedText$, x, 1)
  RichTB.SelStart = StartPlace%
  RichTB.SelLength = 1
  If c$ = "<" Then
    TagStart = x + 1
    TagEnd = InStr(x + 1, FadedText$, ">") - 1
    T$ = LCase$(Mid$(FadedText$, TagStart, (TagEnd - TagStart) + 1))
    x = TagEnd + 1
    RichTB.SelStart = StartPlace%
    RichTB.SelLength = 1
    Select Case T$
      Case "u"
        RichTB.SelUnderline = True
      Case "/u"
        RichTB.SelUnderline = False
      Case "s"
        RichTB.SelStrikeThru = True
      Case "/s"
        RichTB.SelStrikeThru = False
      Case "b"    'start bold
        RichTB.SelBold = True
      Case "/b"   'stop bold
        RichTB.SelBold = False
      Case "i"    'start italic
        RichTB.SelItalic = True
      Case "/i"   'stop italic
        RichTB.SelItalic = False
      
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
          RichTB.SelStart = StartPlace%
          RichTB.SelColor = RGB(RV, GV, BV)
        End If
        If Left$(T$, 9) = "font face" Then
            fontstart% = InStr(T$, Chr(34))
            dafont$ = Right(T$, Len(T$) - fontstart%)
            RichTB.SelStart = StartPlace%
            RichTB.SelFontName = dafont$
        End If
    End Select
  Else  'normal text
    RichTB.SelText = RichTB.SelText + c$
    StartPlace% = StartPlace% + 1
    RichTB.SelStart = StartPlace%
  End If
Next x
End Sub

Function Hex2Dec!(ByVal strHex$)
'+~-> This Will Convert Hex To Dec  -  Duh
'+~-> X = Hex2Dec(HexCodeHere)
'     Text1.Text = X
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For x = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), x, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - x)
  Next x
End Function





Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, thetext$, WavY As Boolean)
'+~-> This Will Fade 10 Colors(No Scroll Bars)
'+~-> Example Call FadeByColor10(Fade_Green, Fade_Red, Fade_Blue, Fade_Green, Fade_Red, Fade_Blue, Fade_Green, Fade_Red, Fade_Blue, Fade_Black, "WhatToSayHere", TrueMeansWavy)
'+~-> Now To Send That To The ChatRoom You Would Go
'+~-> X = FadeByColor10(Fade_Green, Fade_Red, Fade_Blue, Fade_Green, Fade_Red, Fade_Blue, Fade_Green, Fade_Red, Fade_Blue, Fade_Black, "WhatToSayHere", TrueMeansWavy)
'+~-> SendChat(X)
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)
dacolor6$ = RGBtoHEX(Colr6)
dacolor7$ = RGBtoHEX(Colr7)
dacolor8$ = RGBtoHEX(Colr8)
dacolor9$ = RGBtoHEX(Colr9)
dacolor10$ = RGBtoHEX(Colr10)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))
rednum6% = Val("&H" + Right(dacolor6$, 2))
greennum6% = Val("&H" + Mid(dacolor6$, 3, 2))
bluenum6% = Val("&H" + Left(dacolor6$, 2))
rednum7% = Val("&H" + Right(dacolor7$, 2))
greennum7% = Val("&H" + Mid(dacolor7$, 3, 2))
bluenum7% = Val("&H" + Left(dacolor7$, 2))
rednum8% = Val("&H" + Right(dacolor8$, 2))
greennum8% = Val("&H" + Mid(dacolor8$, 3, 2))
bluenum8% = Val("&H" + Left(dacolor8$, 2))
rednum9% = Val("&H" + Right(dacolor9$, 2))
greennum9% = Val("&H" + Mid(dacolor9$, 3, 2))
bluenum9% = Val("&H" + Left(dacolor9$, 2))
rednum10% = Val("&H" + Right(dacolor10$, 2))
greennum10% = Val("&H" + Mid(dacolor10$, 3, 2))
bluenum10% = Val("&H" + Left(dacolor10$, 2))


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, thetext, WavY)

End Function

Function FadeByColor2(Colr1, Colr2, thetext$, WavY As Boolean)
'+~-> This Will Fade 2 Colors(No Scroll Bars)
'+~-> Example Call FadeByColor10(Fade_Green, Fade_Red, "WhatToSayHere", TrueMeansWavy)
'+~-> Now To Send That To The ChatRoom You Would Go
'+~-> X = FadeByColor2(Fade_Green, Fade_Red, "WhatToSayHere", TrueMeansWavy)
'+~-> SendChat(X)
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, thetext, WavY)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, thetext$, WavY As Boolean)
'+~-> This Will Fade 3 Colors(No Scroll Bars)
'+~-> Example Call FadeByColor3(Fade_Green, Fade_Red, Fade_Blue, "WhatToSayHere", TrueMeansWavy)
'+~-> Now To Send That To The ChatRoom You Would Go
'+~-> X = FadeByColor3(Fade_Green, Fade_Red, Fade_Blue, "WhatToSayHere", TrueMeansWavy)
'+~-> SendChat(X)
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, thetext, WavY)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, thetext$, WavY As Boolean)
'+~-> This Will Fade 4 Colors(No Scroll Bars)
'+~-> Example Call FadeByColor4(Fade_Green, Fade_Red, Fade_Blue, Fade_Green, "WhatToSayHere", TrueMeansWavy)
'+~-> Now To Send That To The ChatRoom You Would Go
'+~-> X = FadeByColor4(Fade_Green, Fade_Red, Fade_Blue, Fade_Green, "WhatToSayHere", TrueMeansWavy)
'+~-> SendChat(X)
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, thetext, WavY)

End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, thetext$, WavY As Boolean)
'+~-> This Will Fade 5 Colors(No Scroll Bars)
'+~-> Example Call FadeByColor5(Fade_Green, Fade_Red, Fade_Blue, Fade_Green, Fade_Red, "WhatToSayHere", TrueMeansWavy)
'+~-> Now To Send That To The ChatRoom You Would Go
'+~-> X = FadeByColor5(Fade_Green, Fade_Red, Fade_Blue, Fade_Green, Fade_Red, "WhatToSayHere", TrueMeansWavy)
'+~-> SendChat(X)
dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)
dacolor3$ = RGBtoHEX(Colr3)
dacolor4$ = RGBtoHEX(Colr4)
dacolor5$ = RGBtoHEX(Colr5)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))
rednum3% = Val("&H" + Right(dacolor3$, 2))
greennum3% = Val("&H" + Mid(dacolor3$, 3, 2))
bluenum3% = Val("&H" + Left(dacolor3$, 2))
rednum4% = Val("&H" + Right(dacolor4$, 2))
greennum4% = Val("&H" + Mid(dacolor4$, 3, 2))
bluenum4% = Val("&H" + Left(dacolor4$, 2))
rednum5% = Val("&H" + Right(dacolor5$, 2))
greennum5% = Val("&H" + Mid(dacolor5$, 3, 2))
bluenum5% = Val("&H" + Left(dacolor5$, 2))

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, thetext, WavY)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, thetext$, WavY As Boolean)
'+~-> This Will Fade 5 Colors(Scroll Bars Needed)
'+~-> Example Call Fade5ColorFadeByColor5(Color1.BackColor, Color2.BackColor, Color3.BackColor, Color4.BackColor, Color5.BackColor, "What To Say Here", TrueMeansWavy)
'+~-> Now To Send That To The ChatRoom You Would Go
'+~-> X = Fade5ColorFadeByColor5(Color1.BackColor, Color2.BackColor, Color3.BackColor, Color4.BackColor, Color5.BackColor, "What To Say Here", TrueMeansWavy)
'+~-> SendChat(X)
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(thetext, frthlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFiveColor = Faded1$ + Faded2$ + Faded3$ + Faded4$
End Function
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, thetext$, WavY As Boolean)
'+~-> This Will Fade 10 Colors With ScrollBars
'+~-> Example: X = FadeTenColor(Color1.BackColor, Color2.BackColor, Color3.BackColor, Color4.BackColor, Color5.BackColor, Color6.BackColor, Color7.BackColor, Color8.BackColor, Color9.BackColor, Color10.BackColor, WhatToFade, TrueisWavY)
'              SendChat (X)
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    frthlen% = frthlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    fithlen% = fithlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    sixlen% = sixlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    eightlen% = eightlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    ninelen% = ninelen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Mid(thetext, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(thetext, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(thetext, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(thetext, ninelen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part4
    textlen% = Len(part4$)
    For i = 1 To textlen%
        TextDone$ = Left(part4$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B5 - B4) / textlen% * i) + B4, ((G5 - G4) / textlen% * i) + G4, ((R5 - R4) / textlen% * i) + R4)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded4$ = Faded4$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part5
    textlen% = Len(part5$)
    For i = 1 To textlen%
        TextDone$ = Left(part5$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B6 - B5) / textlen% * i) + B5, ((G6 - G5) / textlen% * i) + G5, ((R6 - R5) / textlen% * i) + R5)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded5$ = Faded5$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part6
    textlen% = Len(part6$)
    For i = 1 To textlen%
        TextDone$ = Left(part6$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B7 - B6) / textlen% * i) + B6, ((G7 - G6) / textlen% * i) + G6, ((R7 - R6) / textlen% * i) + R6)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded6$ = Faded6$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part7
    textlen% = Len(part7$)
    For i = 1 To textlen%
        TextDone$ = Left(part7$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B8 - B7) / textlen% * i) + B7, ((G8 - G7) / textlen% * i) + G7, ((R8 - R7) / textlen% * i) + R7)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded7$ = Faded7$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part8
    textlen% = Len(part8$)
    For i = 1 To textlen%
        TextDone$ = Left(part8$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B9 - B8) / textlen% * i) + B8, ((G9 - G8) / textlen% * i) + G8, ((R9 - R8) / textlen% * i) + R8)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded8$ = Faded8$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    'part9
    textlen% = Len(part9$)
    For i = 1 To textlen%
        TextDone$ = Left(part9$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B10 - B9) / textlen% * i) + B9, ((G10 - G9) / textlen% * i) + G9, ((R10 - R9) / textlen% * i) + R9)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded9$ = Faded9$ + "<Font Color=#" & colorx2 & ">" + TheHTML + LastChr$
    Next i
    
    FadeTenColor = Faded1$ + Faded2$ + Faded3$ + Faded4$ + Faded5$ + Faded6$ + Faded7$ + Faded8$ + Faded9$
End Function


Function InverseColor(OldColor)
'+~-> This Will Get The Opposite Of The Current Color
'+~-> Example X = InverseColor(OldColorHere)
'     Text1.text = X
'+~-> That Will Give You The Inverse In A TextBox
dacolor$ = RGBtoHEX(OldColor)
RedX% = Val("&H" + Right(dacolor$, 2))
GreenX% = Val("&H" + Mid(dacolor$, 3, 2))
BlueX% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - RedX%
newgreen% = 255 - GreenX%
newblue% = 255 - BlueX%
InverseColor = RGB(newred%, newgreen%, newblue%)

End Function

Function MultiFade(NumColors%, TheColors(), thetext$, WavY As Boolean)
'+~-> This Will Let You Fade As Many Colors As You Want
'+~-> Example X = MultiFade(NumberOfColorsToFadeHere, TheColors(), WhatToFadeHere, TrueMeansWavy)
Dim WaveState%
Dim WaveHTML$
WaveState = 0

If NumColors < 1 Then
MsgBox "Error: Attempting to fade less than one color."
MultiFade = thetext
Exit Function
End If

If NumColors = 1 Then
blah$ = RGBtoHEX(TheColors(1))
redpart% = Val("&H" + Right(blah$, 2))
greenpart% = Val("&H" + Mid(blah$, 3, 2))
bluepart% = Val("&H" + Left(blah$, 2))
blah2 = RGB(bluepart%, greenpart%, redpart%)
blah3$ = RGBtoHEX(blah2)

MultiFade = "<Font Color=#" + blah3$ + ">" + thetext
Exit Function
End If

Dim RedList%()
Dim GreenList%()
Dim BlueList%()
Dim DaColors$()
Dim DaLens%()
Dim DaParts$()
Dim Faded$()

ReDim RedList%(NumColors)
ReDim GreenList%(NumColors)
ReDim BlueList%(NumColors)
ReDim DaColors$(NumColors)
ReDim DaLens%(NumColors - 1)
ReDim DaParts$(NumColors - 1)
ReDim Faded$(NumColors - 1)

For Q% = 1 To NumColors
DaColors(Q%) = RGBtoHEX(TheColors(Q%))
Next Q%

For W% = 1 To NumColors
RedList(W%) = Val("&H" + Right(DaColors(W%), 2))
GreenList(W%) = Val("&H" + Mid(DaColors(W%), 3, 2))
BlueList(W%) = Val("&H" + Left(DaColors(W%), 2))
Next W%

textlen% = Len(thetext)
Do: DoEvents
For F% = 1 To (NumColors - 1)
DaLens(F%) = DaLens(F%) + 1: textlen% = textlen% - 1
If textlen% < 1 Then Exit For
Next F%
Loop Until textlen% < 1
    
DaParts(1) = Left(thetext, DaLens(1))
DaParts(NumColors - 1) = Right(thetext, DaLens(NumColors - 1))
    
dastart% = DaLens(1) + 1

If NumColors > 2 Then
For E% = 2 To NumColors - 2
DaParts(E%) = Mid(thetext, dastart%, DaLens(E%))
dastart% = dastart% + DaLens(E%)
Next E%
End If

For r% = 1 To (NumColors - 1)
textlen% = Len(DaParts(r%))
For i = 1 To textlen%
    TextDone$ = Left(DaParts(r%), i)
    LastChr$ = Right(TextDone$, 1)
    ColorX = RGB(((BlueList(r% + 1) - BlueList(r%)) / textlen% * i) + BlueList(r%), ((GreenList%(r% + 1) - GreenList(r%)) / textlen% * i) + GreenList(r%), ((RedList(r% + 1) - RedList(r%)) / textlen% * i) + RedList(r%))
    colorx2 = RGBtoHEX(ColorX)
        
    If WavY = True Then
    WaveState = WaveState + 1
    If WaveState > 4 Then WaveState = 1
    If WaveState = 1 Then WaveHTML = "<sup>"
    If WaveState = 2 Then WaveHTML = "</sup>"
    If WaveState = 3 Then WaveHTML = "<sub>"
    If WaveState = 4 Then WaveHTML = "</sub>"
    Else
    WaveHTML = ""
    End If
        
    Faded(r%) = Faded(r%) + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
Next i
Next r%

For qwe% = 1 To (NumColors - 1)
FadedTxtX$ = FadedTxtX$ + Faded(qwe%)
Next qwe%

MultiFade = FadedTxtX$

End Function


Function RGBtoHEX(RGB)
'+~-> This Will Convert RGB Colors To HEX Colors
'+~-> X = RGBToHEX(RGBColorValueHere)
'     Text1.text = X
    a$ = Hex(RGB)
    b% = Len(a$)
    If b% = 5 Then a$ = "0" & a$
    If b% = 4 Then a$ = "00" & a$
    If b% = 3 Then a$ = "000" & a$
    If b% = 2 Then a$ = "0000" & a$
    If b% = 1 Then a$ = "00000" & a$
    RGBtoHEX = a$
End Function

Function Rich2HTML(RichTXT As Control, StartPos%, EndPos%)
'+~-> This Will Convert Rich Text to HTML
Dim Bolded As Boolean
Dim Undered As Boolean
Dim Striked As Boolean
Dim Italiced As Boolean
Dim LastCRL As Long
Dim LastFont As String
Dim HTMLString As String

For posi% = StartPos To EndPos
RichTXT.SelStart = posi%
RichTXT.SelLength = 1

If Bolded <> RichTXT.SelBold Or posi% = StartPos Then
If RichTXT.SelBold = True Then
HTMLString = HTMLString + "<b>"
Bolded = True
Else
HTMLString = HTMLString + "</b>"
Bolded = False
End If
End If

If Undered <> RichTXT.SelUnderline Or posi% = StartPos Then
If RichTXT.SelUnderline = True Then
HTMLString = HTMLString + "<u>"
Undered = True
Else
HTMLString = HTMLString + "</u>"
Undered = False
End If
End If

If Striked <> RichTXT.SelStrikeThru Or posi% = StartPos Then
If RichTXT.SelStrikeThru = True Then
HTMLString = HTMLString + "<s>"
Striked = True
Else
HTMLString = HTMLString + "</s>"
Striked = False
End If
End If

If Italiced <> RichTXT.SelItalic Or posi% = StartPos Then
If RichTXT.SelItalic = True Then
HTMLString = HTMLString + "<i>"
Italiced = True
Else
HTMLString = HTMLString + "</i>"
Italiced = False
End If
End If

If LastCRL <> RichTXT.SelColor Or posi% = StartPos Then
ColorX = RGB(GetRGB(RichTXT.SelColor).Blue, GetRGB(RichTXT.SelColor).Green, GetRGB(RichTXT.SelColor).Red)
colorhex = RGBtoHEX(ColorX)
HTMLString = HTMLString + "<Font Color=#" & colorhex & ">"
LastCRL = RichTXT.SelColor
End If

If LastFont <> RichTXT.SelFontName Then
HTMLString = HTMLString + "<font face=" + Chr(34) + RichTXT.SelFontName + Chr(34) + ">"
LastFont = RichTXT.SelFontName
End If

HTMLString = HTMLString + RichTXT.SelText
Next posi%

Rich2HTML = HTMLString

End Function

Function HTMLtoRGB(HTMLColor$)
'+~-> This Will Convert HTML Colors Too RGB Colors So You Can Use It In VB
'+~-> X = HTMLtoRGB(HTMLColorHere)
'     Text1.text = X
If Left(HTMLColor$, 1) = "#" Then HTMLColor$ = Right(HTMLColor$, 6)

RedX$ = Left(HTMLColor$, 2)
GreenX$ = Mid(HTMLColor$, 3, 2)
BlueX$ = Right(HTMLColor$, 2)
rgbhex$ = "&H00" + BlueX$ + GreenX$ + RedX$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, thetext$, WavY As Boolean)
'+~-> This Will Fade 4 Colors With ScrollBars
'+~-> Example: X = FadeFourColor(Color1.BackColor, Color2.BackColor, Color3.BackColor, Color4.BackColor, WhatToFade, TrueIsWavy)
'              SendChat (X)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(thetext, fstlen%)
    part2$ = Mid(thetext, fstlen% + 1, seclen%)
    part3$ = Right(thetext, thrdlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    'part3
    textlen% = Len(part3$)
    For i = 1 To textlen%
        TextDone$ = Left(part3$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B4 - B3) / textlen% * i) + B3, ((G4 - G3) / textlen% * i) + G3, ((R4 - R3) / textlen% * i) + R3)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded3$ = Faded3$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    FadeFourColor = Faded1$ + Faded2$ + Faded3$
End Function

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, thetext$, WavY As Boolean)
'+~-> This Will Fade 3 Colors With ScrollBars
'+~-> Example: X = FadeThreeColor(Color1.BackColor, Color2.BackColor, Color3.BackColor, WhatToFade, TrueIsWavy)
'              SendChat (X)
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(thetext)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(thetext, fstlen%)
    part2$ = Right(thetext, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded1$ = Faded1$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    'part2
    textlen% = Len(part2$)
    For i = 1 To textlen%
        TextDone$ = Left(part2$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B3 - B2) / textlen% * i) + B2, ((G3 - G2) / textlen% * i) + G2, ((R3 - R2) / textlen% * i) + R2)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded2$ = Faded2$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    
    
    FadeThreeColor = Faded1$ + Faded2$
End Function

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, thetext$, WavY As Boolean)
'+~-> This Will Fade 2 Colors With ScrollBars
'+~-> Example: X = FadeTwoColor(Color1.BackColor, Color2.BackColor, WhatToFade, TrueIsWavy)
'              SendChat (X)
    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen$ = Len(thetext)
    For i = 1 To textlen$
        TextDone$ = Left(thetext, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If WavY = True Then
        WaveState = WaveState + 1
        If WaveState > 4 Then WaveState = 1
        If WaveState = 1 Then WaveHTML = "<sup>"
        If WaveState = 2 Then WaveHTML = "</sup>"
        If WaveState = 3 Then WaveHTML = "<sub>"
        If WaveState = 4 Then WaveHTML = "</sub>"
        Else
        WaveHTML = ""
        End If
        
        Faded$ = Faded$ + "<Font Color=#" & colorx2 & ">" + WaveHTML + LastChr$
    Next i
    FadeTwoColor = Faded$
End Function

Sub Scroll50Line(text As TextBox)
'+~-> This Will Scroll What You Want With 1 ChatSend About 50 Lines
'+~-> Need A TextBox And A Command Button
'+~-> Example Call Scroll50Line(Text1)
lineee = "Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text"
lineeee = "Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text Text"
Call SendChat("<PRE<" & lineee & lineeee & ">>")
End Sub


Public Sub ScrollList(lst As ListBox)
'+~-> This Will Scroll Names That Are In A ListBoX
'+~-> Example: Call ScrollList(List1)
For x% = 0 To lst.ListCount - 1
SendChat ("Scrolling Name [" & x% & "]" & lst.List(x%))
TimeOut (0.75)
Next x%
End Sub
Function Text_Dots(strin As String)
'+~-> Its Will Return The Text Like Sup.Dawg.Whats.Up
'+~-> X = Text_Dots("Hey")
'+~-> Text1.text = X
'+~-> Mess Around Wit It
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "•"
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Dots = Newsent$

End Function

Function Text_Html(strin As String)
'+~-> Returns The Text In HTML Format (Great Lagger)
'+~-> X = Text_HTML("YO")
'+~-> Text1.text = X
Let inptxt$ = strin
Let lenth% = Len(inptxt$)
Do While NumSpc% <= lenth%
Let NumSpc% = NumSpc% + 1
Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
Let NextChr$ = NextChr$ + "<html>"
Let Newsent$ = Newsent$ + NextChr$
Loop
Text_Html = Newsent$

End Function
Sub PhishPhrases(txt As TextBox)
'+~-> This Is A List Of Phish Phrases My Friend Made It Will Randomize ItsSelf
'+~-> You Need a TextBoX and A Command Button
'+~-> Example Call PhishPhrases(Text1)
Randomize x
phraZes = Int((Val("140") * Rnd) + 1)
If phraZes = "1" Then
txt = "Hi, I'm with AOL's Online Security. We have found hackers trying to get into your MailBox. Please verify your password immediately to avoid account termination.     Thank you.                                    AOL Staff"
ElseIf phraZes = "2" Then
txt = "Hello. I am with AOL's billing department. Due to some invalid information, we need you to verify your log-on password to avoid account cancellation. Thank you, and continue to enjoy America Online."
ElseIf phraZes = "3" Then
txt = "Good Evening. I am with AOL's Virus Protection Group. Due to some evidence of virus uploading, I must validate your sign-on password. Please STOP what you're doing and Tell me your password.       -- AOL VPG"
ElseIf phraZes = "4" Then
txt = "Hello, I am the Head Of AOL's XPI Link Department. Due to a configuration error in your version of AOL, I need you to verify your log-on password to me, to prevent account suspension and possible termination.  Thank You."
ElseIf phraZes = "5" Then
txt = "Hi. You are speaking with AOL's billing manager, Steve Case. Due to a virus in one of our servers, I am required to validate your password. You will be awarded an extra 10 FREE hours of air-time for the inconvenience."
ElseIf phraZes = "6" Then
txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "7" Then
txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "8" Then
txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "9" Then
txt = "Hi, I'm Alex Troph of America Online Sevice Department. Your online account, #3560028, is displaying a billing error. We need you to respond back with your name, address, card number, expiration date, and daytime phone number. Sorry for this inconvenience."
ElseIf phraZes = "10" Then
txt = "Hello, I am a representative of the VISA Corp.  Due to a computer error, we are unable complete your membership to America Online. In order to correct this problem, we ask that you hit the `Respond` key, and reply with your full name and password, so that the proper changes can be made to avoid cancellation of your account. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "11" Then
txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records. Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again. Thank you.  :-)"
ElseIf phraZes = "12" Then
txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Telephone#, Visa Card#, and Expiration date. If this information is not processed promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "13" Then
txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that your validation process is almost complete.  To complete your validation process I need you to please hit the `Respond` key and reply with the following information: Name, Address, Phone Number, City, State, Zip Code,  Credit Card Number, Expiration Date, and Bank Name.  Thank you for your time and cooperation and we hope that you enjoy America Online. :-)"
ElseIf phraZes = "14" Then
txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation."
ElseIf phraZes = "15" Then
txt = "Hello, this is the America Online Billing Department.  Due to a System Crash, we have lost your billing information.  Please hit respond, then enter your Credit Card Number, and experation date.  Thank You, and sorry for the inconvience."
ElseIf phraZes = "16" Then
txt = "ATTENTION:  This is the America Online billing department.  Due to an error that occured in your Membership Enrollment process we did not receieve your billing information.  We need you to reply back with your Full name, Credit card number with Expiration date, and your telephone number.  We are very sorry for this inconvenience and hope that you continue to enjoy America Online in the future."
ElseIf phraZes = "17" Then
txt = "Sorry, there seems to be a problem with your bill. Please reply with your password to verify that you are the account holder.  Thank you."
ElseIf phraZes = "18" Then
txt = "Sorry  the credit card you entered is invalid. Perhaps you mistyped it?  Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it. Thank you and enjoy AOL."
ElseIf phraZes = "19" Then
txt = "Sorry, your credit card failed authorization. Please reply with your credit card number, expiration date, name on card, billing address, and phone number, and we will fix it.  Thank you and enjoy AOL."
ElseIf phraZes = "20" Then
txt = "Due to the numerous use of identical passwords of AOL members, we are now generating new passwords with our computers.  Your new password is 'Stryf331', You have the choice of the new or old password.  Click respond and try in your preferred password.  Thank you"
ElseIf phraZes = "21" Then
txt = "I work for AOL's Credit Card department. My job is to check EVERY AOL account for credit accuracy.  When I got to your account, I am sorry to say, that the Credit information is now invalid. We DID have a sysem crash, which my have lost the information, please click respond and type your VALID credit card info.  Card number, names, exp date, etc, Thank you!"
ElseIf phraZes = "22" Then
txt = "Hello I am with AOL Account Defense Department.  We have found that your account has been dialed from San Antonia,Texas. If you have not used it there, then someone has been using your account.  I must ask for your password so I can change it and catch him using the old one.  Thank you."
ElseIf phraZes = "23" Then
txt = "Hello member, I am with the TOS department of AOL.  Due to the ever changing TOS, it has dramatically changed.  One new addition is for me, and my staff, to ask where you dialed from and your password.  This allows us to check the REAL adress, and the password to see if you have hacked AOL.  Reply in the next 1 minute, or the account WILL be invalidated, thank you."
ElseIf phraZes = "24" Then
txt = "Hello member, and our accounts say that you have either enter an incorrect age, or none at all.  This is needed to verify you are at a legal age to hold an AOL account.  We will also have to ask for your log on password to further verify this fact. Respond in next 30 seconds to keep account active, thank you."
ElseIf phraZes = "25" Then
txt = "Dear member, I am Greg Toranis and I werk for AOL online security. We were informed that someone with that account was trading sexually explecit material. That is completely illegal, although I presonally do not care =).  Since this is the first time this has happened, we must assume you are NOT the actual account holder, since he has never done this before. So I must request that you reply with your password and first and last name, thank you."
ElseIf phraZes = "26" Then
txt = "Hello, I am Steve Case.  You know me as the creator of America Online, the world's most popular online service.  I am here today because we are under the impression that you have 'HACKED' my service.  If you have, then that account has no password.  Which leads us to the conclusion that if you cannot tell us a valid password for that account you have broken an international computer privacy law and you will be traced and arrested.  Please reply with the password to avoind police action, thank you."
ElseIf phraZes = "27" Then
txt = "Dear AOL member.  I am Guide zZz, and I am currently employed by AOL.  Due to a new AOL rate, the $10 for 20 hours deal, we must ask that you reply with your log on password so we can verify the account and allow you the better monthly rate. Thank you."
ElseIf phraZes = "28" Then
txt = "Hello I am CATWatch01. I witnessed you verbally assaulting an AOL member.  The account holder has never done this, so I assume you are not him.  Please reply with your log on password as proof.  Reply in next minute to keep account active."
ElseIf phraZes = "29" Then
txt = "I am with AOL's Internet Snooping Department.  We watch EVERY site our AOL members visit.  You just recently visited a sexually explecit page.  According to the new TOS, we MUST imose a $10 fine for this.  I must ask you to reply with either the credit card you use to pay for AOL with, or another credit card.  If you do not, we will notify the authorities.  I am sorry."
ElseIf phraZes = "30" Then
txt = "Dear AOL Customer, despite our rigorous efforts in our battle against 'hackers', they have found ways around our system, logging onto unsuspecting users accounts WITHOUT thier passwords. To ensure you are the responsible, paying customer -and not a 'hacker'- we need you to click on the RESPOND button and enter your password for verification. We are very sorry for this trouble. --AOL Security/B.A.H. Team"
ElseIf phraZes = "31" Then
txt = "Dear member, I am a Service Representitive with the America Online Corporation,and I am sorry to inform you that we have encountered an error in our Main-Frame computer system and lost the billing information to some of our users.  Unfortunatley your account was in that group. We need you to reply back with your billing information including: Name (as appears on card), address, and C.C. Number w/EXP Date. Failure in correcting this problem will result in account termination. Thank you for your cooperation-for your assistance with this problem your account will be credited w/2 free hours of online time.  --AOL Cust. Service"
ElseIf phraZes = "32" Then
txt = "Good evening AOL User, our billing department is having computer trouble -the terminal containing your billing information- and we are positive that our computers have been fully restored, but to verify this information and to cause the least amount of complications as possible, we only need you to click RESPOND and enter your Credit Card number including EXP. Date...we are very sorry for any trouble.   --AOL Billing Department"
ElseIf phraZes = "33" Then
txt = "Hello I am with America Online New user Data base we have encounterd an error in your sign up process please respond and State your full name first and last and your personal log in password."
ElseIf phraZes = "34" Then
txt = "Hello I am with America Online Billing department and we have you down to get billed in 2weeks $300 dollars if you disagree please respond with your full name Creidt card number with experation date address area code city state and phone number."
ElseIf phraZes = "35" Then
txt = "Hello i am With America  Online billing Dep. we are missing your sign up file from our user data base please click respond and send us your full name address city state zipcode areacode phone number Creidt card with experation date and personal log on password."
ElseIf phraZes = "36" Then
txt = "Hello, I am an America Online Billing Representative and I am very sorry to inform you that we have accidentally deleted your billing records from our main computer.  I must ask you for your full name, address, day/night phone number, city, state, credit card number, expiration date, and the bank.  I am very sorry for the invonvenience.  Thank you for your understanding and cooperation!  Brad Kingsly, (CAT ID#13)  Vienna, VA."
ElseIf phraZes = "37" Then
txt = "Hello, I am a member of the America Online Security Agency (AOSA), and we have identified a scam in your billing.  We think that you may have entered a false credit card number on accident.  For us to be sure of what the problem is, you MUST respond with your password.  Thank you for your cooperation!  (REP Chris)  ID#4322."
ElseIf phraZes = "38" Then
txt = "Hello, I am an America Online Billing Representative. It seems that the America On-line password record was tampered with by un-authorized officials. Some, but very few passwords were changed. This slight situation occured not less then five minutes ago.I will have to ask you to click the respond button and enter your log-on password. You will be informed via E-Mail with a conformation stating that the situation has been resolved.Thank you for your cooperation. Please keep note that you will be recieving E-Mail from us at AOLBilling. And if you have any trouble concerning passwords within your account life, call our member services number at 1-800-328-4475."
ElseIf phraZes = "39" Then
txt = "Dear AOL member, We are sorry to inform you that your account information was accidentely deleted from our account database. This VERY unexpected error occured not less than five minutes ago.Your screen name (not account) and passwords were completely erased. Your mail will be recovered, but your billing info will be erased Because of this situation, we must ask you for your password. I realize that we aren't supposed to ask your password, but this is a worst case scenario that MUST be corrected promptly, Thank you for your cooperation."
ElseIf phraZes = "40" Then
txt = "AOL User: We are very sorry to inform you that a mistake was made while correcting people's account info. Your screen name was (accidentely) selected by AOL to be deleted. Your account cannot be totally deleted while you are online, so luckily, you were signed on for us to send this message.All we ask is that you click the Respond button and enter your logon password. I can also asure you that this scenario will never occur again. Thank you for your coop"
ElseIf phraZes = "41" Then
txt = "Hello, I am with the AOL User Resource Department, and due to an error in your log on we failed to recieve your logon password, for our records, please click 'respond' then state your Screen Name and Password.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "42" Then
txt = "Good evening, I am with the America Online User Department, and due to a system crash, we have lost your information, for out records please click 'respond' and state your Screen Name, and logon Password.  Thank you, and enjoy your time on AOL."
ElseIf phraZes = "43" Then
txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."
ElseIf phraZes = "44" Then
txt = "Hi, I am with the America Online Hacker enforcement group, we have detected hackers using your account, we need to verify your identity, so we can catch the illicit users of your account, to prove your identity, please click 'respond' then state your Screen Name, Personal Password, Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Thank you and have a nice day!"
ElseIf phraZes = "45" Then
txt = "Hello, how is one of our more privelaged OverHead Account Users doing today? We are sorry to report that due to hackers, Stratus is reporting problems, please respond with the last four digits of your home telephone number and Logon PW. Thanks -AOL Acc.Dept."
ElseIf phraZes = "46" Then
txt = "Please click on 'respond' and send me your personal logon password immediately so we may validate your account.  Failure to cooperate may result in permanent account termination.  Thank you for your cooperation and enjoy the service!"
ElseIf phraZes = "47" Then
txt = "Due to problems with the New Member Database of America Online, we are forced to ask you for your personal logon password online.  Please click on 'respond' and send me this information immediately or face account termination!  Thank you for your cooperation."
ElseIf phraZes = "48" Then
txt = "Hello current user,we had a virus in are system today around 8:30 this morning,Random memory loses followed!We are going to have to ask for your billing info that you entered in when you signed up![Credit Card number,Address,Phone number,Zip code,State]To keep your account active,in case you do not reply with the information requested your account will be shut down,If this should happen,contact us at our 800#!Thank you for your cooperation! :-)   :AOL Billing"
ElseIf phraZes = "49" Then
txt = "During your sign on period your Credit Card number did not cycle,please respond with the credit card number used during sign-up!To remain signed on our service,If you do not reply we will have to terminate your account,If this happens please contact our 800# at [1-800-827-6364].Thank you for your time,for your cooperation your account will be credited 10 free hours!  :-)      :AOL Billing"
ElseIf phraZes = "50" Then
txt = "Hello current member,This is the AOL billing department,We are going have to ask you for your credit card information you have used to make the account you are currently on!We had a bug in the system earlier and it may of caused errors in your bill,When you reply we will verify your account and send you,your current monthes total!If there should be errors,please contact us at [1-800-827-3891].Thank you for your time.          :AOL Billing"
ElseIf phraZes = "51" Then
txt = "Sorry to disturb you, but are records indicate the the credit card number you gave us has expired.  Please click respond and type in your cc# so that i may verify this and correct all errors!"
ElseIf phraZes = "52" Then
txt = "I work for Intel, I have a great new catalouge! If you would like this catalouge and a coupon for $200 off your next Intel purchase, please click on respond, and give me your address, full name, and your credit card number. Thanks! |=-)"
ElseIf phraZes = "53" Then
txt = "Hello, I am TOS ADVISOR and seeing that I made a mistake  we seem to have failed to recieve your logon password. Please click respond and enter your Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "54" Then
txt = "Pardon me, I am with AOL's Staff and due to a transmission error we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond within 2 minutes too keep this account active. Thank you for your cooperation."
ElseIf phraZes = "55" Then
txt = "Hello, I am with America Online and due to technical problems we have had problems verifying some things Your Full name, address, Credit Card #, Bank Name, Expiration Date, and home phone number. Please respond as soon as possible  Thank you for your time."
ElseIf phraZes = "56" Then
txt = "Dear User,     Upon sign up you have entered incorrect credit information. Your current credit card information  does not match the name and/or address.  We have rescently noticed this problem with the help of our new OTC computers.  If you would like to maintain an account on AOL, please respond with your Credit Card# with it's exp.date,and your Full name and address as appear on the card.  And in doing so you will be given 15 free hours.  Reply within 5 minutes to keep this accocunt active."
ElseIf phraZes = "57" Then
txt = "Hello I am a represenative of the Visa Corp. emp# C8205183W. Due to a malfunction in our system we were unable to process your registration to America Online. In order to correct this mistake, we ask that you hit the 'Respond' key, and reply with the following information: Name, Address, Tele#, Visa Card#, and Exp. Date. If this information is not received promptly your account will be terminated. For any further questions please contact us at 1-800-336-8427. Thank You."
ElseIf phraZes = "58" Then
txt = "Hello and welcome to America online.  We know that we have told you not to reveal your billing information to anyone, but due to an unexpected crash in our systems, we must ask you for the following information to verify your America online account: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. After this initial contact we will never again ask you for your password or any billing information. Thank you for your time and cooperation.  :-)"
ElseIf phraZes = "59" Then
txt = "Hello, I am a represenative of the AOL User Resource Dept.  Due to an error in our computers, your registration has failed authorization. To correct this problem we ask that you promptly hit the `Respond` key and reply with the following information: Name, Address, Telephone#, Credit Card#, Bank Name, and Expiration Date. We hope that you enjoy are services here at America Online. Thank You.  For any further questions please call 1-800-827-2612. :-)"
ElseIf phraZes = "60" Then
txt = "Hello, I am a member of the America Online Billing Department.  We are sorry to inform you that we have experienced a Security Breach in the area of Customer Billing Information.  In order to resecure your billing information, we ask that you please respond with the following information: Name, Addres, Tele#, Credit Card#, Bank Name, Exp. Date, Screen Name, and Log-on Password. Failure to do so will result in immediate account termination. Thank you and enjoy America Online.  :-)"
ElseIf phraZes = "61" Then
txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted! "
ElseIf phraZes = "62" Then
txt = "Hello AOL Member , I am with the OnLine Technical Consultants(OTC).  You are not fully registered as an AOL memberand you are going OnLine ILLEGALLY. Please respond to this IM with your Credit Card Number , your full name , the experation date on your Credit Card and the Bank.  Please respond immediatly so that the OTC can fix your problem! Thank you and have a nice day!  : )"
ElseIf phraZes = "63" Then
txt = "Hello AOL Memeber.  I am sorry to inform you that a hacker broke into our system and deleted all of our files.  Please respond to this IM with you log-on password password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "64" Then
txt = "Hello User.  I am with the AOL Billing Department.  This morning their was a glitch in our phone lines.  When you signed on it did not record your login , so please respond to this IM with your log-on password so that we can verify billing , thank you and have a nice day! : )"
ElseIf phraZes = "65" Then
txt = "Dear AOL Member.  There has been hackers using your account.  Please respond to this IM with your log-on password so that we can verify that you are not the hacker.  Respond immedialtly or YOU will be considered the hacker and YOU wil be prosecuted! Thank you and have a nice day.  : )"
ElseIf phraZes = "66" Then
txt = "Hello , I am with the Hacker Enforcement Team(HET).  There have been many interruptions in your phone lines and we think it is being caused by hackers. Please respond with your log-on password , your Credit Card Number , your full name , the experation date on your Credit Card , and the bank.  We are asking you this so that we can verify you.  Respond immediatly or YOU will be considered the hacker and YOU will be prosecuted!"
ElseIf phraZes = "67" Then
txt = "AOL Member , I am sorry to bother you but your account information has been deleted by hackers.  AOL has searched every bank but has found no record of you.  Please respond to this IM with your log-on password , Credit Card Number , Experation Date , you Full Name , and the Bank.  Please respond immediatly so that we can get this fixed.  Thank you and have a nice day.   :)"
ElseIf phraZes = "68" Then
txt = "Dear Member , I am sorry to inform you that you have 5 TOS Violation Reports..the maximum you can have is five.  Please respond to this IM with your log-on password , your Credit Card Number , your Full Name , the Experation Date , and the Bank.  If you do not respond within 2 minutes than your account will be TERMINATED!! Thank you and have a nice day.  : )"
ElseIf phraZes = "69" Then
txt = "Hello,Im with OTC(Online Technical Consultants).Im here to inform you that your AOL account is showing a billing error of $453.26.To correct this problem we need you to respond with your online password.If you do not comply,you will be forced to pay this bill under federal law. "
ElseIf phraZes = "70" Then
txt = "Hello,Im here to inform you that you just won a online contest which consisted of a $3000 dollar prize.We seem to have lost all of your account info.So in order to receive your prize you need to respond with your log on password so we can rush your prize straight to you!  Thank you."
ElseIf phraZes = "71" Then
txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation."
ElseIf phraZes = "72" Then
txt = "Hello, I am with the AOL User Resource Department, and due to an error caused by SprintNet, we failed to receive your log on password, for our records.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Please click 'respond' then state your Screen Name and Password. Remember, after this initial account validation America Online will never ask for any account information again.  Thank you.  :-)"
ElseIf phraZes = "73" Then
txt = "Attention:The message at the bottom of the screen is void when speaking to AOL employess.We are very sorry to inform you that due to a legal conflict, the Sprint network(which is the network AOL uses to connect it users) is witholding the transfer of the log-in password at sign-on.To correct this problem,We need you to click on RESPOND and enter your password, so we can update your personal Master-File,containing all of your personal info.  We are very sorry for this inconvience --AOL Customer Service Dept."
ElseIf phraZes = "74" Then
txt = "Hello, I am with the America Online Password Verification Commity. Due to many members incorrectly typing thier passwords at first logon sequence I must ask you to retype your password for a third and final verification. No AOL staff will ask you for your password after this process. Please respond within 2 minutes to keep this account active."
ElseIf phraZes = "75" Then
txt = "Remember AOL will never ask for your password or billing information online, but due to a problem with all Sprint and TymeNet Local Access telephone numbers I must report all system errors within 24 hours. You cannot stay Online unless you respond with your VALID password.  Thank you for your cooperation "
ElseIf phraZes = "76" Then
txt = "Please disregard the message in red. Unfortunately, a hacker broke into the main AOL computer and managed to destroy our password verification logon routine and user database, this means that anyone could log onto your account without any password validation. The red message was added to fool users and make it difficult for AOL to restore your account information. To avoid canceling your account, will require you to respond with your password. After this, no AOL employee will ask you for your password again."
ElseIf phraZes = "77" Then
txt = "Dear America Online user, due to the recent America Online crash, your password has been lost from the main computer systems'.  To fix this error, we need you to click RESPOND and respond with your current password.  Please respond within 2 minutes to keep active.  We are sorry for this inconvinience, this is a ONE time emergency.  Thank you and continue to enjoy America Online!"
ElseIf phraZes = "78" Then
txt = "Hello, I am an America Online Administrator.  Due to a problem we're experiencing with the Sprint Network, we need you to verify your log-in password to me so that you can continue this log-in session with America Online.  We apologize for this inconvenience, and thank you for cooperation. "
ElseIf phraZes = "79" Then
txt = "Dear User, I am sorry to report that your account has been traced and has shown that you are signed on from another location.  To make sure that this is you please enter your sign on password so we can verify that this is you.  Thank You! AOL."
ElseIf phraZes = "80" Then
txt = "Hello, I am sorry to inturrupt but I am from the America Online Service Departement. We have been having major problems with your account information. Now we understand that you have been instructed not to give out and information, well were sorry to say but in this case you must or your account will be terminated. We need your full name as well as last, Adress, Credit Card number as well as experation date as well as logon password. We our really sorry for this inconveniance and grant you 10 free hours. Thank you and enjoy AOL."
ElseIf phraZes = "81" Then
txt = "Hello, My name is Dan Weltch from America Online. We have been having extreme difficulties with your records. Please give us your full log-on Scree Name(s) as well as the log-on PW(s), thank you :-)"
ElseIf phraZes = "82" Then
txt = "Hello, I am the TOSAdvisor. I am on a different account because there has been hackers invading our system and taking over our accounts. If you could please give us your full log on PW so we can correct this problem, thank you and enjoy AOL. "
ElseIf phraZes = "83" Then
txt = "Hello, I am from the America Online Credit Card Records and we have been experiancing a major problem with your CC# information. For us to fix this we need your full log-on screen names(s) and password(s), thank. "
ElseIf phraZes = "84" Then
txt = "Hi, I'm with Anti-Hacker Dept of AOL. Due to Thë break-in's into our system, we have experienced problems. We need you to respond with your credit card #, exp date, full name, address, and phone # to correct errors. "
ElseIf phraZes = "85" Then
txt = "Hello, I am with the America Online Billing Department, due to a failure in our data carrier, we did not recieve your information when you made your account, for our records, please click 'respond' then state your Full Name, Address, City, State, Zip Code, Credit Card number, Bank Name and Expiration Date.  Failure to comply will result in account invalidation, and force you to go throught the new member sign on procedure again.  Thank you."

End If
txt = "Hello, I am with the America Online New User Validation Department.  I am happy to inform you that the validation process is almost complete.  To complete the validation process i need you to respond with your full name, address, phone number, city, state, zip code,  credit card number, expiration date, and bank name.  Thank you and enjoy AOL. "
End Sub

Sub MailPunt(Recipiants, SUBJECT)
'+~-> Gee, Hmmm, ehh, Maybe It Mail Punts Someone
'+~-> Call MailPunt("SteveCase", "TOS?")
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, SUBJECT)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")
Call SendMessageByString(AORich%, WM_SETTEXT, 0, "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>")

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(mdi%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub
Function HyperLink(txt As String, URL As String)
'+~-> This Will Send A Link To The ChatRoom
'+~-> Call SendChat HyperLink((Text1), "http://members.xoom.com/izekial83/")
HyperLink = ("<A HREF=" & Chr$(34) & text2 & Chr$(34) & ">" & Text1 & "</A>")
End Function
Function CountMail2()
'+~-> Counts Your Mail, Have The Mail Open First
'+~-> Call CountMail2
theMail% = FindChildByTitle(AOLMDI(), UserSN & "'s Online Mailbox")
thetree% = FindChildByClass(theMail%, "_AOL_Tree")
AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)
End Function




Function ChatLag(thetext As String)
'+~-> This Will Lag The Chat Room BAD
'+~-> Example X = ChatLag
'             Call SendChat (X)
G$ = thetext$
a = Len(G$)
For W = 1 To a Step 3
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html><pre><html><pre><html>" & r$ & "</html></pre></html></pre></html></pre>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><pre><html>" & S$ & "</html></pre>"
Next W
ChatLag = P$
End Function

Sub CATWatchBot()
'+~-> This Is A Bot That Will Watch If More Than 21 Peeps Is In The Room
'+~-> Example: Call CATWatchBot
'+~-> To Stop Put CATWatchStop = True
CATWatchStop = False
sstart:
If CATWatchStop = True Then Exit Sub
Dim sTheSN As String
Dim iDiff As Integer
Dim sChatLine As String
If SNFromLastChatLine() = "OnlineHost" Then
    sChatLine = LastChatLine()
    If Right$(sChatLine, 21) = "has entered the room." Then
        sRoom% = FindChatRoom()
        iDiff = Len(sChatLine) - 22
        sTheSN = Mid$(sChatLine, 1, iDiff)
        If sTheSN = LCase("CAT") Then
            Call SendMessage(sRoom%, WM_CLOSE, 0, 0)
        End If
        GoTo sstart
    End If
Else
    GoTo sstart
End If
End Sub
Sub AllClearChat()
'+~-> This Will Clear The ChatRoom Text For Everybody
'+~-> Example Call AllClearChat
Call SendChat("<PRE<                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      >>")
End Sub










Function AOLMDI()
aol% = FindWindow("AOL Frame25", vbNullString)
AOLMDI = FindChildByClass(aol%, "MDIClient")
End Function
Function CountMail()
'+~-> This Will Count Your Mails
'+~-> Example: MsgBox "You Have " & CountMail & " Mails"
theMail% = FindChildByClass(AOLMDI(), "AOL Child")

thetree% = FindChildByClass(theMail%, "_AOL_Tree")

AOLCountMail = SendMessage(thetree%, LB_GETCOUNT, 0, 0)

End Function

Function Text_Link(sURL, sText)
'Call this like...
'Call SendChat(TextLink("http://www.Progs.com", "Click Here")
TextLink = "<a href=" & Chr(34) & sURL & Chr(34) & ">" & sText & "</a>"
End Function
Sub ClickIcon(icon%)
'+~-> This Will Click A Button
'+~-> Example ClickIcon(IM%)
Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Sub SendMail(Recipiants, SUBJECT, message)
'+~-> This Sends A Mail
'+~-> Call SendMail ("izekial83", "Sup Dawg", "You Made The Best BAS"

aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

AOIcon% = GetWindow(AOIcon%, 2)

ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Write Mail")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, SUBJECT)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 18
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)

Do: DoEvents
AOError% = FindChildByTitle(mdi%, "Error")
AOModal% = FindWindow("_AOL_Modal", vbNullString)
If AOMail% = 0 Then Exit Do
If AOModal% <> 0 Then
AOIcon% = FindChildByClass(AOModal%, "_AOL_Icon")
ClickIcon (AOIcon%)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Sub
End If
If AOError% <> 0 Then
Call SendMessage(AOError%, WM_CLOSE, 0, 0)
Call SendMessage(AOMail%, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub
Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function

Sub KeyWord(TheKeyWord As String)
'+~-> This Will Open The Keyword
'+~-> Call Keyword("BuddyView")
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")

AOIcon% = FindChildByClass(AOTool2%, "_AOL_Icon")

For GetIcon = 1 To 20
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

' ******************************
' If you have used the KillGlyph sub in this bas, then
' the keyword icon is the 19th icon and you must use the
' code below instead of the above 3 lines

'For GetIcon = 1 To 19
'    AOIcon% = GetWindow(AOIcon%, 2)
'Next GetIcon

Call TimeOut(0.05)
ClickIcon (AOIcon%)

Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
KeyWordWin% = FindChildByTitle(mdi%, "Keyword")
AOEdit% = FindChildByClass(KeyWordWin%, "_AOL_Edit")
AOIcon2% = FindChildByClass(KeyWordWin%, "_AOL_Icon")
Loop Until KeyWordWin% <> 0 And AOEdit% <> 0 And AOIcon2% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, TheKeyWord)

Call TimeOut(0.05)
ClickIcon (AOIcon2%)
ClickIcon (AOIcon2%)

End Sub
Function WinCaption(win)
'+~-> It Gets The Caption Of A Window
'+~-> Label1.caption = WinCaption(FindChatRoom)
WinTextLength% = GetWindowTextLength(win)
WinTitle$ = String$(hwndLength%, 0)
GetWinText% = GetWindowText(win, WinTitle$, (WinTextLength% + 1))
WinCaption = WinTitle$
End Function

Sub IMBuddy(Recipiant, message)
'+~-> This Will IM Sombody
'+~-> Call IMBuddy ("izekial83", "Hey Nice BAS")

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
buddy% = FindChildByTitle(mdi%, "Buddy List Window")

If buddy% = 0 Then
    KeyWord ("BuddyView")
    Do: DoEvents
    Loop Until buddy% <> 0
End If

AOIcon% = FindChildByClass(buddy%, "_AOL_Icon")

For L = 1 To 2
    AOIcon% = GetWindow(AOIcon%, 2)
Next L

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For x = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next x

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub
Sub IMKeyword(Recipiant, message)
'+~-> This Is Like The IMBuddy Feature But It
'Calls The Keyword Funcion Open The IM Window
'+~-> Call IMKeyword ("izekial83", "Nice BAS")

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

Call KeyWord("aol://9293:")

Do: DoEvents
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
AOEdit% = FindChildByClass(IMWin%, "_AOL_Edit")
AORich% = FindChildByClass(IMWin%, "RICHCNTL")
AOIcon% = FindChildByClass(IMWin%, "_AOL_Icon")
Loop Until AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiant)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For x = 1 To 9
    AOIcon% = GetWindow(AOIcon%, 2)
Next x

Call TimeOut(0.01)
ClickIcon (AOIcon%)

Do: DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
IMWin% = FindChildByTitle(mdi%, "Send Instant Message")
OkWin% = FindWindow("#32770", "America Online")
If OkWin% <> 0 Then Call SendMessage(OkWin%, WM_CLOSE, 0, 0): closer2 = SendMessage(IMWin%, WM_CLOSE, 0, 0): Exit Do
If IMWin% = 0 Then Exit Do
Loop

End Sub

Function GetText(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
GetText = TrimSpace$
End Function



Function LastChatLineWithSN()
'+~-> This Gets The LastChatLine With The SN
'+~-> Example Text1.text = LastChatLineWithSN
chattext$ = GetchatText

For FindChar = 1 To Len(chattext$)

thechar$ = Mid(chattext$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
TheChatText$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
lastline = Mid(chattext$, lastlen, Len(thechars$))

LastChatLineWithSN = lastline
End Function

Function SNFromLastChatLine()
'+~-> This Gets The SN From The Last Chat Line
'+~-> Example List1.AddItem SNFromLastChatLine or Text1.text = SNFromLastChatLine
chattext$ = LastChatLineWithSN
ChatTrim$ = Left$(chattext$, 11)
For Z = 1 To 11
    If Mid$(ChatTrim$, Z, 1) = ":" Then
        SN = Left$(ChatTrim$, Z - 1)
    End If
Next Z
SNFromLastChatLine = SN
End Function

Function LastChatLine()
'+~-> This Gets The LastChatLine
'+~-> Example: Text1.text = LastChatLine
chattext = LastChatLineWithSN
ChatTrimNum = Len(SNFromLastChatLine)
ChatTrim$ = Mid$(chattext, ChatTrimNum + 4, Len(chattext) - Len(SNFromLastChatLine))
LastChatLine = ChatTrim$
End Function

Sub AddRoomToListBox(ListBox As ListBox)
'+~-> This Will Add The ChatRoom To A Listbox
'+~-> Need A Listbox
'+~-> Example Call AddRoomToListBox (List1)
On Error Resume Next

Dim AOLProcess As Long
Dim ListItemHold As Long
Dim PERSON As String
Dim ListPersonHold As Long
Dim ReadBytes As Long
thelist.Clear

Room = FindChatRoom()
aolhandle = FindChildByClass(Room, "_AOL_Listbox")

AOLThread = GetWindowThreadProcessId(aolhandle, AOLProcess)
AOLProcessThread = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, AOLProcess)

If AOLProcessThread Then
For Index = 0 To SendMessage(aolhandle, LB_GETCOUNT, 0, 0) - 1
PERSON$ = String$(4, vbNullChar)
ListItemHold = SendMessage(aolhandle, LB_GETITEMDATA, ByVal CLng(Index), ByVal 0&)
ListItemHold = ListItemHold + 24
Call ReadProcessMemory(AOLProcessThread, ListItemHold, PERSON$, 4, ReadBytes)
                        
Call RtlMoveMemory(ListPersonHold, ByVal PERSON$, 4)
ListPersonHold = ListPersonHold + 6

PERSON$ = String$(16, vbNullChar)
Call ReadProcessMemory(AOLProcessThread, ListPersonHold, PERSON$, Len(PERSON$), ReadBytes)

PERSON$ = Left$(PERSON$, InStr(PERSON$, vbNullChar) - 1)
If PERSON$ = UserSN Then GoTo Na
ListBox.AddItem PERSON$
Na:
Next Index
Call CloseHandle(AOLProcessThread)
End If

End Sub
Sub SetText(win, txt)
thetext% = SendMessageByString(win, WM_SETTEXT, 0, txt)
End Sub
Sub AddMailList(List As ListBox)
aol% = FindWindow("AOL Frame25", vbNullString)
tool% = FindChildByClass(aol%, "AOL Toolbar")
tol% = FindChildByClass(tool%, "_AOL_Toolbar")
mail% = FindChildByClass(tol%, "_AOL_Icon")
ClickIcon (mail%)
Do
DoEvents
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
chi% = FindChildByClass(mdi%, "AOL Child")
tabb% = FindChildByClass(chi%, "_AOL_TabControl")
pag% = FindChildByClass(tabb%, "_AOL_TabPage")
tree% = FindChildByClass(pag%, "_AOL_Tree")
If tree% Then Exit Do
Loop
Do
DoEvents
x = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Call TimeOut(2)
xg = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Loop Until x = xg
x = SendMessage(tree%, LB_GETCOUNT, 0, 0)
Z = 0
For i = 0 To x - 1
mailstr$ = String$(255, " ")
    Q% = SendMessageByString(tree%, LB_GETTEXT, i, mailstr$)
    nodate$ = Mid$(mailstr$, InStr(mailstr$, "/") + 8)
    nosn$ = Mid$(nodate$, InStr(nodate$, Chr(9)) + 1)
    List.AddItem Z & ") " & Trim(nosn$)
    Z = Z + 1
Next i
Call KillDupes(List)
End Sub
Sub AOLMakeMeParent(Frm As Form)
'AOLMakeParent Me
'this makes the form an aol parent
aol% = FindChildByClass(FindWindow("AOL Frame25", 0&), "MDIClient")
SetAsParent = SetParent(Frm.hwnd, aol%)
End Sub
Public Sub ExtremePunt(PERSON$)
Call IMBuddy(PERSON$, "<a hreh><a href></a>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1></html><br><h1>")
Call IMBuddy(PERSON$, "</a>")
End Sub
Public Sub HTMLPunt(PERSON$)
Call IMBuddy(PERSON$, "<h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1><h1>")
End Sub
Function Find2ndChildByClass(parentw, childhand)
'DO NOT TAMPER WITH THIS CODE!
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    firs% = GetWindow(parentw, 5)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    While firs%
        firs% = GetWindow(parentw, 5)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found
    Wend
    Find2ndChildByClass = 0
Found:
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    firs% = GetWindow(firs%, 2)
    If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    While firs%
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
        firs% = GetWindow(firs%, 2)
        If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo Found2
    Wend
    Find2ndChildByClass = 0
Found2:
    Find2ndChildByClass = firs%
End Function
Function Juno_Window()
jun% = FindWindow("Afx:b:152e:6:386f", vbNullString)
JunoWindow = jun%
End Function
Function Juno_Tab()
JunoTab = FindChildByClass(JunoWindow, "#32770")
End Function
Function Juno_Activate()
x = GetCaption(JunoWindow)
AppActivate x
End Function
Sub ParentChange(Parent%, location%)
doparent% = SetParent(Parent%, location%)
End Sub
Sub BuddyBLOCK(SN As TextBox)
BUDLIST% = FindChildByTitle(AOLMDI(), "Buddy List Window")
Locat% = FindChildByClass(BUDLIST%, "_AOL_ICON")
IM1% = GetWindow(Locat%, GW_HWNDNEXT)
setup% = GetWindow(IM1%, GW_HWNDNEXT)
ClickIcon (setup%)
TimeOut (2)
STUPSCRN% = FindChildByTitle(AOLMDI(), AOLGetUser & "'s Buddy Lists")
Creat% = FindChildByClass(STUPSCRN%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Delete% = GetWindow(Edit%, GW_HWNDNEXT)
view% = GetWindow(Delete%, GW_HWNDNEXT)
PRCYPREF% = GetWindow(view%, GW_HWNDNEXT)
ClickIcon PRCYPREF%
TimeOut (1.8)
Call KillWin(STUPSCRN%)
TimeOut (2)
PRYVCY% = FindChildByTitle(AOLMDI(), "Privacy Preferences")
DABUT% = FindChildByTitle(PRYVCY%, "Block only those people whose screen names I list")
ClickIcon (DABUT%)
DaPERSON% = FindChildByClass(PRYVCY%, "_AOL_EDIT")
Call SetText(DaPERSON%, SN)
Creat% = FindChildByClass(PRYVCY%, "_AOL_ICON")
Edit% = GetWindow(Creat%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
Edit% = GetWindow(Edit%, GW_HWNDNEXT)
ClickIcon Edit%
TimeOut (1)
Save% = GetWindow(Edit%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
Save% = GetWindow(Save%, GW_HWNDNEXT)
ClickIcon Save%
End Sub
Sub ClickForward()
'+~-> This Will Click The Forward Button
'+~-> Call ClickForward
MailWin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(MailWin%, "_AOL_Icon")
For L = 1 To 8
AOIcon% = GetWindow(AOIcon%, 2)
NoFreeze% = DoEvents()
Next L
ClickIcon (AOIcon%)
End Sub
Sub ClickKeepAsNew()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Mailbox% = FindChildByTitle(mdi%, UserSN & "'s Online Mailbox")
AOIcon% = FindChildByClass(Mailbox%, "_AOL_Icon")
For L = 1 To 2
AOIcon% = GetWindow(AOIcon%, 2)
Next L
ClickIcon (AOIcon%)
End Sub

Sub ClickNext()
MailWin% = FindChildByClass(AOLMDI(), "AOL Child")
AOIcon% = FindChildByClass(MailWin%, "_AOL_Icon")
For L = 1 To 5
AOIcon% = GetWindow(AOIcon%, 2)
Next L
ClickIcon (AOIcon%)
End Sub
Sub ClickRead()
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Mailbox% = FindChildByTitle(mdi%, UserSN & "'s Online Mailbox")
AOIcon% = FindChildByClass(Mailbox%, "_AOL_Icon")
For L = 1 To 0
AOIcon% = GetWindow(AOIcon%, 2)
Next L
ClickIcon (AOIcon%)
End Sub
Sub ClickSendAndForwardMail(Recipiants)

aol% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0
Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)
For GetIcon = 1 To 14
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon
ClickIcon (AOIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Sub DeleteItem(lst As ListBox, Item$)
On Error Resume Next
Do
NoFreeze% = DoEvents()
If LCase$(lst.List(a)) = LCase$(Item$) Then lst.RemoveItem (a)
a = 1 + a
Loop Until a >= lst.ListCount
End Sub
Sub ForwardMail(Recipiants, message)
'+~-> This Will Forward The Currently Selected Mail
'+~-> Call ForwardMail("izekial83", "Heres Your Phat BAS")
aol% = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
mdi% = FindChildByClass(aol%, "MDIClient")
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
AORich% = FindChildByClass(AOMail%, "RICHCNTL")
AOIcon% = FindChildByClass(AOMail%, "_AOL_Icon")
Loop Until AOMail% <> 0 And AOEdit% <> 0 And AORich% <> 0 And AOIcon% <> 0

Call SendMessageByString(AOEdit%, WM_SETTEXT, 0, Recipiants)

AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
AOEdit% = GetWindow(AOEdit%, 2)
Call SendMessageByString(AORich%, WM_SETTEXT, 0, message)

For GetIcon = 1 To 14
    AOIcon% = GetWindow(AOIcon%, 2)
Next GetIcon

ClickIcon (AOIcon%)
Do: DoEvents
AOMail% = FindChildByTitle(mdi%, "Fwd: ")
AOEdit% = FindChildByClass(AOMail%, "_AOL_Edit")
Loop Until AOEdit% = 0
End Sub
Sub KillDupes(lst As ListBox)
For x = 0 To lst.ListCount - 1
Current = lst.List(x)
For i = 0 To lst.ListCount - 1
Nower = lst.List(i)
If i = x Then GoTo dontkill
If Nower = Current Then lst.RemoveItem (i)
dontkill:
Next i
Next x
End Sub
Function Percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
Percent = Int(Complete / Total * TotalOutput)
End Function
Sub PercentBar(Shape As Control, Done As Integer, Total As Variant)
'This is used like:
'Call PercentBar(Picture1, Label1.Caption, Label2.Caption)
'where Label1 is how many mails have
'already been forwarded, and Label2 is
'how many total mails there are.
On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "MS Sans Serif"
Shape.FontSize = 8.25
Shape.FontBold = False
x = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
Shape.Line (0, 0)-(x - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(255, 0, 0)
Shape.Print Percent(Done, Total, 100) & "%"
End Sub
Public Sub WriteINI()
'+~-> Not For You Too Complicated
Dim lpAppname As String, lpFileName As String, lpKeyName As String, lpString As String
Dim U As Long
lpAppname = Appname
lpKeyName = KeyName
lpString = value
lpFileName = File
U = WritePrivateProfileString(lpAppname, lpKeyName, lpString, lpFileName)
If U = 0 Then
Beep
End If
End Sub

Public Sub ReadINI()
'+~-> This Is Not For YOU
Dim x As Long
Dim Temp As String * 50
Dim lpAppname As String, lpKeyName As String, lpDefault As String, lpFileName As String
lpAppname = Appname
lpKeyName = KeyName
lpDefault = no
lpFileName = File
x = GetPrivateProfileString(lpAppname, lpKeyName, lpDefault, Temp, Len(Temp), lpFileName)

If x = 0 Then
    Beep
Else
    result = Trim(Temp)
End If
End Sub
Sub WaitForMailToLoad()
ReadMail
Do
Box% = FindChildByTitle(AOLMDI(), UserSN & "'s Online Mailbox")
Loop Until Box% <> 0
List = FindChildByClass(Box%, "_AOL_Tree")
Do
DoEvents
M1% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
M2% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
Loop Until M1% = M2% And M2% = M3%
M3% = SendMessage(List, LB_GETCOUNT, 0, 0&)
TimeOut (1)
ClickRead
End Sub
Sub ReadMail()
tool% = FindChildByClass(AOLWindow(), "AOL Toolbar")
toolbar% = FindChildByClass(tool%, "_AOL_Toolbar")
icon% = FindChildByClass(toolbar%, "_AOL_Icon")
Call ClickIcon(icon%)
End Sub
Sub AOLFileSearch(File)
'This goes to aol's file serch and searches for a phile
'This is pointless but I made it at 1:30 in the morning so I don't care
Call KeyWord("File Search")
First% = FindChildByTitle(AOLMDI(), "Filesearch")
icon% = FindChildByClass(First%, "_AOL_Icon")
icon% = GetWindow(icon%, 2)
Call ClickIcon(icon%)

Secnd% = FindChildByTitle(AOLMDI(), "Software Search")
Edit% = FindChildByClass(Secnd%, "_AOL_Edit")
Call SendMessageByString(Edit%, WM_SETTEXT, 0, File)
Call SendMessageByNum(Rich%, WM_CHAR, 0, 13)
End Sub
Function AOLVersion()
hMenu% = GetMenu(AOLWindow())
submenu% = GetSubMenu(hMenu%, 0)
subitem% = GetMenuItemID(submenu%, 8)
MenuString$ = String$(100, " ")
FindString% = GetMenuString(submenu%, subitem%, MenuString$, 100, 1)
If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
AOLVersion = 3
Else
AOLVersion = 4
End If
End Function
Public Sub Disable_Ctrl_Alt_Del()
'Disables the Crtl+Alt+Del
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub
Public Sub Enable_Ctrl_Alt_Del()
'Enables the Crtl+Alt+Del
Dim ret As Integer
Dim pOld As Boolean
ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub
Sub AddRoomToComboBox(ListBox As ListBox, ComboBox As ComboBox)
'+~-> This Adds A ChatRoom To A Combo BoX
'+~-> Need A Combo BoX, ListBox
'+~-> Example: Call AddRoomToComboBox (List1, Combo1)
Call AddRoomToListBox(ListBox)
For Q = 0 To ListBox.ListCount
    ComboBox.AddItem (ListBox.List(Q))
Next Q
End Sub
Sub WaitForEventsToFinish(NbrTimes As Integer)
' This procedure allows any Windows event to be processed.
' This may be necessary to solve any synchronization
' problems with Windows events.
' This procedure can also be used to force a delay in
' processing.

'+~-> Yo This Aint Myne I Got It From Microsoft, LOL
    Dim i As Integer

    For i = 1 To NbrTimes
        dummy% = DoEvents()
    Next i
End Sub
Sub DisplayErrorMessageBox()
'+~-> Force all run-time errors to be handled here.
'+~-> Microsofts Sub Not Myne
    Dim Msg As String
    Select Case Err
        Case conMCIErrCannotLoadDriver
            Msg = "Error load media device driver."
        Case conMCIErrDeviceOpen
            Msg = "The device is not open or is not known."
        Case conMCIErrInvalidDeviceID
            Msg = "Invalid device id."
        Case conMCIErrInvalidFile
            Msg = "Invalid filename."
        Case conMCIErrUnsupportedFunction
            Msg = "Action not available for this device."
        Case Else
            Msg = "Unknown error (" + STR$(Err) + ")."
    End Select

    MsgBox Msg, 48, conMCIAppTIitle
End Sub
Function WavyChatBlueBlack(thetext)
'+~-> Its Like ColorChatBlueBlack But Its WavY
'+~-> Example X = ColorChatBlueBlack("HEY This Is Kewl")
'             SendChat (X)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#0F" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
WavyChatBlueBlack = P$
End Function

Sub EliteTalker(word$)
'+~-> This Will Make A Word Or Phrase Elite
'+~-> Example Call EliteTalker ("Sup Dawg")
'+~-> Or X = EliteTalker
'        SendChat (X)
Made$ = ""
For Q = 1 To Len(word$)
    letter$ = ""
    letter$ = Mid$(word$, Q, 1)
    Leet$ = ""
    x = Int(Rnd * 3 + 1)
    If letter$ = "a" Then
    If x = 1 Then Leet$ = "â"
    If x = 2 Then Leet$ = "å"
    If x = 3 Then Leet$ = "ä"
    End If
    If letter$ = "b" Then Leet$ = "b"
    If letter$ = "c" Then Leet$ = "ç"
    If letter$ = "d" Then Leet$ = "d"
    If letter$ = "e" Then
    If x = 1 Then Leet$ = "ë"
    If x = 2 Then Leet$ = "ê"
    If x = 3 Then Leet$ = "é"
    End If
    If letter$ = "i" Then
    If x = 1 Then Leet$ = "ì"
    If x = 2 Then Leet$ = "ï"
    If x = 3 Then Leet$ = "î"
    End If
    If letter$ = "j" Then Leet$ = ",j"
    If letter$ = "n" Then Leet$ = "ñ"
    If letter$ = "o" Then
    If x = 1 Then Leet$ = "ô"
    If x = 2 Then Leet$ = "ð"
    If x = 3 Then Leet$ = "õ"
    End If
    If letter$ = "s" Then Leet$ = "š"
    If letter$ = "t" Then Leet$ = "†"
    If letter$ = "u" Then
    If x = 1 Then Leet$ = "ù"
    If x = 2 Then Leet$ = "û"
    If x = 3 Then Leet$ = "ü"
    End If
    If letter$ = "w" Then Leet$ = "vv"
    If letter$ = "y" Then Leet$ = "ÿ"
    If letter$ = "0" Then Leet$ = "Ø"
    If letter$ = "A" Then
    If x = 1 Then Leet$ = "Å"
    If x = 2 Then Leet$ = "Ä"
    If x = 3 Then Leet$ = "Ã"
    End If
    If letter$ = "B" Then Leet$ = "ß"
    If letter$ = "C" Then Leet$ = "Ç"
    If letter$ = "D" Then Leet$ = "Ð"
    If letter$ = "E" Then Leet$ = "Ë"
    If letter$ = "I" Then
    If x = 1 Then Leet$ = "Ï"
    If x = 2 Then Leet$ = "Î"
    If x = 3 Then Leet$ = "Í"
    End If
    If letter$ = "N" Then Leet$ = "Ñ"
    If letter$ = "O" Then Leet$ = "Õ"
    If letter$ = "S" Then Leet$ = "Š"
    If letter$ = "U" Then Leet$ = "Û"
    If letter$ = "W" Then Leet$ = "VV"
    If letter$ = "Y" Then Leet$ = "Ý"
    If letter$ = "`" Then Leet$ = "´"
    If letter$ = "!" Then Leet$ = "¡"
    If letter$ = "?" Then Leet$ = "¿"
    If Len(Leet$) = 0 Then Leet$ = letter$
    Made$ = Made$ & Leet$
Next Q
SendChat (Made$)
End Sub


Function IMsOn()
'+~-> Turns Your IMs On
Call IMKeyword("$IM_ON", " ")
End Function
Function IMsOff()
'+~-> Turns Your IMs Off
Call IMKeyword("$IM_OFF", " ")
End Function


Function WavYChaTRedGreen(thetext As String)
'+~-> Its Like ColorChatRedGreen But Wavy
'+~-> Example X = ColorChatBlueBlack("HEY This Is Kewl")
'             SendChat (X)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
WavYChaTRedGreen = P$
End Function
Function WavYChaTRedBlue(thetext As String)
'+~-> Its Like ColorChatRedBlue But Wavy
'+~-> Example X = ColorChatBlueBlack("HEY This Is Kewl")
'             SendChat (X)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sup>" & r$ & "</sup>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & "><sub>" & S$ & "</sub>" & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
WavYChaTRedBlue = P$
End Function

Sub Attentioner(thetext)
'+~-> This Will Have Your Text Stand Out
'+~-> Example Call Attention(Text1)
x = FadeByColor3(FADE_BLUE, FADE_GREEN, FADE_BLACK, "¸,.»¬=æ¤º²°A T T E N T I O N°²º¤æ=¬».,¸", True)
P = FadeByColor3(FADE_BLUE, FADE_GREEN, FADE_BLACK, (thetext), False)
Call SendChat(x)
Call SendChat(P)
Call SendChat(x)
End Sub



Function RoomBuster(Room As TextBox, counter As Label)
'+~-> This Is A Room Buster ;-) It Keeps Trying To Get Into The Room Until Your In
'+~-> Call RoomBuster((Text1), (Label1))
d = FindChatRoom4
If d Then KillWin (d)

Do: DoEvents
Call KeyWord("aol://2719:2-2-" + Room + "")
waitforok
counter = counter + 1
If FindChatRoom Then Exit Do
If text2 = 1 Then Exit Do
Loop
End Function
Public Function CountMail3()
'+~-> This Is Another Count Mail Sub
'+~-> Just Put Call CountMail2
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Mailbox% = FindChildByTitle(mdi%, UserSn4 & "'s Online Mailbox")
tabd% = FindChildByClass(Mailbox%, "_AOL_TabControl")
tabp% = FindChildByClass(tabd%, "_AOL_TabPage")
aoltree% = FindChildByClass(tabp%, "_AOL_Tree")

DE = SendMessage(aoltree%, LB_GETCOUNT, 0, 0)
'txtlen% = SendMessageByNum(aoltree%, LB_GETTEXTLEN, ndex, 0&)
'Txt$ = String(txtlen% + 1, 0&)
'X = SendMessageByString(aoltree%, LB_GETTEXT, ndex, Txt$)
MsgBox "You Have" & " " & DE & " Mails"
End Function
Function KillGlyph()
'+~-> Kills the annoying spinning AOL logo in the toobar
'+~-> Call KillGlyph
aol% = FindWindow("AOL Frame25", vbNullString)
AOTooL% = FindChildByClass(aol%, "AOL Toolbar")
AOTool2% = FindChildByClass(AOTooL%, "_AOL_Toolbar")
Glyph% = FindChildByClass(AOTool2%, "_AOL_Glyph")
Call SendMessage(Glyph%, WM_CLOSE, 0, 0)
End Function

Function CoLoRChaTBlueBlack(thetext As String)
'+~-> This Will Make Every Other Letter Blue And Black
'+~-> Example X = ColorChatBlueBlack("HEY This Is Kewl")
'             SendChat (X)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#00F" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#F0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
CoLoRChaTBlueBlack = P$
End Function
Function ColorChatRedGreen(thetext)
'+~-> This Will Make Every Other Letter Red And Green
'+~-> Example X = ColorChatBlueBlack("HEY This Is Kewl")
'             SendChat (X)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#006400" & Chr$(34) & ">" & T$
Next W
ColorChatRedGreen = P$

End Function
Function ColorChatRedBlue(thetext)
'+~-> This Will Make Every Other Letter Red And Blue
'+~-> Example X = ColorChatBlueBlack("HEY This Is Kewl")
'             SendChat (X)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & r$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & U$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" & S$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & ">" & T$
Next W
ColorChatRedBlue = P$

End Function

Function TrimTime()
b$ = Left$(Time$, 5)
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = STR$(HourA)
TrimTime = HourH$ & Right$(b$, 3) & " " & Ap$
End Function
Function TrimTime2()
b$ = Time$
HourH$ = Left$(b$, 2)
HourA = Val(HourH$)
If HourA >= 12 Then Ap$ = "PM"
If HourA = 24 Or HourA < 12 Then Ap$ = "AM"
If HourA > 12 Then
    HourA = HourA - 12
End If
If HourA = 0 Then HourA = 12
HourH$ = STR$(HourA)
TrimTime2 = HourH$ & ":" & Right$(b$, 5) & " " & Ap$
End Function




Function SNfromIM()
'+~-> Gets The SN From An IM
'+~-> Example Text1.text = SNFromIM or List1.AddItem SNFromIM

aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient") '

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMCap$ = GetCaption(IM%)
TheSN$ = Mid(IMCap$, InStr(IMCap$, ":") + 2)
SNfromIM = TheSN$

End Function

Sub playwav(File)
'+~-> This Will Play A Wav File
'+~-> Need A Command Button
'+~-> Call PlayWav("Exit.wav")
SoundName$ = File
   wFlags% = SND_ASYNC Or SND_NODEFAULT
   x% = sndPlaySound(SoundName$, wFlags%)

End Sub

Sub UpdateMenu()
'Microsoft
    frmEditor.mnuFileArray(0).Visible = True            ' Make the initial element visible and display separator bar.
    ArrayNum = ArrayNum + 1                             ' Increment the Index property of the menu control array.
    ' Check to see if Filename is already on the menu list.
    For i = 0 To ArrayNum - 1
        If frmEditor.mnuFileArray(i).Caption = FileName Then
            ArrayNum = ArrayNum - 1
            Exit Sub
        End If
    Next i
    
    ' If filename is not on the menu list, add the menu item.
    Load frmEditor.mnuFileArray(ArrayNum)               ' Create a new menu control.
    frmEditor.mnuFileArray(ArrayNum).Caption = FileName ' Set the caption of the new menu item.
    frmEditor.mnuFileArray(ArrayNum).Visible = True     ' Make the new menu item visible.
End Sub
Sub KillModal()
Modal% = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal%, WM_CLOSE, 0, 0)
End Sub

Sub waitforok()
'+~-> This Will Wait For The OK Message From AOL
Do
DoEvents
okw = FindWindow("#32770", "America Online")
If proG_STAT$ = "OFF" Then
Exit Sub
Exit Do
End If

DoEvents
Loop Until okw <> 0
   
    okb = FindChildByTitle(okw, "OK")
    okd = SendMessageByNum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = SendMessageByNum(okb, WM_LBUTTONUP, 0, 0&)


End Sub

Function WavY(thetext As String)
'+~-> This Will Make Your Text WaVY
'+~-> X = Wavy("Hey Man Dis Shyt Be WavY")
'+~-> SendChat (X)
G$ = thetext
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & S$ & "</sub>" & T$
Next W
WavY = P$

End Function

Sub CenterForm(F As Form)
'+~-> This Will Center The Form
'+~-> Example Call CenterForm Me or Call CenterForm (Form69)
F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
F.Left = Screen.Width / 2 - F.Width / 2
End Sub
Sub Scroll3Line(txt As TextBox)
'+~-> This Will Scroll Your Message 3 Times Per ChatSend
'+~-> Example Call Scroll3Line(Text1)
For fakeoh = 1 To 4
Call SendChat((txt) & (String((116 - Len(txt)), Chr$(4))) & (txt))
Next fakeoh
TimeOut (8)
End Sub
Function CustTalk(Comb1$, Comb2$, Comb3$, Chec1$, Chec2$, Chec3$, Comb4$, txt$, Combt1$, Chec4$, Chec5$, Txt2$, Chec6$)
'+~-> This Is A Comlicated Custom Talker Sub Its Not For YOU
'+~-> Unless Ya Wanna Figure It Out
If Comb2$ = 0 Then Color$ = "#FF0000"
If Comb2$ = 1 Then Color$ = "#00FF00"
If Comb2$ = 2 Then Color$ = "#0000FF"
If Comb2$ = 3 Then Color$ = "#FFFF00"
If Comb2$ = 4 Then Color$ = "#FFFFFF"
If Comb2$ = 5 Then Color$ = "#000000"
If Comb2$ = 6 Then Color$ = "#FF00FF"
If Comb2$ = 7 Then Color$ = "#C0C0C0"
If Comb2$ = 8 Then Color$ = "#FF8080"
If Comb2$ = 9 Then Color$ = "#00C0C0"
If Chec1$ = 1 Then
    Bold$ = "<B>"
Else
    Bold$ = "</B>"
End If
If Chec3$ = 1 Then
    Italic$ = "<I>"
Else
    Italic$ = "</I>"
End If
If Chec4$ = 1 Then
    Underline$ = "<U>"
Else
    Underline$ = "</U>"
End If
If Chec5$ = 1 Then
    Linkk$ = "<A HREF=" & Chr(34) & Txt2$ & Chr(34) & ">"
    Linkkk$ = "</A>"
Else
    Linkk$ = ""
    Linkkk$ = ""
End If
If Comb4$ = 0 Then Talk$ = txt$
If Comb4$ = 1 Then Talk$ = Text_Hacker(txt$)
If Comb4$ = 2 Then Talk$ = Text_Elite(txt$)
If Comb4$ = 3 Then Talk$ = Text_Big(txt$)
If Comb4$ = 4 Then Talk$ = Text_Small(txt$)
If Comb4$ = 5 Then Talk$ = Text_Spaced(txt$)
If Comb4$ = 6 Then Talk$ = Text_backwards(txt$)
If Comb4$ = 7 Then Talk$ = Text_Scrambled(txt$)
Fontt$ = Combt1$
If Chec2$ = 1 Then
    Wavyyy$ = WavY(txt$)
Else
    Wavyyy$ = txt$
End If
If Chec6$ = 1 Then
    If Comb2$ = 0 Then FirstColor$ = FADE_RED
    If Comb2$ = 1 Then FirstColor$ = FADE_GREEN
    If Comb2$ = 2 Then FirstColor$ = FADE_BLUE
    If Comb2$ = 3 Then FirstColor$ = FADE_YELLOW
    If Comb2$ = 4 Then FirstColor$ = FADE_WHITE
    If Comb2$ = 5 Then FirstColor$ = FADE_BLACK
    If Comb2$ = 6 Then FirstColor$ = FADE_PURPLE
    If Comb2$ = 7 Then FirstColor$ = FADE_GREY
    If Comb2$ = 8 Then FirstColor$ = FADE_PINK
    If Comb2$ = 9 Then FirstColor$ = FADE_TURQUOISE
    If Comb3$ = 0 Then SecondColor$ = FADE_RED
    If Comb3$ = 1 Then SecondColor$ = FADE_GREEN
    If Comb3$ = 2 Then SecondColor$ = FADE_BLUE
    If Comb3$ = 3 Then SecondColor$ = FADE_YELLOW
    If Comb3$ = 4 Then SecondColor$ = FADE_WHITE
    If Comb3$ = 5 Then SecondColor$ = FADE_BLACK
    If Comb3$ = 6 Then SecondColor$ = FADE_PURPLE
    If Comb3$ = 7 Then SecondColor$ = FADE_GREY
    If Comb3$ = 8 Then SecondColor$ = FADE_PINK
    If Comb3$ = 9 Then SecondColor$ = FADE_TURQUOISE
    If Chec2$ = True Then
        Fadee$ = FadeByColor2(FirstColor$, SecondColor$, Talk$, True)
    Else
        Fadee$ = FadeByColor2(FirstColor$, SecondColor$, Talk$, False)
    End If
Else
    Fadee$ = Wavyyy$
End If
CustTalkk$ = (Linkk$ & Bold$ & Italic$ & Underline$ & "<FONT FACE=" & Chr(34) & Fontt$ & Chr(34) & " COLOR=" & Chr(34) & Color$ & Chr(34) & ">" & Fadee$ & Linkkk$)
CustTalk = CustTalkk$
End Function


Function Text_Scrambled(thetext)
'+~-> Scrambles TheText
'+~-> X = Text_Scrambled ("Im Scrambled Text, ehhhhh, Phat")
'+~-> Text1.text = X
findlastspace = Mid(thetext, Len(thetext), 1)
If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If
For scrambling = 1 To Len(thetext)
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$
If thechar$ = " " Then
chars$ = Mid(Char$, 1, Len(Char$) - 1)
firstchar$ = Mid(chars$, 1, 1)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
For SpeedBack = Len(midchar$) To 1 Step -1
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe
cityz:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs
sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "
sniffs:
Char$ = ""
backchar$ = ""
End If
Next scrambling
Text_Scrambled = scrambled$
Exit Function
End Function




Sub RespondIM(message)
'+~-> This Will Find An IM Sent To You Sends Your Message And Closes The IM Window
'+~-> Call RespondIM("Go Away Im Busy")
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Sub
Greed:
E = FindChildByClass(IM%, "RICHCNTL")

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)
'e = GetWindow(e, GW_HWNDNEXT)

E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
e2 = GetWindow(E, GW_HWNDNEXT) 'Send Text
E = GetWindow(e2, GW_HWNDNEXT) 'Send Button
Call SendMessageByString(e2, WM_SETTEXT, 0, message)
ClickIcon (E)
Call TimeOut(0.8)
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
E = FindChildByClass(IM%, "RICHCNTL")
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT)
E = GetWindow(E, GW_HWNDNEXT) 'cancel button...
'to close the IM window
ClickIcon (E)
End Sub

Function MessageFromIM()
'+~-> This Will Retrieve The Message From And IM
'+~-> Need A TextBoX, Command Button
'+~-> Text1.text = MessageFromIM
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")

IM% = FindChildByTitle(mdi%, ">Instant Message From:")
If IM% Then GoTo Greed
IM% = FindChildByTitle(mdi%, "  Instant Message From:")
If IM% Then GoTo Greed
Exit Function
Greed:
IMtext% = FindChildByClass(IM%, "RICHCNTL")
IMmessage = GetText(IMtext%)
SN = SNfromIM()
snlen = Len(SNfromIM()) + 3
blah = Mid(IMmessage, InStr(IMmessagge, SN) + snlen)
MessageFromIM = Left(blah, Len(blah) - 1)
End Function

Sub RunMenu(menu1 As Integer, menu2 As Integer)
'+~-> This Will Open An AOL Menu Like File Signoff
'+~-> Example (File, Signoff) Would Be Like Call RunMenu(4,2) I Did This OffLine So Im Not Sure
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
AOLSubMenu% = GetSubMenu(AOLMenus%, menu1)
AOLItemID = GetMenuItemID(AOLSubMenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, AOLItemID, 0&)

End Sub
Sub CloseFile(FileName As String)
'Made By Microsoft
Dim F As Integer
On Error GoTo CloseError    ' If there is an error, display the error message below.
    
    If Dir(FileName) <> "" Then         ' File already exists, so ask if the user wants to overwrite the file.
        response = MsgBox("Overwrite existing file?", vbYesNo + vbQuestion + vbDefaultButton2)
        If response = vbNo Then Exit Sub
    End If
    F = FreeFile
    Open FileName For Output As F       ' Otherwise, open the filename for output.
    Print #F, frmEditor!txtEdit.text    ' Print the current text to the opened file.
    Close F                             ' Close the file.
    FileName = "Untitled"               ' Reset the caption of the main form.
    Exit Sub
CloseError:
    MsgBox "Error occurred while trying to close file, please retry.", 48
    Exit Sub
End Sub
Sub DoUnLoadPreCheck(UnloadMode As Integer)
'Microsoft
    If UnloadMode = 0 Or UnloadMode = 3 Then
            Unload frmAbout
            Unload frmEditor
            End
    End If
End Sub
Sub OpenFile(FileName As String)
'Microsoft
Dim F As Integer
    If "Text Editor: " + FileName = frmEditor.Caption Then  ' Avoid opening a file if it is already loaded.
        Exit Sub
    Else
        On Error GoTo errhandler
            F = FreeFile
            Open FileName For Input As F                    ' Open the file selected in the File Open About dialog box.
            frmEditor!txtEdit.text = Input(LOF(F), F)
            Close F                                         ' Close the file.
            ' frmEditor.mnuFileItem(3).Enabled = True         ' Enable the Close command on the File menu.
            UpdateMenu
            frmEditor.Caption = "Text Editor: " + FileName
            Exit Sub
    End If
errhandler:
        MsgBox "Error encountered while trying to open file, please retry.", 48, "Text Editor"
        Close F
        Exit Sub
End Sub


Sub RunMenuByString(Application, StringSearch)
'+~-> This Will Open The AOL Menu By The Caption
'+~-> Call RunMenuByString("&Sign Off", "&Sign Off")
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



Function Upchat()
'+~-> This Will Enable You To Do Stuff While Uploading
'+~-> Call UpChat
'+~-> The Below Stuff Is To Minimize A Download Also If You Want
'AOL% = FindWindow("AOL Frame25", vbNullString)
'MDI% = FindChildByClass(AOL%, "MDIClient")
'UpWin% = FindChildByTitle(MDI%, "File Transfer")
'Call ShowWindow(UpWin%, SM_MINIMIZE)
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Do
AOModal% = FindChildByClass(mdi%, "_AOL_Modal")
Loop Until AOModal% <> 0
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(aol%, 1)
Call EnableWindow(Upp%, 0)
End Function
Function UnUpchat()
'+~-> This Will Disable Up-Chat
'+~-> call UnUpchat
aol% = FindWindow("AOL Frame25", vbNullString)
mdi% = FindChildByClass(aol%, "MDIClient")
Do
AOModal% = FindChildByClass(mdi%, "_AOL_Modal")
Loop Until AOModal% <> 0
AOGauge% = FindChildByClass(AOModal%, "_AOL_Gauge")
If AOGauge% <> 0 Then Upp% = AOModal%
Call EnableWindow(Upp%, 1)
Call EnableWindow(aol%, 0)
End Function

Function HideAOL()
'+~-> This Will Hide AOL
'+~-> Example Call HideAOL or Better Yet Copy The
'Code Into A Command Button
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 0)
End Function

Function ShowAOL()
'+~-> Hides AOL
'+~-> Call HideAOL
aol% = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(aol%, 5)
End Function

Sub PWSScanner(Dalabel, FilePath$, FileName$, status As Label)
'+~-> Made By GeNOzIdE Modified By izekial83
Dim TheFileLen, NumOne, GenOiZBack, GenOziDe, TheFileInfo$, PWS, PWS2, PWS3, VirusedFile, LengthOfFile, TotalRead, TheTab, TheMSg, TheMsg2, TheMsg3, TheMsg4, TheMsg5, TheDots
StopPWScanner = 0
If FileName$ = "" Then GoTo Errorr
FileName$ = FilePath$ & "\" & FileName$
Dalabel.Caption = FileName$
If Right$(FilePath$, 1) = "\" Then FileName$ = FilePath$ & FileName$
If Not IFileExists(FileName$) Then MsgBox "File Not Found!", 16, "Error": GoTo Errorr
TheFileLen = FileLen(FileName$)
status.Caption = TheFileLen
NumOne = 1
GenOiZBack = 2
GenOziDe = 3
Do While GenOziDe > GenOiZBack
PentiumRest% = DoEvents()
If StopPWScanner = 1 Then GoTo Errorr
Open FileName$ For Binary As #1
If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": GoTo Errorr
TheFileInfo$ = String(32000, 0)
Get #1, NumOne, TheFileInfo$
Close #1
Open FileName$ For Binary As #2
If Err Then MsgBox "An unexpected error occured while opening file!", 16, "Error": GoTo Errorr
PWS = InStr(1, LCase$(TheFileInfo$), "main.idx" + Chr(0), 1)
If PWS Then
Geno:
Mid(TheFileInfo$, PWS) = "GenOziDe  "
PWS2 = Mid(LCase$(TheFileInfo$), PWS + 8 + 1, 8)
PWS2 = Trm(PWS2)
PWS3 = Mid(LCase$(TheFileInfo$), PWS + 8 + 1 + Len(PWS), 1)
If PWS3 <> Chr(0) Then GoTo DeliriuM
If Len(PWS2) < 4 Then GoTo DeliriuM
If Len(PWS2) = "" Then GoTo DeliriuM
DeliriuM:
 PWS = InStr(1, LCase$(TheFileInfo$), "main.idx" + Chr(0), 1)
If PWS <> 0 Then VirusedFile = FileName$: MsgBox VirusedFile & " is a Password Stealer!", 16, "Password Stealer": Close #2: Exit Sub
End If
TotalRead = TotalRead + 32000
status.Caption = Val(TotalRead)
LengthOfFile = LOF(2)
Close #2
If TotalRead > LengthOfFile Then: status.Caption = LengthOfFile: GoTo GOD
DoEvents
Loop
GOD:
TheTab = Chr$(9) & Chr$(9)
TheDots = "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
 '  ---------------------------------------------------------
TheMSg = TheDots & Chr(13) & "File Information:" & Chr(13) & Chr(13)
TheMsg2 = TheMSg & FileName$ & " <- Is Not A Trojan" & Chr(13) & Chr(13)
TheMsg3 = TheMsg2 & FileName$ & " <-  Was Scanned With The PWSD" & Chr(13) & Chr(13)
TheMsg4 = TheMsg3 & "Scanned - 100% of - " & FileName$ & Chr(13) & Chr(13)
TheMsg5 = TheMsg3 & FileName$ & " <- Is Safe For Use!" & Chr(13) & Chr(13) & TheDots
MsgBox TheMsg5, 55, "File Is Clean!"
Errorr:
PentiumRest% = DoEvents()
status.Caption = ""
Close #1
PentiumRest% = DoEvents()
Close #2
PentiumRest% = DoEvents()
Exit Sub
status.Caption = "Ready..."
End Sub
