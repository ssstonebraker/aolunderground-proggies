Attribute VB_Name = "x2k"
',,,,,,,,        ,,,,,,,     ,,,,,,,,,,,,,,,,     ,,,,,,       ,,,,,,,,,
' ´;;;;;;,     ,;;;;;;´    ,;;;;;;;;;;;;;;;;;;,   ;;;;;;     ,;;;;;;;;´             .bas / kast
'   ´;;;;;;, ,;;;;;;´      ;;;;;;´´´´´´´´´;;;; ;;;;;;;   ,;;;;;;;´                32-bit / vb 4,5,6
'     ´;;;;;;;;;;;´        ´´´´´          ;;;;; ;;;;;;,,;;;;;;;´                   for win95x
'      ,;;;;;;;;;,          ,,,,,;;;;;;;;;;;;;´   ;;;;;;;;;;;;;;
'    ,;;;;;;;;;;;;;,        ;;;;;;´´´´´´´´´´   ;;;;;;´´ ;;;;;;;,
'  ,;;;;;;;´ ´;;;;;;;,     ;;;;;,,,,,,,,,,,,,,,  ;;;;;;     ;;;;;;;,
',;;;;;;;´     ´;;;;;;;,   ;;;;;;;;;;;;;;;;;;;;  ;;;;;;     ;;;;;;;,






' declare's
Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SHAddToRecentDocs Lib "Shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function keybd_event Lib "user32" Alias "Keybd_Event" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal Flags As Long, ByVal ExtraInfo As Long)
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function Shell_NotifyIcon Lib "Shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'const's
Public Const LB_GETCOUNT = &H18B
Public Const LB_SETCURSEL = &H186

Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4

Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const SND_LOOP = &H8

Public Const SPI_SCREENSAVERRUNNING = 97

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SW_NORMAL = 1

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

'type-declares

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type
Public Sub clickbutton(Button As Long)
    Call SendMessageLong(Button&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessageLong(Button&, WM_KEYUP, VK_SPACE, 0&)
End Sub



Sub formrestore(Frm As Form)
Frm.WindowState = 0
End Sub

Function showstartbutton()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
    Call ShowWindow(Button&, 5)
End Function
Function showtaskBar()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
    Call ShowWindow(Bar&, 5)
End Function

Function randomnumber(finished)
Randomize
randomnumber = Int((Val(finished) * Rnd) + 1)
End Function

Function percent(Complete As Integer, Total As Variant, TotalOutput As Integer) As Integer
On Error Resume Next
percent = Int(Complete / Total * TotalOutput)
End Function

Sub percentbar(Shape As Control, Done As Integer, Total As Variant)
On Error Resume Next
Shape.AutoRedraw = True
Shape.FillStyle = 0
Shape.DrawStyle = 0
Shape.FontName = "Arial"
Shape.FontSize = 9
Shape.FontBold = True
X = Done / Total * Shape.Width
Shape.Line (0, 0)-(Shape.Width, Shape.Height), RGB(255, 255, 255), BF
Shape.Line (0, 0)-(X - 10, Shape.Height), RGB(0, 0, 255), BF
Shape.CurrentX = (Shape.Width / 2) - 100
Shape.CurrentY = (Shape.Height / 2) - 125
Shape.ForeColor = RGB(255, 0, 0)
Shape.Print percent(Done, Total, 100) & "%"
End Sub
Sub openexe(path As String)
X% = Shell(path, 1): NoFreeze% = DoEvents(): Exit Sub
End Sub
Sub killdupes(lst As ListBox)
For X = 0 To lst.ListCount - 1
current = lst.List(X)
For i = 0 To lst.ListCount - 1
nower = lst.List(i)
If i = X Then GoTo dontkill
If nower = current Then lst.RemoveItem (i)
dontkill:
Next i
Next X
End Sub

Function hidetaskbar()
Dim Bar As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(Bar&, 0)
End Function
Function hidestartbutton()
Dim Bar As Long, Button As Long
Bar& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(Bar&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, 0)
End Function

Sub getwin(kast&)
kast& = GetWindow(kast&, 2)
End Sub
Function gettext(child)
GetTrim = SendMessageByNum(child, 14, 0&, 0&)
TrimSpace$ = Space$(GetTrim)
GetString = SendMessageByString(child, 13, GetTrim + 1, TrimSpace$)
gettext = TrimSpace$
End Function

Public Sub waitforlisttoload(lnglist As Long)
    Dim lngcount As Long
    Do: DoEvents
        Let lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&): Call TimeOut(2&)
        If lngcount& = SendMessage(lnglist&, LB_GETCOUNT, 0&, 0&) Then Exit Do
    Loop
End Sub

Public Sub formcenter(frmform As Form)
    Let frmform.Top = (Screen.Height * 0.85) / 2& - frmform.Height / 2&
    Let frmform.Left = Screen.Width / 2& - frmform.Width / 2&
End Sub
Public Sub windowmaximize(Window As Long)
    Call ShowWindow(Window, SW_MAXIMIZE)
End Sub

Public Sub windowhide(Window As Long)
    Call ShowWindow(Window, SW_HIDE)
End Sub

Public Sub savecombobox(path As String, thecombo As ComboBox)
    Dim Index As Long
    On Error Resume Next
    Open path$ For Output As #1&
    For Index& = 0& To thecombo.ListCount - 1&
        Print #1&, thecombo.List(Index&)
    Next Index&
    Close #1&
End Sub

Public Function replacestring(Replace As String, What As String)
    Dim lngpos As Long
    Do While InStr(1&, Replace$, What$)
        DoEvents
        Let lngpos& = InStr(1&, Replace$, What$)
        Let Replace$ = Left$(Replace$, (lngpos& - 1&)) & Right$(Replace$, Len(Replace$) - (lngpos& + Len(What$) - 1&))
    Loop
    Let replacestring$ = Replace$
End Function
Public Sub printtext(Text As String)
    Dim lngoldcursor As Long
    Let lngoldcursor& = Screen.MousePointer
    Let Screen.MousePointer = 11&
    Printer.Print (Text$)
    Printer.NewPage
    Printer.EndDoc
    Let Screen.MousePointer = lngoldcursor&
End Sub
Public Sub pause(length As Long)
    Dim current As Long
    Let current& = Timer
    Do Until (Timer - current&) >= length&
        DoEvents
    Loop
End Sub

Public Sub loadonstartup()
    Call writeini("windows", "load", App.path & "\" & App.EXEName, "c:\window\win.ini")
End Sub

Public Sub loadlistbox(path As String, thelist As ListBox)
    Dim strlinetext As String
    On Error Resume Next
    Open path$ For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            thelist.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Sub loadcombobox(path As String, thecombo As ComboBox)
    Dim strlinetext As String
    On Error Resume Next
    Open path$ For Input As #1&
        While Not EOF(1&)
            Input #1&, strlinetext$
            DoEvents
            thecombo.AddItem strlinetext$
        Wend
    Close #1&
End Sub
Public Function listfindstring(List As ListBox, FindString As String) As Long
    Dim Index As Long
    If List.ListCount = 0 Then Exit Function
    For Index& = 0 To List.ListCount - 1
        Let List.ListIndex = Index&
        If UCase(List.Text) = UCase(FindString$) Then
            Let listfindstring& = Index&
            Exit Function
            If Err Then Exit Function
        End If
    Next Index&
End Function
Public Function getfontcount() As Long
    Let getfontcount& = Screen.FontCount
End Function

Public Function getcaption(Window As Long) As String
    Dim strBuffer As String, lngtextlen As Long
    Let lngtextlen& = GetWindowTextLength(Window&)
    Let strBuffer$ = String$(lngtextlen&, 0&)
    Call GetWindowText(Window&, strBuffer$, lngtextlen& + 1&)
    Let getcaption$ = strBuffer$
End Function

Public Sub formtotray(frmform As Form)
    Dim systray As NOTIFYICONDATA
    With systray
        Let .cbSize = Len(systray)
        Let .uId = vbNull
        Let .hwnd = frmform.hwnd
        Let .ucallbackMessage = WM_MOUSEMOVE
        Let .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        Let .hIcon = frmform.Icon
        Let .szTip = frmform.Caption
    End With
    Call Shell_NotifyIcon(NIM_ADD, systray)
    frmform.Hide
End Sub

Public Sub formontop(frmform As Form, ontop As Boolean)
    If ontop = True Then Call SetWindowPos(frmform.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
    If ontop = False Then Call SetWindowPos(frmform.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub
Public Sub formfromtray(frmform As Form)
    Dim systray As NOTIFYICONDATA
    With systray
        Let .cbSize = Len(systray)
        Let .hwnd = frmform.hwnd
        Let .uId = vbNull
    End With
    Call Shell_NotifyIcon(NIM_DELETE, systray)
    frmform.Show
End Sub
Public Sub formexitup(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(str(Int(TheForm.Top) - 300))
    Loop Until TheForm.Top < -TheForm.Width
End Sub
Public Sub formexitright(Form As Form)
    Do
        Form.Left = Trim(str(Int(Form.Left) + 300))
        DoEvents
    Loop Until Form.Left > 11000
    If Form.Left > 11000 Then End
End Sub

Public Sub formexitleft(Form As Form)
    Do
        Form.Left = Trim(str(Int(Form.Left) - 300))
        DoEvents
    Loop Until Form.Left < -6300
    If Form.Left < -6300 Then End
End Sub

Public Sub formexitdown(Form As Form)
    Do
        Form.Top = Trim(str(Int(Form.Top) + 300))
        DoEvents
    Loop Until Form.Top > 10000
        If Form.Top > 10000 Then End
End Sub
Sub filedelete(file$)
Dim NoFreeze%
If Not fileexists(file$) Then Exit Sub
Kill file$
NoFreeze% = DoEvents()
End Sub
Public Sub filecopy(file$, Destination$)
    If Not fileexists(file$) Then Exit Sub
    If InStr(file$, ".") = 0& Then Exit Sub
    If InStr(Destination$, "\") = 0& Then Exit Sub
    Call filecopy(file$, Destination$)
End Sub

Public Sub fadepicbox(picbox As Object, color1 As Long, color2 As Long)
    Dim lngcon As Long, longcon As Long, lnghlfwidth As Long, lngcolorval1 As Long
    Dim lngcolorval2 As Long, lngcolorval3 As Long, lngrgb1 As Long, lngrgb2 As Long
    Dim lngrgb3 As Long, lngyval As Long, strcolor1 As String, strcolor2 As String
    Dim strred1 As String, strgreen1 As String, strblue1 As String, strred2 As String
    Dim strgreen2 As String, strblue2 As String, lngred1 As Long, lnggreen1 As Long
    Dim lngblue1 As Long, lngred2 As Long, lnggreen2 As Long, lngblue2 As Long
    Let picbox.AutoRedraw = True
    Let picbox.DrawStyle = 6&
    Let picbox.DrawMode = 13&
    Let picbox.DrawWidth = 2&
    Let lngcon& = 0&
    Let lnghlfwidth& = picbox.Width / 2&
    Let strcolor1$ = gethexfromrgb(color1&)
    Let strcolor2$ = gethexfromrgb(color2&)
    Let strred1$ = "&h" & Right$(strcolor1$, 2&)
    Let strgreen1$ = "&h" & Mid$(strcolor1$, 3&, 2&)
    Let strblue1$ = "&h" & Left$(strcolor1$, 2&)
    Let strred2$ = "&h" & Right$(strcolor2$, 2&)
    Let strgreen2$ = "&h" & Mid$(strcolor2$, 3&, 2&)
    Let strblue2$ = "&h" & Left$(strcolor2$, 2&)
    Let lngred1& = Val(strred1$)
    Let lnggreen1& = Val(strgreen1$)
    Let lngblue1& = Val(strblue1$)
    Let lngred2& = Val(strred2$)
    Let lnggreen2& = Val(strgreen2$)
    Let lngblue2& = Val(strblue2$)
    Do: DoEvents
        On Error Resume Next
        Let lngcolorval1& = lngred2& - lngred1&
        Let lngcolorval2& = lnggreen2& - lnggreen1&
        Let lngcolorval3& = lngblue2& - lngblue1&
        Let lngrgb1& = (lngcolorval1& / lnghlfwidth& * lngcon&) + lngred1&
        Let lngrgb2& = (lngcolorval2& / lnghlfwidth& * lngcon&) + lnggreen1&
        Let lngrgb3& = (lngcolorval3& / lnghlfwidth& * lngcon&) + lngblue1&
        picbox.Line (lngyval&, 0&)-(lngyval& + 2&, picbox.Height), RGB(lngrgb1&, lngrgb2&, lngrgb3&), BF
        Let lngyval& = lngyval& + 10&
        Let lngcon& = lngcon& + 5&
    Loop Until lngcon& > lnghlfwidth&
End Sub
Public Function gethexfromrgb(rgbvalue As Long) As String
    Dim hexstate As String, hexlen As Long
    Let hexstate$ = Hex(rgbvalue&)
    Let hexlen& = Len(hexstate$)
    Select Case hexlen&
        Case 1&
            Let gethexfromrgb$ = "00000" & hexstate$
            Exit Function
        Case 2&
            Let gethexfromrgb$ = "0000" & hexstate$
            Exit Function
        Case 3&
            Let gethexfromrgb$ = "000" & hexstate$
            Exit Function
        Case 4&
            Let gethexfromrgb$ = "00" & hexstate$
            Exit Function
        Case 5&
            Let gethexfromrgb$ = "0" & hexstate$
            Exit Function
        Case 6&
            Let gethexfromrgb$ = "" & hexstate$
            Exit Function
        Case Else
            Exit Function
    End Select
End Function

Public Sub ctrlaltdel(enabled As Boolean)
    Dim lnggogo As Long, pOld As Boolean
    Let lnggogo& = SystemParametersInfo(SPI_SCREENSAVERRUNNING, enabled, pOld, 0&)
End Sub

Public Sub removelistitem(thelist As ListBox, entry As String)
    Dim lngindex As Long
    If thelist.ListCount = 0& Then Exit Sub
    For lngindex& = 0& To thelist.ListCount - 1&
        Let thelist.ListIndex = lngindex&
        If LCase$(thelist.List(lngindex&)) = LCase$(entry$) Then
            Call thelist.RemoveItem(lngindex&)
            Exit Sub
            If Err Then Exit Sub
        End If
    Next lngindex&
End Sub
Public Sub clickicondouble(lnghwnd As Long)
    Call SendMessage(lnghwnd&, WM_LBUTTONDBLCLK, 0&, 0&)
End Sub

Function iswindows95() As Boolean
Const dwMask95 = &H2&
If iswindows95 = True Then
iswindows95 = 1
Else
iswindows95 = False
iswindows95 = 1
End If
End Function
Function iswindows98() As Boolean
Const dwMask98 = &H2&
If iswindows98 = True Then
iswindows98 = 1
Else
iswindows98 = False
iswindows98 = 1
End If
End Function

Function ucase16(ByVal str As String)
#If Win16 Then
    ucase16 = UCase$(str)
#Else
    ucase16 = str
#End If
End Function
Sub msgbox_error(Text As String, program As String)
MsgBox Text, 16, program
End Sub

Function clipboardgettext()
Txt$ = Clipboard.gettext
clipboardgettext = Txt$
End Function
Public Sub clicklistindex(lnglist As Long, lngindex As Long)
    Call SendMessage(lnglist&, LB_SETCURSEL, CLng(lngindex&), 0&)
End Sub
Public Function fileexists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        fileexists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        fileexists = True
    Else
        fileexists = False
    End If
End Function
Public Function linecount(MyString As String) As Long
    Dim Spot As Long, Count As Long
    If Len(MyString$) < 1 Then
        linecount& = 0&
        Exit Function
    End If
    Spot& = InStr(MyString$, Chr(13))
    If Spot& <> 0& Then
        linecount& = 1
        Do
            Spot& = InStr(Spot + 1, MyString$, Chr(13))
            If Spot& <> 0& Then
                linecount& = linecount& + 1
            End If
        Loop Until Spot& = 0&
    End If
    linecount& = linecount& + 1
End Function

Public Sub playmidi(midifile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(midifile$)
    If SafeFile$ <> "" Then
        Call mciSendString("play " & midifile$, 0&, 0, 0)
    End If
End Sub
Public Sub runwebpage(YourForm As Form, URL As String)
    Call ShellExecute(YourForm.hwnd, "Open", URL, "", "", SW_NORMAL)
End Sub
Public Sub clickicon(lngico As Long)
    Call SendMessageLong(lngico&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(lngico&, WM_LBUTTONUP, 0&, 0&)
End Sub



Public Sub runmenu(hwnd As Long, FirstMenu As Long, SecondMenu As Long)
    Dim a As Long, B As Long, c As Long
    a = GetMenu(hwnd)
    B = GetSubMenu(a, FirstMenu)
    c = GetMenuItemID(B, SecondMenu)
    Call SendMessageByNum(hwnd, WM_COMMAND, c, 0&)
End Sub
Public Function countchr(ToCountIn As String, ToCount As String, CaseSensitive As Boolean) As Long
    Dim a As Long, B As String, c As Long
    c = 0
    For a = 1 To Len(ToCountIn)
        B = Mid$(ToCountIn, a, 1)
        If CaseSensitive = True Then
            If B = ToCount Then
                c = c + 1
            End If
        ElseIf CaseSensitive = False Then
            If LCase(B) = LCase(ToCount) Then
                c = c + 1
            End If
        End If
    Next a
    countchr = c
End Function
Function trayhide()
Dim ShelltrayWnd As Long, TraynotifyWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(TraynotifyWnd&, SW_HIDE)
End Function

Function trayshow()
Dim ShelltrayWnd As Long, TraynotifyWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(TraynotifyWnd&, SW_SHOW)
End Function

Public Sub beep()
'beeps a sound on your speaker
MessageBeep -1&
End Sub
Function desktophide()
Dim gone As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long
gone& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(gone&, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer& = FindWindowEx(SHELLDLLDefView&, 0&, "Internet Explorer_Server", vbNullString)
Call ShowWindow(InternetExplorerServer&, SW_HIDE)

End Function
Function printscreen()
Const VK_SNAPSHOT = &H2C
    Call keybd_event(VK_SNAPSHOT, 1, 0&, 0&)
End Function
Function desktopshow()
Dim gone As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long
gone& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(gone&, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer& = FindWindowEx(SHELLDLLDefView&, 0&, "Internet Explorer_Server", vbNullString)
Call ShowWindow(InternetExplorerServer&, SW_SHOW)
End Function
Function fileopen(ThePath As String) As String
    Dim a As Long, B As String
    a = FreeFile
    Open ThePath For Binary As #a
    B = String$(LOF(a), " ")
    Get #a, , B
    fileopen = B
    Close #a
End Function
Public Function getline(ToGetFrom As String, LineNumber As Long) As String
    Dim a As Long, B As Long, c As String
    On Error Resume Next
    c = ToGetFrom
    If LineNumber = 1 Then
        B = InStr(c, Chr(13))
        c = Left(c, B - 1)
        getline = c
        Exit Function
    End If
    For a = 1 To LineNumber - 1
        B = InStr(c, Chr(13))
        c = Mid$(c, B + 2)
    Next a
    B = InStr(c, Chr(13))
    c = Left(c, B - 1)
    getline = c
End Function
Public Function getwinparent(hwnd As Long) As Long
    Dim a As Long
    a = GetParent(hwnd)
    getwinparent = a
End Function
Public Sub formcoolexit(Form As Form)
    Dim sStart As Integer, GoNow As Long
    GoNow& = Form.Height / 2
    For sStart% = 1 To GoNow&
    DoEvents
        Form.Height = Form.Height - 10
        Form.Top = (Screen.Height - Form.Height) \ 2
        If Form.Height <= 11 Then GoTo Finish
    Next sStart%
Finish:
        Form.Height = 30
        GoNow& = Form.Width / 2
    For sStart% = 1 To GoNow&
    DoEvents
        Form.Width = Form.Width - 10
        Form.Left = (Screen.Width - Form.Width) \ 2
        If Form.Width <= 11 Then Exit Sub
    Next sStart%
    Form.end
End Sub

Function freeprocess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
End Function
Sub formmaximize(Frm As Form)
Frm.WindowState = 2
End Sub
Sub formminimize(Frm As Form)
Frm.WindowState = 1
End Sub

Public Sub settext(Window As Long, Text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub
Public Sub windowshow(Window As Long)
    Call ShowWindow(Window&, SW_SHOW)
End Sub
Public Sub formdrag(frmform As Form)
    Call ReleaseCapture
    Call SendMessage(frmform.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub showctl()
'show's your open program in ctl + alt + del
Dim X As Long
Dim k As Long
X = GetCurrentProcessId()
k = RegisterServiceProcess(X, RSP_SIMPLE_SERVICE)
End Sub


Public Sub hidectl()
'hide's from ctl alt del
Dim X As Long
Dim k As Long
X = GetCurrentProcessId()
k = RegisterServiceProcess(X, RSP_UNREGISTER_SERVICE)
End Sub



Public Sub windowclose(Window As Long)
    Call SendMessageLong(Window, WM_CLOSE, 0&, 0&)
End Sub

Public Sub wavloop(file As String)
    Dim lngflags As Long
        Let lngflags& = SND_ASYNC Or SND_LOOP
        Call sndPlaySound(file$, lngflags&)
End Sub

Public Sub wavloopstop()
    Dim lngflags As Long
    Let lngflags& = SND_ASYNC Or SND_LOOP
    Call sndPlaySound("", lngflags&)
End Sub

Public Sub wavstop()
    Call sndPlaySound("", SND_FLAG)
End Sub

Public Function trimspaces(Text As String) As String
    Let trimspaces$ = replacestring(Text$, " ", "")
End Function
Public Sub savelistbox(path As String, thelist As ListBox)
    Dim Index As Long
    On Error Resume Next
    Open path$ For Output As #1&
    For Index& = 0& To thelist.ListCount - 1&
        Print #1&, thelist.List(Index&)
    Next Index&
    Close #1&
End Sub

Public Sub writetoini(Section As String, Key As String, KeyValue As String, fullpath As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, fullpath$)
End Sub
Public Function getfromini(Section As String, Key As String, fullpath As String) As String
   Dim buffer As String
   Let buffer$ = String$(750, Chr$(0&))
   Let getfromini$ = Left$(buffer$, GetPrivateProfileString(Section$, ByVal LCase$(Key$), "", buffer, Len(buffer), fullpath$))
End Function
Public Sub trans_form(Frm As Form)
Const RGN_DIFF = 4
Const RGN_OR = 2
Dim outer_rgn As Long
Dim inner_rgn As Long
Dim wID As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
Dim ctl_left As Single
Dim ctl_top As Single
Dim ctl_right As Single
Dim ctl_bottom As Single
Dim control_rgn As Long
Dim combined_rgn As Long
Dim ctl As Control

    If Frm.WindowState = vbMinimized Then Exit Sub
    wID = ScaleX(Width, vbTwips, vbPixels)
    hgt = ScaleY(Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wID, hgt)

    border_width = (wID - ScaleWidth) / 2
    title_height = hgt - border_width - ScaleHeight
    inner_rgn = CreateRectRgn( _
        border_width, _
        title_height, _
        wID - border_width, _
        hgt - border_width)
    combined_rgn = CreateRectRgn(0, 0, 0, 0)
    CombineRgn combined_rgn, outer_rgn, _
        inner_rgn, RGN_DIFF
    For Each ctl In Controls
        If ctl.Container Is Frm Then
            ctl_left = ScaleX(ctl.Left, Frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = ScaleX(ctl.Top, Frm.ScaleMode, vbPixels) _
                + title_height
            ctl_right = ScaleX(ctl.Width, Frm.ScaleMode, vbPixels) _
                + ctl_left
            ctl_bottom = ScaleX(ctl.Height, Frm.ScaleMode, vbPixels) _
                + ctl_top
            control_rgn = CreateRectRgn( _
                ctl_left, ctl_top, _
                ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, _
                control_rgn, RGN_OR
        End If
    Next ctl
    SetWindowRgn hwnd, combined_rgn, True
End Sub

Private Sub untrans_form(Frm As Form)
Const RGN_DIFF = 4
Const RGN_OR = 2
Dim outer_rgn As Long
Dim inner_rgn As Long
Dim wID As Single
Dim hgt As Single
Dim border_width As Single
Dim title_height As Single
Dim ctl_left As Single
Dim ctl_top As Single
Dim ctl_right As Single
Dim ctl_bottom As Single
Dim control_rgn As Long
Dim combined_rgn As Long
Dim ctl As Control

    If Frm.WindowState = vbMinimized Then
    Exit Sub
    End If
    wID = ScaleX(Width, vbTwips, vbPixels)
    hgt = ScaleY(Height, vbTwips, vbPixels)
    outer_rgn = CreateRectRgn(0, 0, wID, hgt)

    border_width = (wID - ScaleWidth) / 2
    title_height = hgt - border_width - ScaleHeight
    inner_rgn = CreateRectRgn( _
        border_width, _
        title_height, _
        wID - border_width, _
        hgt - border_width)

    

    ' Create the control regions.
    For Each ctl In Controls
        If ctl.Container Is Frm Then
            ctl_left = ScaleX(ctl.Left, Frm.ScaleMode, vbPixels) _
                + border_width
            ctl_top = ScaleX(ctl.Top, Frm.ScaleMode, vbPixels) _
                + title_height
            ctl_right = ScaleX(ctl.Width, Frm.ScaleMode, vbPixels) _
                + ctl_left
            ctl_bottom = ScaleX(ctl.Height, Frm.ScaleMode, vbPixels) _
                + ctl_top
            control_rgn = CreateRectRgn( _
                ctl_left, ctl_top, _
                ctl_right, ctl_bottom)
            CombineRgn combined_rgn, combined_rgn, _
                control_rgn, RGN_OR
        End If
    Next ctl
    SetWindowRgn hwnd, combined_rgn, True
End Sub

