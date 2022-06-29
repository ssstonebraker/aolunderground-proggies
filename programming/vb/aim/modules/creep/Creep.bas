Attribute VB_Name = "Creep"
'-[Creep.bas for AIM 3.0/3.5/Maybe 4.0]-
'-[Made By: Creep
'-[Screen Name: iicreepLL@aol.com
'-[Credits: Monkefade.bas for some subs
'-[Web Page: http://fly.to/phrozen







Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Declare Function mciSendString Lib "MMSystem" Alias "mcisendstring" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal wReturnLength As Integer, ByVal hCallback As Integer) As Long

Const SND_SYNC = &H0

    Public Const EM_GETLINECOUNT = &HBA
    Public Const EM_GETLINE = &HC4
    Public Const SND_ASYNC = &H1
    Public Const SND_NODEFAULT = &H2
    Public Const SND_MEMORY = &H4
    Public Const SND_LOOP = &H8
    Public Const SND_NOSTOP = &H10
    Public Const WM_CLOSE = &H10
    Public Const WM_SETTEXT = &HC
    Public Const WM_LBUTTONUP = &H202
    Public Const WM_LBUTTONDOWN = &H201
    Public Const SW_HIDE = 0
    Public Const SW_SHOW = 5
    Public Const SW_RESTORE = 9
    Public Const SW_MAXIMIZE = 3
    Public Const SW_MINIMIZE = 6
    Public Const GW_HWNDNEXT = 2
    Public Const HWND_NOTOPMOST = -2
    Public Const HWND_TOPMOST = -1
    Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1



Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETtext = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185


Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT




Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10

Public Const VK_UP = &H26

Public Const WM_CHAR = &H102

Public Const WM_COMMAND = &H111
Public Const WM_gettext = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100

Public Const WM_LBUTTONDBLCLK = &H203


Public Const WM_MOVE = &HF012

Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const ENTER_KEY = 13
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        X As Long
        Y As Long
End Type
' The Fades are all borrowed from MonkEFade3.bas so give him
'Proper credit.

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

''''
Public Sub SetText(Window As Long, text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub
Public Function GetText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_gettext, TextLength& + 1, buffer$)
    GetText$ = buffer$
End Function
Function GetCaption(hwnd)
hwndLength% = GetWindowTextLength(hwnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hwnd, hwndTitle$, (hwndLength% + 1))
GetCaption = hwndTitle$
End Function
Function AIM_Attention(message As String)
AIM_SendChat ("_.·´¯`·..·•A•T•T•E•N•T•I•O•N•·..·´¯`·._")
AIM_SendChat message
AIM_SendChat ("¯`·._.····•A•T•T•E•N•T•I•O•N•····._.·`¯")
End Function
Public Sub AIM_Button(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub

Function AIM_Click_SendIm()
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ourparent& = FindWindowEx(AIMwindow&, 0, "_Oscar_TabGroup", vbNullString)
ourhandle& = FindWindowEx(ourparent&, 0, "_Oscar_IconBtn", vbNullString)
Call ICON(ourhandle)
End Function
Function AIM_Click_Invite()
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Tabgroup& = FindWindowEx(AIMwindow&, 0, "_Oscar_TabGroup", vbNullString)
IconGroup& = FindWindowEx(Tabgroup&, 0, "_Oscar_IconBtn", vbNullString)
InviteIcon& = FindWindowEx(Tabgroup&, IconGroup&, "_Oscar_IconBtn", vbNullString)
Call ICON(InviteIcon)
End Function
Function AIM_Click_Voice()
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Tabgroup& = FindWindowEx(AIMwindow&, 0, "_Oscar_TabGroup", vbNullString)
Iconthing1& = FindWindowEx(Tabgroup&, 0, "_Oscar_IconBtn", vbNullString)
Iconthing2& = FindWindowEx(Tabgroup&, Iconthing1&, "_Oscar_IconBtn", vbNullString)
VoiceIcon& = FindWindowEx(Tabgroup&, Iconthing2&, "_Oscar_IconBtn", vbNullString)
Call ICON(VoiceIcon)
End Function
Function AIM_Exit()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "&Close")
End Function
Function AIM_SwitchSN()
Call RunMenuByString(FindWindow("_Oscar_BuddyListWin", vbNullString), "S&witch Screen Name...")
End Function
Function Form_Ontop(frm As Form)
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Function
Function Form_NotOntop(frm As Form)
SetWinOnTop = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Function
Function AIM_Get_ChatName()
Dim chat As Long
On Error Resume Next
chat& = FindWindow("AIM_ChatWnd", vbNullString)
RoomName$ = GetCaption(chat&)
Room$ = Mid(RoomName$, InStr(RoomName$, ":") + 2)
AIM_Get_ChatName = Room$
End Function
Function AIM_Get_IMSender()
Do
DoEvents
IMwindow& = FindWindow("AIM_IMessage", vbNullString)
Loop Until IMwindow <> 0

TotalCaption$ = GetCaption(IMwindow)
SNfromcaption$ = InStr(TotalCaption$, "-")
SNfromcaption$ = SNfromcaption$ - 1
SNfromcaption$ = Left(TotalCaption$, SNfromcaption$)
AIM_Get_IMSender = SNfromcaption$
End Function
Function AIM_Get_User()
Dim aimcaption As String
Dim AIMwindow As Long
AIMwindow = FindWindow("_Oscar_BuddyListWin", vbNullString)
aimcaption = GetCaption(AIMwindow)
pos$ = InStr(aimcaption, "'")
SN$ = Left(aimcaption, pos - 1)
AIM_Get_User = SN
End Function
Public Sub AIM_icon(aIcon As Long)
    Call SendMessage(aIcon&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(aIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Function AIM_IsOnline()
Dim aim As Long
aim& = FindWindow("_Oscar_BuddyListWin", vbNullString)
IsAIMOnline = aim&
End Function
Function AIM_KillAd()
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
AIMthing1& = FindWindowEx(AIMwindow&, 0, "WndAte32Class", vbNullString)
AIMad& = FindWindowEx(AIMthing1&, 0, "Ate32Class", vbNullString)
ShowWindow AIMad, SW_HIDE

End Function
Function AIM_MassIM(TheList As ListBox, message As String)
If TheListList.ListCount = 0 Then
Do: DoEvents: Loop
End If
TheListList.Enabled = False
i = TheListList.ListCount - 1
TheListList.ListIndex = 0
For X = 0 To i
TheListList.ListIndex = X
Call AIM_SendIM(TheListList.text, message)
TimeOut (0.8)
Next X
TheListList.Enabled = True
End Function
Function AIM_RespondToIM(message As String)

Do
DoEvents
IMwindow& = FindWindow("AIM_IMessage", vbNullString)
Loop Until IMwindow <> 0

wndateclass& = FindWindowEx(IMwindow&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
AIMmessagebox& = FindWindowEx(IMwindow&, wndateclass&, "wndate32class", vbNullString)
SetText AIMmessagebox, message

SendIcon& = FindWindowEx(IMwindow&, 0, "_Oscar_IconBtn", vbNullString)
AIM_icon SendIcon
End Function
Function AIM_ShowAd()
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
AIMthing1& = FindWindowEx(AIMwindow&, 0, "WndAte32Class", vbNullString)
AIMad& = FindWindowEx(AIMthing1&, 0, "Ate32Class", vbNullString)
ShowWindow AIMad, SW_SHOW
End Function
Function AIM_GotoUrl(URL As String)
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
TheUrl& = FindWindowEx(AIMwindow&, 0, "Edit", vbNullString)
SetText TheUrl, URL
DoEvents
Goicon& = FindWindowEx(AIMwindow&, 0, "_Oscar_IconBtn", vbNullString)
AIM_icon Goicon
End Function
Function AIM_Hide()
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ShowWindow AIMwindow, SW_HIDE
End Function
Function AIM_Show()
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ShowWindow AIMwindow, SW_SHOW
End Function
Function AIM_SendChat(what As String)

chatroom& = FindWindow("AIM_ChatWnd", vbNullString)
chatparent& = FindWindowEx(chatroom&, 0, "WndAte32Class", vbNullString)
ourhandle& = FindWindowEx(chatparent&, 0, "Ate32Class", vbNullString)
ChatBox& = FindWindowEx(chatroom&, chatparent&, "WndAte32Class", vbNullString)
SetText ChatBox, what

DoEvents

Hand1& = FindWindowEx(chatroom&, 0, "_Oscar_IconBtn", vbNullString)
Hand2& = FindWindowEx(chatroom&, Hand1&, "_Oscar_IconBtn", vbNullString)
Hand3& = FindWindowEx(chatroom&, Hand2&, "_Oscar_IconBtn", vbNullString)
Hand4& = FindWindowEx(chatroom&, Hand3&, "_Oscar_IconBtn", vbNullString)
SendButton& = FindWindowEx(chatroom&, Hand4&, "_Oscar_IconBtn", vbNullString)
AIM_icon SendButton
End Function
Function AIM_ClearChat()
chatroom& = FindWindow("AIM_ChatWnd", vbNullString)
chatparent& = FindWindowEx(chatroom&, 0, "WndAte32Class", vbNullString)
ourhandle& = FindWindowEx(chatparent&, 0, "Ate32Class", vbNullString)
SetText ourhandle, ""
End Function
Function AIM_SendIM(SN As String, message As String)

AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
ourparent& = FindWindowEx(AIMwindow&, 0, "_Oscar_TabGroup", vbNullString)
ourhandle& = FindWindowEx(ourparent&, 0, "_Oscar_IconBtn", vbNullString)
Call AIM_icon(ourhandle)

Do
DoEvents
IMwindow& = FindWindow("AIM_IMessage", vbNullString)
Loop Until IMwindow <> 0


IMPlace& = FindWindowEx(IMwindow&, 0, "_Oscar_PersistantCombo", vbNullString)
SNbox& = FindWindowEx(IMPlace&, 0, "Edit", vbNullString)
Call SetText(SNbox, SN)

wndateclass& = FindWindowEx(IMwindow&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
MessageBox& = FindWindowEx(IMwindow&, wndateclass&, "wndate32class", vbNullString)
Call SetText(MessageBox, message)

DoEvents

SendIcon& = FindWindowEx(IMwindow&, 0, "_Oscar_IconBtn", vbNullString)
Call AIM_icon(SendIcon)
End Function

Function AIM_SendInvite(ScreenNames As String, message As String, RoomName As String)
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Tabgroup& = FindWindowEx(AIMwindow&, 0, "_Oscar_TabGroup", vbNullString)
IconGroup& = FindWindowEx(Tabgroup&, 0, "_Oscar_IconBtn", vbNullString)
InviteIcon& = FindWindowEx(Tabgroup&, IconGroup&, "_Oscar_IconBtn", vbNullString)
Call AIM_icon(InviteIcon)

Do
DoEvents
InviteWindow& = FindWindow("AIM_ChatInviteSendWnd", vbNullString)
Loop Until InviteWindow <> 0

InviteSNs& = FindWindowEx(InviteWindow&, 0, "Edit", vbNullString)
Call SetText(InviteSNs, ScreenNames)

DoEvents

InvitePart& = FindWindowEx(InviteWindow&, 0, "Edit", vbNullString)
InviteMessage& = FindWindowEx(InviteWindow&, InvitePart&, "Edit", vbNullString)
Call SetText(InviteMessage, message)

InviteBox1& = FindWindowEx(InviteWindow&, 0, "Edit", vbNullString)
InviteBox2& = FindWindowEx(InviteWindow&, InviteBox1&, "Edit", vbNullString)
InviteRoomName& = FindWindowEx(InviteWindow&, InviteBox2&, "Edit", vbNullString)
Call SetText(InviteRoomName, RoomName)

DoEvents

IconPart1& = FindWindowEx(InviteWindow&, 0, "_Oscar_IconBtn", vbNullString)
IconPart2& = FindWindowEx(InviteWindow&, IconPart1&, "_Oscar_IconBtn", vbNullString)
SendIcon& = FindWindowEx(InviteWindow&, IconPart2&, "_Oscar_IconBtn", vbNullString)
Call AIM_icon(SendIcon)
End Function
Function AIM_SendLink_Chat(Link As String, message As String)
AIM_SendChat "<a href=""" + Link + """>" + message + ""
End Function
Function AIM_SendLink_IM(who As String, Link As String, message As String)
AIM_SendIM who, "<a href=""" + Link + """>" + message + ""
End Function


Sub FadeForm(FormX As Form, Colr1, Colr2)
'by monk-e-god (modified from a sub by MaRZ)
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
'by aDRaMoLEk
FadedText$ = Replacer(FadedText$, Chr(13), "+chr13+")
OSM = PicB.ScaleMode
PicB.ScaleMode = 3
TextOffX = 0: TextOffY = 0
StartX = 2: StartY = 0
PicB.Font = "Arial": PicB.FontSize = 10
PicB.FontBold = False: PicB.FontItalic = False: PicB.FontUnderline = False: PicB.FontStrikethru = False
PicB.AutoRedraw = True: PicB.ForeColor = 0&: PicB.Cls
For X = 1 To Len(FadedText$)
  C$ = Mid$(FadedText$, X, 1)
  If C$ = "<" Then
    tagstart = X + 1
    tagend = InStr(X + 1, FadedText$, ">") - 1
    t$ = LCase$(Mid$(FadedText$, tagstart, (tagend - tagstart) + 1))
    X = tagend + 1
    Select Case t$
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
        If Left$(t$, 10) = "font color" Then 'change font color
          ColorStart = InStr(t$, "#")
          ColorString$ = Mid$(t$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          PicB.ForeColor = RGB(RV, GV, BV)
        End If
        If Left$(t$, 9) = "font face" Then 'added by monk-e-god
            fontstart% = InStr(t$, Chr(34))
            dafont$ = Right(t$, Len(t$) - fontstart%)
            PicB.Font = dafont$
        End If
    End Select
  Else  'normal text
    If C$ = "+" And Mid(FadedText$, X, 7) = "+chr13+" Then ' added by monk-e-god
        StartY = StartY + 16
        TextOffX = 0
        X = X + 6
    Else
        PicB.CurrentY = StartY + TextOffY
        PicB.CurrentX = StartX + TextOffX
        PicB.Print C$
        TextOffX = TextOffX + PicB.TextWidth(C$)
    End If
  End If
Next X
PicB.ScaleMode = OSM
End Sub

Function GetRGB(ByVal CVal As Long) As COLORRGB
  GetRGB.Blue = Int(CVal / 65536)
  GetRGB.Green = Int((CVal - (65536 * GetRGB.Blue)) / 256)
  GetRGB.Red = CVal - (65536 * GetRGB.Blue + 256 * GetRGB.Green)
End Function
Sub FadePreview2(RichTB As Control, ByVal FadedText As String)
'Modified by monk-e-god for use in a RichTextBox

'NOTE: RichTB must be a RichTextBox.
'NOTE: You cannot preview wavy fades with this sub.
Dim StartPlace%
StartPlace% = 0
RichTB.SelStart = StartPlace%
RichTB.SelBold = False: RichTB.SelItalic = False: RichTB.SelUnderline = False: RichTB.SelStrikeThru = False
RichTB.SelColor = 0&: RichTB.text = ""
For X = 1 To Len(FadedText$)
  C$ = Mid$(FadedText$, X, 1)
  RichTB.SelStart = StartPlace%
  RichTB.SelLength = 1
  If C$ = "<" Then
    tagstart = X + 1
    tagend = InStr(X + 1, FadedText$, ">") - 1
    t$ = LCase$(Mid$(FadedText$, tagstart, (tagend - tagstart) + 1))
    X = tagend + 1
    RichTB.SelStart = StartPlace%
    RichTB.SelLength = 1
    Select Case t$
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
        If Left$(t$, 10) = "font color" Then 'change font color
          ColorStart = InStr(t$, "#")
          ColorString$ = Mid$(t$, ColorStart + 1, 6)
          RedString$ = Left$(ColorString$, 2)
          GreenString$ = Mid$(ColorString$, 3, 2)
          BlueString$ = Right$(ColorString$, 2)
          RV = Hex2Dec!(RedString$)
          GV = Hex2Dec!(GreenString$)
          BV = Hex2Dec!(BlueString$)
          RichTB.SelStart = StartPlace%
          RichTB.SelColor = RGB(RV, GV, BV)
        End If
        If Left$(t$, 9) = "font face" Then
            fontstart% = InStr(t$, Chr(34))
            dafont$ = Right(t$, Len(t$) - fontstart%)
            RichTB.SelStart = StartPlace%
            RichTB.SelFontName = dafont$
        End If
    End Select
  Else  'normal text
    RichTB.SelText = RichTB.SelText + C$
    StartPlace% = StartPlace% + 1
    RichTB.SelStart = StartPlace%
  End If
Next X
End Sub

Function Hex2Dec!(ByVal strHex$)
'by aDRaMoLEk
  If Len(strHex$) > 8 Then strHex$ = Right$(strHex$, 8)
  Hex2Dec = 0
  For X = Len(strHex$) To 1 Step -1
    CurCharVal = GETVAL(Mid$(UCase$(strHex$), X, 1))
    Hex2Dec = Hex2Dec + CurCharVal * 16 ^ (Len(strHex$) - X)
  Next X
End Function

Function GETVAL%(ByVal strLetter$)
'by aDRaMoLEk
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

Function FadeCLRBars(RedBar As Control, GreenBar As Control, BlueBar As Control)
'This gets a color from 3 scroll bars
FadeCLRBars = RGB(RedBar.Value, GreenBar.Value, BlueBar.Value)

'Put this in the scroll event of the
'3 scroll bars RedScroll1, GreenScroll1,
'& BlueScroll1.  It changes the backcolor
'of ColorLbl when you scroll the bars
'ColorLbl.BackColor = CLRBars(RedScroll1, GreenScroll1, BlueScroll1)

End Function

Function FadeByColor10(Colr1, Colr2, Colr3, Colr4, Colr5, Colr6, Colr7, Colr8, Colr9, Colr10, TheText$, Wavy As Boolean)

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


FadeByColor10 = FadeTenColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, rednum6%, greennum6%, bluenum6%, rednum7%, greennum7%, bluenum7%, rednum8%, greennum8%, bluenum8%, rednum9%, greennum9%, bluenum9%, rednum10%, greennum10%, bluenum10%, TheText, Wavy)

End Function

Function FadeByColor2(Colr1, Colr2, TheText$, Wavy As Boolean)

dacolor1$ = RGBtoHEX(Colr1)
dacolor2$ = RGBtoHEX(Colr2)

rednum1% = Val("&H" + Right(dacolor1$, 2))
greennum1% = Val("&H" + Mid(dacolor1$, 3, 2))
bluenum1% = Val("&H" + Left(dacolor1$, 2))
rednum2% = Val("&H" + Right(dacolor2$, 2))
greennum2% = Val("&H" + Mid(dacolor2$, 3, 2))
bluenum2% = Val("&H" + Left(dacolor2$, 2))

FadeByColor2 = FadeTwoColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, TheText, Wavy)

End Function
Function FadeByColor3(Colr1, Colr2, Colr3, TheText$, Wavy As Boolean)

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

FadeByColor3 = FadeThreeColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, TheText, Wavy)

End Function
Function FadeByColor4(Colr1, Colr2, Colr3, Colr4, TheText$, Wavy As Boolean)

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

FadeByColor4 = FadeFourColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, TheText, Wavy)

End Function

Function FadeByColor5(Colr1, Colr2, Colr3, Colr4, Colr5, TheText$, Wavy As Boolean)

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

FadeByColor5 = FadeFiveColor(rednum1%, greennum1%, bluenum1%, rednum2%, greennum2%, bluenum2%, rednum3%, greennum3%, bluenum3%, rednum4%, greennum4%, bluenum4%, rednum5%, greennum5%, bluenum5%, TheText, Wavy)

End Function

Function FadeFiveColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    
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
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Right(TheText, frthlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
Function FadeTenColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, R5%, G5%, B5%, R6%, G6%, B6%, R7%, G7%, B7%, R8%, G8%, B8%, R9%, G9%, B9%, R10%, G10%, B10%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    
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
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Mid(TheText, fstlen% + seclen% + 1, thrdlen%)
    part4$ = Mid(TheText, fstlen% + seclen% + thrdlen% + 1, frthlen%)
    part5$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + 1, fithlen%)
    part6$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + 1, sixlen%)
    part7$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + 1, sevlen%)
    part8$ = Mid(TheText, fstlen% + seclen% + thrdlen% + frthlen% + fithlen% + sixlen% + sevlen% + 1, eightlen%)
    part9$ = Right(TheText, ninelen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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

dacolor$ = RGBtoHEX(OldColor)
RedX% = Val("&H" + Right(dacolor$, 2))
GreenX% = Val("&H" + Mid(dacolor$, 3, 2))
BlueX% = Val("&H" + Left(dacolor$, 2))
newred% = 255 - RedX%
newgreen% = 255 - GreenX%
newblue% = 255 - BlueX%
InverseColor = RGB(newred%, newgreen%, newblue%)

End Function

Function MultiFade(NumColors%, TheColors(), TheText$, Wavy As Boolean)

Dim WaveState%
Dim WaveHTML$
WaveState = 0

If NumColors < 1 Then
MsgBox "Error: Attempting to fade less than one color."
MultiFade = TheText
Exit Function
End If

If NumColors = 1 Then
blah$ = RGBtoHEX(TheColors(1))
redpart% = Val("&H" + Right(blah$, 2))
greenpart% = Val("&H" + Mid(blah$, 3, 2))
bluepart% = Val("&H" + Left(blah$, 2))
blah2 = RGB(bluepart%, greenpart%, redpart%)
blah3$ = RGBtoHEX(blah2)

MultiFade = "<Font Color=#" + blah3$ + ">" + TheText
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

For q% = 1 To NumColors
DaColors(q%) = RGBtoHEX(TheColors(q%))
Next q%

For w% = 1 To NumColors
RedList(w%) = Val("&H" + Right(DaColors(w%), 2))
GreenList(w%) = Val("&H" + Mid(DaColors(w%), 3, 2))
BlueList(w%) = Val("&H" + Left(DaColors(w%), 2))
Next w%

textlen% = Len(TheText)
Do: DoEvents
For F% = 1 To (NumColors - 1)
DaLens(F%) = DaLens(F%) + 1: textlen% = textlen% - 1
If textlen% < 1 Then Exit For
Next F%
Loop Until textlen% < 1
    
DaParts(1) = Left(TheText, DaLens(1))
DaParts(NumColors - 1) = Right(TheText, DaLens(NumColors - 1))
    
dastart% = DaLens(1) + 1

If NumColors > 2 Then
For E% = 2 To NumColors - 2
DaParts(E%) = Mid(TheText, dastart%, DaLens(E%))
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
        
    If Wavy = True Then
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
Sub PlayWav(wav As String)
X = sndPlaySound("" + Sound + "", 1):
     NoFreeze% = DoEvents()
End Sub
Function AIM_ChangeMainCaption(what As String)
AIMwindow& = FindWindow("_Oscar_BuddyListWin", vbNullString)
TheCaption& = SendMessageByString(AIMwindow&, WM_SETTEXT, 0, what)
End Function
Function AIM_ChangeIMCaption(what As String)
IMwindow& = FindWindow("AIM_IMessage", vbNullString)
TheCaption& = SendMessageByString(IMwindow&, WM_SETTEXT, 0, what)
End Function
Function AIM_ChangeChatCaption(what As String)
ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
TheCaption& = SendMessageByString(ChatWindow&, WM_SETTEXT, 0, what)
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
Function RGBtoHEX(RGB)

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

Function HTMLtoRGB(TheHTML$)

'converts HTML such as 0000FF to an
'RGB value like &HFF0000 so you can
'use it in the FadeByColor functions
If Left(TheHTML$, 1) = "#" Then TheHTML$ = Right(TheHTML$, 6)

RedX$ = Left(TheHTML$, 2)
GreenX$ = Mid(TheHTML$, 3, 2)
BlueX$ = Right(TheHTML$, 2)
rgbhex$ = "&H00" + BlueX$ + GreenX$ + RedX$ + "&"
HTMLtoRGB = Val(rgbhex$)
End Function
Function KillWindow(Window%)
KillWindow = SendMessageByNum(Window%, WM_CLOSE, 0, 0)
End Function

Function FadeFourColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, R4%, G4%, B4%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    
    Do: DoEvents
    fstlen% = fstlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    seclen% = seclen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    thrdlen% = thrdlen% + 1: textlen% = textlen% - 1
    If textlen% < 1 Then Exit Do
    Loop Until textlen% < 1
    
    part1$ = Left(TheText, fstlen%)
    part2$ = Mid(TheText, fstlen% + 1, seclen%)
    part3$ = Right(TheText, thrdlen%)
    
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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

Function FadeThreeColor(R1%, G1%, B1%, R2%, G2%, B2%, R3%, G3%, B3%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen% = Len(TheText)
    fstlen% = (Int(textlen%) / 2)
    part1$ = Left(TheText, fstlen%)
    part2$ = Right(TheText, textlen% - fstlen%)
    'part1
    textlen% = Len(part1$)
    For i = 1 To textlen%
        TextDone$ = Left(part1$, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen% * i) + B1, ((G2 - G1) / textlen% * i) + G1, ((R2 - R1) / textlen% * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
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
        
        If Wavy = True Then
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

Function FadeTwoColor(R1%, G1%, B1%, R2%, G2%, B2%, TheText$, Wavy As Boolean)

    Dim WaveState%
    Dim WaveHTML$
    WaveState = 0
    
    textlen$ = Len(TheText)
    For i = 1 To textlen$
        TextDone$ = Left(TheText, i)
        LastChr$ = Right(TextDone$, 1)
        ColorX = RGB(((B2 - B1) / textlen$ * i) + B1, ((G2 - G1) / textlen$ * i) + G1, ((R2 - R1) / textlen$ * i) + R1)
        colorx2 = RGBtoHEX(ColorX)
        
        If Wavy = True Then
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

