Attribute VB_Name = "RogersAIM35"
'RogersAIM3.5 the FIRST module for AIM 3.5
'Created 11/11/99
'oirogers5@aol.com
'RogersFX on Dalnet in #project1
'http://rogersfx.spedia.net
'Peace: Argon, Funky, Snippa, Dos, Syber, Amber, KnK
'Gpx, Hider, Herb, Bud, Kaos, Yatzee, Chronoso, Mash
'Mi, Intx, Fri, all the rest of the crew
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETLINE = &HC4
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long)
Declare Function mciSendString Lib "MMSystem" Alias "mcisendstring" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal wReturnLength As Integer, ByVal hCallback As Integer) As Long
Const SND_SYNC = &H0
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
Sub Click0r(Icon1&)
    Call SendMessage(Icon1&, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessage(Icon1&, WM_LBUTTONUP, 0&, 0&)
End Sub
Function IMButton()
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
IMButton = oscariconbtn&
End Function
Function ChatInviteButton()
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
ChatInviteButton = oscariconbtn&
End Function
Function VoiceButton()
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
VoiceButton = oscariconbtn&
End Function
Function SendIM(Person$, message$)
Call Click0r(IMButton)
Dim aimimessage&
Do
DoEvents
IMwin& = FindWindow("aim_imessage", vbNullString)
Loop Until IMwin <> 0

Dim oscarpersistantcombo&
Dim edit&
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
edit& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
Call SetText(edit&, Person$)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
Call SetText(wndateclass&, message$)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
Call Click0r(oscariconbtn&)

End Function
Public Sub SetText(Window As Long, Text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub
Sub SendInvite(Who$, message$, Room$)
Dim aimchatinvitesendwnd&
Dim edit&
Dim oscariconbtn&
Call Click0r(ChatInviteButton)
aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
edit& = FindWindowEx(aimchatinvitesendwnd&, 0&, "edit", vbNullString)
Call SetText(edit, Who$)
edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
Call SetText(edit, message$)
edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
Call SetText(edit, Room$)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
butt& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call Click0r(butt&)
End Sub
Sub CloseDecline()
Dim X&
Dim Button&
Button& = FindWindow("button", vbNullString)
Call Click0r(Button&)

End Sub
Public Function ChatLink(site, word)
ChatLink = "<A HREF= " + Chr(34) + site + Chr(34) + ">" + word = "</A>"
End Function
Sub SendVoiceChat(Person$)
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
butt& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call Click0r(butt&)
Do
DoEvents
talkwin& = FindWindow("#32770", vbNullString)
butt& = FindWindowEx(talkwin&, 0&, "Button", vbNullString)
stat& = FindWindowEx(talkwin&, 0&, "Static", vbNullString)
Loop Until talkwin <> 0
Call SetText(stat&, Person$)
Call Click0r(butt&)
End Sub
Sub SendChat(String1$)
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
chatty& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
Call SetText(chatty, String1)
oscariconbtn& = FindWindowEx(aimchatwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
butt& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call Click0r(butt&)
End Sub
Function GetChatText()
Dim aimchatwnd&
Dim wndateclass&
Dim ateclass&
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
GetChatText = GetText(ateclass&)
End Function
Public Function GetSNfromIM()
IM& = FindWindow("AIM_IMessage", vbNullString)
IMcap = GetCaption(IM&)
q = InStr(IMcap, "-")
q = q - 1
b = Left(IMcap, q)
GetSNfromIM = b

End Function
Public Function UserSN()
Dim blist As Long
Dim cap As String
blist = FindWindow("_Oscar_BuddyListWin", vbNullString)
cap = GetCaption(blist)
MsgBox cap
pos = InStr(cap, "'")
MsgBox pos

SN = Left(cap, pos - 1)
UserSN = SN


End Function
Sub KillAd()
Dim oscarbuddylistwin&
Dim wndateclass&
Dim ateclass&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

Call ShowWindow(ateclass&, SW_HIDE)
End Sub
Sub ShowAd()
Dim oscarbuddylistwin&
Dim wndateclass&
Dim ateclass&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

Call ShowWindow(ateclass&, SW_SHOW)
End Sub
Public Function GetCaption(WindowHandle As Long) As String
    'From Dos
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function
Function CountLines(Text As Object)
    Dim a
    a = SendMessage(Text.hwnd, EM_GETLINECOUNT, 0, 0)
    CountLines = a
End Function
Function GetLine(Text1 As TextBox, Lineh As Integer)
Dim q As String

Dim m_sLineString As String * 1056
m_sLineString = Space$(1056)
q = SendMessage(Text1.hwnd, EM_GETLINE, Lineh, ByVal m_sLineString)
GetLine = q

End Function
Public Sub Loadlistbox(TheList As ListBox, Directory As String)


    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
            TheList.AddItem MyString$
        Wend


        Close #1
    End Sub
Public Sub SaveListBox(TheList As ListBox, Directory As String)


    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&


    Close #1
End Sub
Public Sub LoadComboBox(TheList As ComboBox, Directory As String)


    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
            TheList.AddItem MyString$
        Wend


        Close #1
    End Sub
Public Sub SaveComboBox(TheList As ComboBox, Directory As String)


    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To TheList.ListCount - 1
        Print #1, TheList.List(SaveList&)
    Next SaveList&


    Close #1
End Sub
Sub PrintIt(Text As TextBox)


    Printer.Print "" + Text.Text + Str(Printer.Page)
    Printer.NewPage
    Printer.Print "" + Text.Text + Str(Printer.Page)
    Printer.EndDoc
End Sub
Public Function SaveIt(Data, File)
    Open File For Output As #1
    Write #1, Data
    Close #1

End Function
Public Function LoadIt(File, Data)

    Dim a As String
    Open File For Input As 1
    a = Input(LOF(1), 1)
    Close 1
    Data = a

End Function
Sub WAVPlay(File)

    Dim SoundName As String
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(SoundName$, wFlags%)
End Sub
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function
Sub RespondIM(what$)
Dim aimimessage&
Dim wndateclass&
Dim ateclass&
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
textboxer& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
send0r& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
Call SetText(textboxer&, what$)
Call Click0r(send0r&)

End Sub
Public Function GetText(WindowHandle As Long) As String
    'thanks dos
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetText$ = buffer$
End Function
