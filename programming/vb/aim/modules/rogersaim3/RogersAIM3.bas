Attribute VB_Name = "RogersAIM3"
'Rogers AIM Bas 3.0
'For AIM 3.0
'By Rogers
'Thanks to:
'Pat or JK his spy provided easy coding when I was
'too lazy for all that typing ;)
'oirogers5@aol.com
'http://rogers.ownz.com
'http://oirogers5.cjb.net
'assorted subs taken from my module
'chrome32.bas
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
Public Function GetSNfromIM()
IM = FindWindow("AIM_IMessage", vbNullString)
IMcap = GetCaption(IM)
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
Public Sub SetText(Window As Long, Text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, Text$)
End Sub
Public Sub SendChat(What)
aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
wndateclass& = FindWindowEx(aimchatwnd&, wndateclass&, "wndate32class", vbNullString)
Call SetText(wndateclass, What)
Call SendMessageLong(wndateclass, WM_CHAR, ENTER_KEY, 0&)

End Sub
Sub KillAd()

blist = FindWindow("_Oscar_BuddyListWin", vbNullString)
ad = FindChildByClass(blist, "Ate32Class")
Call ShowWindow(Hide2%, SW_HIDE)
End Sub
Sub ChatInvite(Who, message, roomname)
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call ClickIcon(oscariconbtn)
Do
inv = FindWindow("aim_chatinvitesendwnd", vbNullString)
Loop Until inv <> 0
edit& = FindWindowEx(inv, 0&, "edit", vbNullString)
Call SetText(edit, Who)
edit& = FindWindowEx(inv, edit&, "edit", vbNullString)
Call SetText(edit, message)
edit& = FindWindowEx(inv, edit&, "edit", vbNullString)
Call SetText(edit, roomname)
aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
oscarstatic& = FindWindowEx(aimchatinvitesendwnd&, 0&, "_oscar_static", vbNullString)
oscarstatic& = FindWindowEx(aimchatinvitesendwnd&, oscarstatic&, "_oscar_static", vbNullString)
Call ClickIcon(oscarstatic)

End Sub
Sub UnKillAd()

blist = FindWindow("_Oscar_BuddyListWin", vbNullString)
ad = FindChildByClass(blist, "Ate32Class")
Call ShowWindow(Hide2%, SW_SHOW)
End Sub

Public Function IMText()
one = FindWindow("AIM_IMessage", vbNullString)
two = FindChildByClass(one, "Ate32Class")
IMText = GetText(two)
End Function
Public Sub SendIM(Who, What)
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call ClickIcon(oscariconbtn)
Do
DoEvents
Loop Until FindIM <> 0
aimimessage& = FindWindow("aim_imessage", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
wndateclass& = FindWindowEx(aimimessage&, wndateclass&, "wndate32class", vbNullString)
Call SetText(wndateclass, What)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscarpersistantcombo& = FindWindowEx(aimimessage&, 0&, "_oscar_persistantcombo", vbNullString)
edit& = FindWindowEx(oscarpersistantcombo&, 0&, "edit", vbNullString)
Call SetText(edit, What)
aimimessage& = FindWindow("aim_imessage", vbNullString)
oscariconbtn& = FindWindowEx(aimimessage&, 0&, "_oscar_iconbtn", vbNullString)
Call ClickIcon(oscariconbtn)

End Sub
Public Function FindIM() As Long
IM = FindWindow("AIM_IMessage", vbNullString)
FindIM = IM

End Function
Public Sub StopIt()
Do
DoEvents
Loop

End Sub
Public Function GetCaption(WindowHandle As Long) As String
    'From Dos
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function
Sub ClickIcon(icon)

Call SendMessage(icon, WM_LBUTTONDOWN, 0, 0&)
Call SendMessage(icon, WM_LBUTTONUP, 0, 0&)
End Sub
Function CountLines(Text As TextBox)
    Dim a
    a = SendMessage2(Text.hwnd, EM_GETLINECOUNT, 0, 0)
    CountLines = a
End Function
Function GetLine(Text1 As TextBox, Lineh As Integer)
Dim q As String

Dim m_sLineString As String * 1056
m_sLineString = Space$(1056)
q = SendMessage2(Text1.hwnd, EM_GETLINE, Lineh, ByVal m_sLineString)
GetLine = q

End Function
Public Sub Loadlistbox(thelist As ListBox, Directory As String)


    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
            thelist.AddItem MyString$
        Wend


        Close #1
    End Sub
Public Sub SaveListBox(thelist As ListBox, Directory As String)


    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.list(SaveList&)
    Next SaveList&


    Close #1
End Sub
Public Sub LoadComboBox(thelist As ComboBox, Directory As String)


    Dim MyString As String
    On Error Resume Next
    Open Directory$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
            thelist.AddItem MyString$
        Wend


        Close #1
    End Sub
Public Sub SaveComboBox(thelist As ComboBox, Directory As String)


    Dim SaveList As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveList& = 0 To thelist.ListCount - 1
        Print #1, thelist.list(SaveList&)
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
Public Sub ChatLink(site, word)
ChatLink = "<A HREF= " + Chr(34) + site + Chr(34) + ">" + word = "</A>"
End Sub
