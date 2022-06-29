Attribute VB_Name = "FlymanAim"
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
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
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
Public Sub SetText(Window As Long, text As String)
    Call SendMessageByString(Window&, WM_SETTEXT, 0&, text$)
End Sub
Sub Send_Text(String1$)
'this goes to a chat room
'Ex:Call Send_Text(Text1)
'or
'Ex:Call Send_Text("The Text here")
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
Butt& = FindWindowEx(aimchatwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call Click0r(Butt&)
End Sub
Sub WAVPlay(File)
'Ex: C:Wavs\Whatever.wav
    Dim SoundName As String
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(SoundName$, wFlags%)
End Sub
Sub AboutBas()
'This was made for aim 3.5
'I had 1person i think if that help me
'on something.
'Theres Like 70subs or more?
'If you need anything,or if there are bugs
'contact me:
'Mail: psei901428@hotmail.com
'Aim: Tourq
End Sub
Sub Clear_Chat(text$)
'Ex: Call Clear_Chat("")
Dim Parent As Long, Child1 As Long, Child2 As Long, child3 As Long, textset As Long
Parent& = FindWindow("AIM_ChatWnd", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "WndAte32Class", vbNullString)
Child2& = FindWindowEx(Child1&, 0&, "Ate32Class", vbNullString)
textset& = SendMessageByString(Child2&, WM_SETTEXT, 0, text$)
End Sub

Sub Click(TheIcon&)
'This was not coded by me,thanks to
'digitial
    Dim Klick As Long
    Klick& = SendMessage(TheIcon&, WM_LBUTTONDOWN, 0, 0&)
    Klick& = SendMessage(TheIcon&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Im_Direct(SN As String)
'thanks to trend again
'Ex: Call Im_Direct(text1)
Dim Parent As Long, Child1 As Long
Call gobar("aim:goim?screenname=" & SN$)
Parent& = FindWindow("AIM_IMessage", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call RunMenuByString(Parent, "Send IM I&mage")
End Sub
Sub Get_Files(SN$)
'thanks to trend for this
'Ex:Call Get_Files(Text1)
'or
'Ex:Call Get_Files("")
Call gobar("aim:getfile?screenname=" & SN$)
End Sub
Sub gobar(URL$)
Dim Parent As Long, Child1 As Long, textset As Long, Child2 As Long, textset2 As Long
Parent& = FindWindow("_Oscar_BuddyListWin", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "Edit", vbNullString)
textset& = SendMessageByString(Child1&, WM_SETTEXT, 0, URL$)
Child2& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call Click(Child2&)
textset2& = SendMessageByString(Child1&, WM_SETTEXT, 0, "Search The Web")
End Sub
Sub Clear_Im(text$)
'Ex Call Clear_Im("")
Dim Parent As Long, Child1 As Long, Child2 As Long, textset As Long, send As Long
Parent& = FindWindow("AIM_IMessage", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "WndAte32Class", vbNullString)
Child2& = FindWindowEx(Child1&, 0&, "Ate32Class", vbNullString)
textset& = SendMessageByString(Child2&, WM_SETTEXT, 0, text$)
End Sub
Sub Send_Im(SN$, text$)
'Ex: Call Send_Im(Text1,"Your Text")
'or
'Ex: Call Send_Im(Text1,Text2)
'or
'Ex: Call Send_Im("Tourq","Hey I am Useing your .bas")
'or
'Ex: Call Send_Im("Tourq",Text1)
Dim Parent As Long, Child1 As Long
Call gobar("aim:goim?screenname=" & SN$ & "&message=" & text$)
Parent& = FindWindow("AIM_IMessage", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call Click(Child1&)
End Sub
Sub RunMenuByString(Application, StringSearch)
'I got this from someone!

    Dim ToSearch As Integer, MenuCount As Integer, FindString
    Dim ToSearchSub As Integer, MenuItemCount As Integer, GetString
    Dim SubCount As Integer, MenuString As String, GetStringMenu As Integer
    Dim MenuItem As Integer, RunTheMenu As Integer
    
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
Sub File_Send(SN As String)
'Ex: Call File_Send(Text1)
'or
'Ex: Call File_Send("")
Dim Parent As Long, Child1 As Long
Call gobar("aim:goim?screenname=" & SN & "")
Parent& = FindWindow("AIM_IMessage", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call RunMenuByString(Parent, "Send &File")
End Sub
Sub DirectIm_Close(SN As String)
'Ex: Call DirectIm_Close(Text1)
'or
'Ex: Call DirectIm_Close("")
Dim Parent As Long, Child1 As Long
Call gobar("aim:goim?screenname=" & SN$)
Parent& = FindWindow("AIM_IMessage", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call RunMenuByString(Parent, "&Close IM Image")
End Sub


Function Blue_Link(Link$, Message$)
'Ex:Call Blue_Link("Http://www.deadbyte.com","My Page")
'or
'Ex:Call Blue_Link(Text1,Text2)
Send_Text "<a href=""" + (Text1) + """><font color=#0000ff>" + (Text2) + ""
End Function
Function Dot_Talk(strin As String) As String
'This was from chaos's .bas
'Ex:Moo$ = Dot_Talk(text1.text)
'Send_Text(moo$)

    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NumSpc As Integer, NewSent As String, Dotz As String
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + "•"
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    Dotz$ = NewSent$
    Dot_Talker = Dotz$
End Function
Function Space_Talk(strin As String) As String
'Ex: Moo$ = Space_Talk(text1.text)
'Send_Text(moo$)

    Dim NextChr As String, inptxt As String, lenth As Integer
    Dim NumSpc As Integer, NewSent As String, Spac As String
    
    Let inptxt$ = strin
    Let lenth% = Len(inptxt$)
    Do While NumSpc% <= lenth%
        Let NumSpc% = NumSpc% + 1
        Let NextChr$ = Mid$(inptxt$, NumSpc%, 1)
        Let NextChr$ = NextChr$ + " "
        Let NewSent$ = NewSent$ + NextChr$
    Loop
    Spac$ = NewSent$
    Space_Talk = Spac$
End Function
Sub Pause(interval)
'Ex: Pause(1.0)
    Dim Current
    
    Current = Timer
    Do While Timer - Current < Val(interval)
        DoEvents
    Loop
End Sub
Sub Im_Stamp_On()
'Ex: Call Im_Stamp_On
    Dim IMWin As Long
    
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMWin&, "Timestamp")
End Sub
Sub IM_Stamps_Off()
'Ex: Call IM_Stamps_Off
    Dim IMWin As Long
    
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMWin&, "Timestamp")
End Sub
Sub IM_Close()
'Ex: Call IM_Close
    Dim IMWin As Long

    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    If IMWin& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If
Start:
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call Win_Killwin(IMWin&)
End Sub
Sub IM_GetInfo()
'Ex: Call IM_GetInfo

    Dim IMWin As Long
    
    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(IMWin&, "Info")
End Sub
Sub Im_Link(who As String, Address As String, text As String, Closez As Boolean)
'Ex: Call Im_Link(text1,"www.deadbyte.com","Flymans Site"
'or
'Ex: Call Im_Link("Tourq",text1,"My Site")
    If Closez = True Then
        Call Send_Im(who$, "<A HREF=""" + Address$ + """>" + text$ + "", True)
    Else
        Call Send_Im(who$, "<A HREF=""" + Address$ + """>" + text$ + "", False)
    End If
End Sub
Sub Im_Bold(SN$, text$)
'Ex: Call Im_Bold("Tourq","<b>What to say?</b>")
'Warning: I could have done it the real way
'but it take more time
Dim Parent As Long, Child1 As Long
Call gobar("aim:goim?screenname=" & SN$ & "&message=" & text$)
Parent& = FindWindow("AIM_IMessage", vbNullString)
Child1& = FindWindowEx(Parent&, 0&, "_Oscar_IconBtn", vbNullString)
Call Click(Child1&)
End Sub
Sub Im_Talk()
'Ex: Call Im_Talk
'thanks go to blackout for this
    Dim talkb As Long, fullwindow As Long, fullbutton As Long, Klick As Long
    talkb& = FindWindow("AIM_IMessage", vbNullString)
    Call RunMenuByString(talkb&, "Connect to &Talk")
End Sub
Sub Im_Warn()
'Ex: Call Im_Warn
    Dim IMWin As Long, some As Long, Warn As Long, Click As Long

    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    some& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Warn& = FindWindowEx(IMWin&, some&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(Warn&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(Warn&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub IM_Block()
'Ex: Call Im_Block
    Dim IMWin As Long, some As Long, Warn As Long, Block As Long
    Dim Click As Long

    IMWin& = FindWindow("AIM_IMessage", vbNullString)
    some& = FindWindowEx(IMWin&, 0, "_Oscar_IconBtn", vbNullString)
    Warn& = FindWindowEx(IMWin&, some&, "_Oscar_IconBtn", vbNullString)
    Block& = FindWindowEx(IMWin&, Warn&, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(Block&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(Block&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Chat_Sounds_On()
'Ex: Call Chat_Sounds_On
'thanks to digital
    Dim ChatWindow As Long, ZeeWin As Long, PrefWin As Long
    Dim Buttin2 As Long, Buttin As Long, PlayMess As Long
    Dim Buttin1 As Long, Buttin22 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, PlaySend As Long
    Dim OKbuttin As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin&, "Button", vbNullString)
    PlayMess& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Call SendMessage(PlayMess&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlayMess&, WM_KEYUP, VK_SPACE, 0&)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin22& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin22&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    PlaySend& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Call SendMessage(PlaySend&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlaySend&, WM_KEYUP, VK_SPACE, 0&)

    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Chat_Sounds_Off()
'Ex: Call Chat_Sounds_Off
'thanks to digital
    Dim ChatWindow As Long, ZeeWin As Long, PrefWin As Long
    Dim Buttin2 As Long, Buttin As Long, PlayMess As Long
    Dim Buttin1 As Long, Buttin22 As Long, Buttin3 As Long
    Dim Buttin4 As Long, Buttin5 As Long, PlaySend As Long
    Dim OKbuttin As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call RunMenuByString(ChatWindow&, "&Edit Chat Preferences...")

    PrefWin& = FindWindow("#32770", "Buddy Chat")
    ZeeWin& = FindWindowEx(PrefWin&, 0, "#32770", vbNullString)
    Buttin& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin2& = FindWindowEx(ZeeWin&, Buttin&, "Button", vbNullString)
    PlayMess& = FindWindowEx(ZeeWin&, Buttin2&, "Button", vbNullString)
    Call SendMessage(PlayMess&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlayMess&, WM_KEYUP, VK_SPACE, 0&)
    Buttin1& = FindWindowEx(ZeeWin&, 0, "Button", vbNullString)
    Buttin22& = FindWindowEx(ZeeWin&, Buttin1&, "Button", vbNullString)
    Buttin3& = FindWindowEx(ZeeWin&, Buttin22&, "Button", vbNullString)
    Buttin4& = FindWindowEx(ZeeWin&, Buttin3&, "Button", vbNullString)
    Buttin5& = FindWindowEx(ZeeWin&, Buttin4&, "Button", vbNullString)
    PlaySend& = FindWindowEx(ZeeWin&, Buttin5&, "Button", vbNullString)
    Call SendMessage(PlaySend&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(PlaySend&, WM_KEYUP, VK_SPACE, 0&)

    OKbuttin& = FindWindowEx(PrefWin&, 0, "Button", vbNullString)
    Call SendMessage(OKbuttin&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(OKbuttin&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Sub Chat_Macrokill_Smile()
'Ex: Call Chat_Macrokill_Smile
Send_Text (":):-):-(;-):):-(:-):-(;-):):-):-(;-):):-):-(;-):):-):-(:):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):-):-(;-):):-):-(;-):):-):-(;-):):-(:-);-):):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-)")
Send_Text (":):-):-(;-):):-(:-):-(;-):):-):-(;-):):-):-(;-):):-):-(:):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):-):-(;-):):-):-(;-):):-):-(;-):):-(:-);-):):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-(;-):):-):-)")
End Sub
Sub Close_Chat()
'Ex: Call Chat_Close
    Dim ChatWindow As Long

    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    If ChatWindow& <> 0& Then
        GoTo Start
    Else
      Exit Sub
    End If
Start:
    ChatWindow& = FindWindow("AIM_ChatWnd", vbNullString)
    Call Win_Killwin(ChatWindow&)
End Sub
Sub chat_ignore(Person As String)
'Ex: Call chat_ignore(text1)
'or
'Ex: Call chat_ignore("Screenname")
Dim ChatRoom As Long, LopGet, MooLoo, Moo2
Dim name As String, NameLen, buffer As String
Dim TabPos, NameText As String, text As String
Dim mooz, Well As Integer, BuddyTree As Long
Person = LCase(Person)
ChatRoom = FindWindow("AIM_ChatWnd", vbNullString)
If ChatRoom <> 0 Then
Do
BuddyTree = FindWindowEx(ChatRoom, 0, "_Oscar_Tree", vbNullString)
Loop Until BuddyTree& <> 0
LopGet = SendMessage(BuddyTree&, LB_GETCOUNT, 0, 0)
For MooLoo = 0 To LopGet - 1
    Call SendMessageByString(BuddyTree, LB_SETCURSEL, MooLoo, 0)
    NameLen = SendMessage(BuddyTree, LB_GETTEXTLEN, MooLoo, 0)
    buffer = String(NameLen, "[]")
    Moo2 = SendMessageByString(BuddyTree, LB_GETtext, MooLoo, buffer)
    TabPos = InStr(buffer, Chr$(9))
    NameText = Right$(buffer, (Len(buffer) - (TabPos)))
    TabPos = InStr(NameText, Chr$(9))
    text = Right(NameText, (Len(NameText) - (TabPos)))
    text = Replace(text, " ", "")
    Person = Replace(Person, " ", "")
    If LCase(text) = LCase(Person) Then
     End If
Next MooLoo
End If
End Sub
Sub Clicks(item, clickmode As Integer)
Select Case clickmode
Case 1
Call SendMessageByNum(item, WM_LBUTTONDOWN, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONUP, 0, 0&)
Case 2
Call SendMessageByNum(item, WM_LBUTTONDOWN, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONUP, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONDOWN, 0, 0&)
Call SendMessageByNum(item, WM_LBUTTONUP, 0, 0&)
End Select
End Sub
Sub Bot_FakeProg()
Send_Text (Text1)
Send_Text (Text2)
Send_Text (Text3)
End Sub
Sub Bot_FakeVirii()
'Ex: Call Bot_FakeVirii
Send_Text ("Loging in to login.oscar.aol.com,Port 5190")
Send_Text ("Picking up Information")
Send_Text ("Sending Virus To Ever User")
Send_Text ("25% Done")
Send_Text ("86% Done")
Send_Text ("100% Done")
Send_Text ("Upload Of Virus Completed!")
End Sub
Private Sub Bot_Fighter()
'The Text1 and Text2 are were you put
'the people to fight are!
Send_Text ("People To Fight Are")
Pause (2.9)
Send_Text (Text1)
Pause (2.9)
Send_Text ("V.S.")
Pause (2.9)
Send_Text (Text2)
Pause (2.8)
Send_Text (Text1)
Send_Text ("Grabs Player By The nuts")
Pause (2.9)
Send_Text (Text2)
Pause (2.9)
Send_Text ("Likes it so doesn't do anything,but moans")
Pause (2.8)
Send_Text (Text1)
Pause (2.8)
Send_Text ("Cuts Dick off!")
Pause (2.9)
Send_Text (Text1)
Pause (2.9)
Send_Text ("Wins!")
End Sub

Sub Bot_AFK()
'How to do this:

'Add A New form

'Add 1 timer,set the timers interval to 60000

'Add 2 command buttons,Name Command 1 To: Start

'Name Command2,to : Stop

'Add 1 Textbox,in text box caption put: Reason for afk?

'Add 2 Labels,name Label1's caption to: 0

'On Label2 for that caption put: Mins

'In command1 put:Timer1.Enabled = True

'In command2 put:Timer1.Enabled = FalseLabel1.Caption = 0

'In Timer1 put:
'DoEvents
'send_text ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & l$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I>AFK Bot</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")
'send_text ("<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + "<~{-} " & l$ & "<FONT COLOR=" & Chr$(34) & "#0000FF" & Chr$(34) & "><I> I've ben gone for: ") + Label1 + (" Mins</I>" & aa$ & "<FONT COLOR=" & Chr$(34) & "#FF0000" & Chr$(34) & ">" + " {-}~>")
'send_text ("Oº°˜¨¨Reason Im Afk:")
'send_text (Text1)
'Label1.Caption = Val(Label1) + 1
End Sub
Sub Bot_Idle()
'How it works:
'New form
'Add 2 Command Buttons
'Add 1 Timer

'Name Command 1's caption to: Start
'Name Command 2's caption to: Stop

'In Command1 put:
'send_Text ("Whatever Prog Name Idle Bot")
'Timer1.Enabled = True

'In Command2 put:
'Timer1.Enabled = False
'send_Text ("Im Back Room!!")

'In timer1 put:
'DoEvents
'Send_Text ("Prog Name Idle Bot")
End Sub
Sub Macros_Pussy()
'Ex: Call Macros_Pussy
Send_Text ("   ;;;;;;;;;;  ;;;    ;;;  ;;;;;;;;;;  ;;;;;;;;;; ;;;     ;;;")
Send_Text ("   ¸¸¸¸¸¸¸;;;  ;;;    ;;;  ;;;¸¸¸¸¸¸¸  ;;;¸¸¸¸¸¸¸ ;;;¸¸¸¸¸;;;")
Send_Text ("   ;;;´´´´´´´  ;;;    ;;;  ´´´´´´´;;;  ´´´´´´´;;;  ´´´´´´´´´")
Send_Text ("   ;;;         ´;;;;;;;;´ ;;;;;;;;;;; ;;;;;;;;;;;    ;;;;;")
End Sub
Sub Macros_Tits()
'Ex: Call Macros_Tits
Send_Text ("  ;;;;;;;;;;;;;; ;;; ;;;;;;;;;;;;;; ;;;;;;;;;;")
Send_Text ("      ;;;;;;     ;;;     ;;;;;;     ;;;¸¸¸¸¸¸¸")
Send_Text ("      ;;;;;;     ;;;     ;;;;;;     ´´´´´´´;;;")
Send_Text ("      ;;;;;;     ;;;     ;;;;;;    ;;;;;;;;;;; ")
Send_Text ("                 Flyman")
End Sub
Sub Macros_BRB()
'Ex: Call Macros_BRB
Send_Text ("   ;;;;;;;;;;  ;;;;;;;;;;  ;;;;;;;;;;")
Send_Text ("   ¸¸¸¸¸¸¸;;´  ¸¸¸¸¸¸¸;;;  ¸¸¸¸¸¸¸;;´")
Send_Text ("   ;;;´´´´;;;  ;;;´´;;;;´  ;;;´´´´;;;")
Send_Text ("   ;;;;;;;;;;  ;;;   ´;;;¸ ;;;;;;;;;;")
Send_Text ("                       ´;;;¸")
Send_Text ("                         ´´´´")
End Sub
Sub Macros_Phish()
'Ex: Call Macros_Phish
Send_Text ("><>")
End Sub
Sub Macros_Flyman()
'Ex: Call Macros_Flyman
Send_Text ("  ¸;;;;;;;;;;;      ;;;   ;;; ¸;;;¸  ¸;;;; ;;;;;;;;¸ ;;;;¸¸ ;;;")
Send_Text ("  ;;;;;;;; ;;;      ´;;;;;;;; ;;;;;¸¸;;;;; ;;;;;;;;; ;;;´;;;;;;")
Send_Text ("  ;;;      ´;;;;;;;;   ;;;    ;;; ;;;; ;;; ;;;;;;;;; ;;;  ´;;;;")
End Sub
Sub Macros_Sperm()
Send_Text ("`·.¸.·´¯`·.¸.·O   Spërm")
End Sub
Sub Spammer()
Clear_Chat (text)
Send_Text (Text1)
Send_Text (Text1)
Send_Text (Text1)
Send_Text (Text1)
Send_Text (Text1)
Send_Text (Text1)
End Sub
Sub Form_Caption_AddLetter(Form1 As Form1, newletter As String, oldcaption As String)
'Ex:addletter Me, "W", Me.caption
   'addletter Me, "e", Me.caption
   'addletter Me, "l", Me.caption
   'addletter Me, "c", Me.caption
   'addletter Me, "o", Me.caption
   'addletter Me, "m", Me.caption
   'addletter Me, "e", Me.caption
    Dim total As Integer, spaces As Integer
    total = Len(temp)
    spaces = (form99.Width / 50) - (total)
    For X = spaces To Len(temp) Step -1
        form99.Caption = oldcaption & Space(X) & newletter
        DoEvents
        Next X
    End Sub


Public Sub Chat_AntiPunt()
'Ex: Call Chat_AntiPunt
On Error Resume Next
Dim aimchatwnd As Long
Dim wndateclass As Long
Dim ateclass As Long
Dim sHold As String
Dim txT As String
 aimchatwnd& = FindWindow("aim_chatwnd", vbNullString)
 wndateclass& = FindWindowEx(aimchatwnd&, 0&, "wndate32class", vbNullString)
 ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)
  txT$ = GetHalfText(ateclass&)
If InStr(1, LCase(txT$), LCase(">p<")) And InStr(1, LCase(txT$), LCase(">u<")) And InStr(1, LCase(txT$), LCase(">n<")) And InStr(1, LCase(txT$), LCase(">t<")) Then
    SetText ateclass&, "Pünt Was Found In Chat"
End If
If InStr(1, LCase(txT$), LCase("punter")) Then
    SetText ateclass&, "Pünt Was Found In Chat"
End If
If InStr(1, LCase(txT$), LCase(">e<")) And InStr(1, LCase(txT$), LCase(">r<")) And InStr(1, LCase(txT$), LCase(">r<")) And InStr(1, LCase(txT$), LCase(">o<")) Then
    SetText ateclass&, "Error String Was Found In Chat"
End If
If InStr(1, LCase(txT$), Chr(9)) Then 'looks for distort
    SetText ateclass&, "Distort Was Found In the Chat"
End If
If InStr(1, LCase(txT$), LCase(".clear")) Then 'looks for .clear in the room
    SetText ateclass&, "Chat cleared"
End If
End Sub
Public Sub Aim_AddBuddyList(List As ListBox)
'Ex: Call Aim_AddBuddyList(List1)
On Error Resume Next
Dim oscarbuddylistwin As Long
Dim oscartabgroup As Long
Dim oscartree As Long
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscartree& = FindWindowEx(oscartabgroup&, 0&, "_oscar_tree", vbNullString)
 List.Clear
  If oscartree& <> 0 Then
    LB2LB oscartree&, List   '
  End If
End Sub
Sub Chat_IpSniffer()
'programed by clash
'How this works:
'Add 2 Command Buttons
'Add 1 ListBox
'Add 1 Winsock

'In command 1 Put:
'Winsock1.LocalPort = 80
'Text1.Text = "<A href=""Http://" & Winsock1.LocalIP & """>Free Porn Pics!</A>"

'In Command2 Put:
'Send_Text (Text1)

'In The winsock put:
'List1.AddItem Winsock1.RemoteHostIP

'Name Command's 1 caption:
'&Generate

'Name Command's 2 caption:
'&Send Sniff!
End Sub
Sub Im_IpSniffer()
'coded by clash
'How it works:
'New Form
'Add 2 Command Buttons
'Add 1 Winsock Control
'Add 1 ListBox

'In command 1 Put:
'Winsock1.LocalPort = 80
'Text1.Text = "<A href=""Http://" & Winsock1.LocalIP & """>Free Porn Pics!</A>"

'In Command2 Put:
'Send_Im (Text1)

'In The winsock put:
'List1.AddItem Winsock1.RemoteHostIP

'Name Command's 1 caption:
'&Generate

'Name Command's 2 caption:
'&Send Sniff!
End Sub
Sub Form_Protection()
'How this works:
'Add 1 TextBox
'Add 2 CommandButtons
'Add 1 Timer

'In timer1 put:
'Text1.Text = "Can't Get in"

'Name Commands 1 Caption: Enter
'In Command1 Put:
'If Text1.Text = "Flyman" Then Secret.Show
'Timer1.Enabled = True

'Name Commands ' Caption: Exit
'In Command2 Put:
'Send_Text("Im A Loser,Im Not Leet,No Secret Are For me")
'MsgBox "Nice Try Buddy"
'End
End Sub
 Function Bot_Lamerizer(Nam As String)
'Thanks go to Blackout
'Ex: Call Bot_Lamerizer(text1)
'You gotta add a text box to enter
'who to lamerize
Call Send_Text("·  lamerizer found " & Nam & "  ·")
Pause (0.9)
Dim X As Integer
Dim lcse As String
Dim letr As String
Dim dis As String

For X = 1 To Len(Nam)
lcse$ = LCase(Nam)
letr$ = Mid(lcse$, X, 1)
If letr$ = "a" Then Let dis$ = "a-is for the animals your momma fucks": GoTo Dissem
If letr$ = "b" Then Let dis$ = "b-is for all the boys you love": GoTo Dissem
If letr$ = "c" Then Let dis$ = "c-is for the cunt you are": GoTo Dissem
If letr$ = "d" Then Let dis$ = "d-is for all the times your dissed": GoTo Dissem
If letr$ = "e" Then Let dis$ = "e-is for that egghead of yours": GoTo Dissem
If letr$ = "f" Then Let dis$ = "f-is for the friday nights you stay home": GoTo Dissem
If letr$ = "g" Then Let dis$ = "g-is for the girls who hate you": GoTo Dissem
If letr$ = "h" Then Let dis$ = "h-is for the ho your momma is": GoTo Dissem
If letr$ = "i" Then Let dis$ = "i-is for the idiotic dumbass you are": GoTo Dissem
If letr$ = "j" Then Let dis$ = "j-is for all the times you jerkoff to your dog": GoTo Dissem
If letr$ = "k" Then Let dis$ = "k-is for you self esteem that the cool kids killed": GoTo Dissem
If letr$ = "l" Then Let dis$ = "l-is for the lame ass you are": GoTo Dissem
If letr$ = "m" Then Let dis$ = "m-is for the many men you sucked": GoTo Dissem
If letr$ = "n" Then Let dis$ = "n-is for the nights you spent alone": GoTo Dissem
If letr$ = "o" Then Let dis$ = "o-is for the sex operation you had": GoTo Dissem
If letr$ = "p" Then Let dis$ = "p-is for the times people p on you": GoTo Dissem
If letr$ = "q" Then Let dis$ = "q-is for the queer you are": GoTo Dissem
If letr$ = "r" Then Let dis$ = "r-is for all the times i raped your sister": GoTo Dissem
If letr$ = "s" Then Let dis$ = "s-is for your lover Steve Case": GoTo Dissem
If letr$ = "t" Then Let dis$ = "t-is for the tits youll never see": GoTo Dissem
If letr$ = "u" Then Let dis$ = "u-is for your underwear hangin on the flagpole": GoTo Dissem
If letr$ = "v" Then Let dis$ = "v-is for the victories you'll never have": GoTo Dissem
If letr$ = "w" Then Let dis$ = "w-is for the 400 pounds you wiegh":  GoTo Dissem
If letr$ = "x" Then Let dis$ = "x-is for all the lamers who" & Chr(34) & "[x]'ed" & Chr(34) & " you online": GoTo Dissem
If letr$ = "y" Then Let dis$ = "y-is for the question of, y your even alive?": GoTo Dissem
If letr$ = "z" Then Let dis$ = "z-is for zero which is what you are":  GoTo Dissem

If letr$ = "1" Then Let dis$ = "1-is for how many inches your dick is": GoTo Dissem
If letr$ = "2" Then Let dis$ = "2-is for the 2 dollars you make an hour": GoTo Dissem
If letr$ = "3" Then Let dis$ = "3-is for the amount of men your girl takes at once": GoTo Dissem
If letr$ = "4" Then Let dis$ = "4-is for your mom bein a whore":  GoTo Dissem
If letr$ = "5" Then Let dis$ = "5-is for 5 times an hour you whack off": GoTo Dissem
If letr$ = "6" Then Let dis$ = "6-is for the years you been single": GoTo Dissem
If letr$ = "7" Then Let dis$ = "7-is for the times your girl cheated on you..with me": GoTo Dissem
If letr$ = "8" Then Let dis$ = "8-is for how many people beat the hell outta you today": GoTo Dissem
If letr$ = "9" Then Let dis$ = "9-is for how many boyfriends your momma has": GoTo Dissem
If letr$ = "0" Then Let dis$ = "0-is for the amount of girls you get": GoTo Dissem

Dissem:
Call Send_Text(dis$)

Pause (0.9)
Next X

End Function
Sub KillAd()
'Ex: Call KillAd
Dim oscarbuddylistwin&
Dim wndateclass&
Dim ateclass&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
wndateclass& = FindWindowEx(oscarbuddylistwin&, 0&, "wndate32class", vbNullString)
ateclass& = FindWindowEx(wndateclass&, 0&, "ate32class", vbNullString)

Call ShowWindow(ateclass&, SW_HIDE)
End Sub
Function Button()
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
Sub greets()
'Thanks go to:
'B0rf  :I forget,but im sure its important, :p
'Clash :Figured out the last chatline!
'Doce  :Got My Module on his page!
'Knk   :Has my module on his page!
'Quirk :Inspires me,lame eh? no not at all bitch!
'Trend :Sorry about last module
'Zb    :Nice person,listend to me bitch(gets alot of credit) ps.We hate mastafagg!!
End Sub
Sub Hide_Aim()
'Ex: Call Hide_Aim
    Dim BuddyList As Long, X As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    X& = ShowWindow(BuddyList&, SW_HIDE)
End Sub
Sub Hide_GoToBar()
'Ex: Call Hide_GoToBar
    Dim BuddyList As Long, STWbox As Long, GoButtin As Long
    Dim X  As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    X& = ShowWindow(STWbox&, SW_HIDE)
    X& = ShowWindow(GoButtin&, SW_HIDE)
End Sub
Sub Show_GoToBar()
'Ex: Show_GoToBar
    Dim BuddyList As Long, STWbox As Long, GoButtin As Long
    Dim X  As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    X& = ShowWindow(STWbox&, SW_SHOW)
    X& = ShowWindow(GoButtin&, SW_SHOW)
End Sub
Sub Show_Aim()
'Ex: Call Show_Aim
    Dim BuddyList As Long, X As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    X& = ShowWindow(BuddyList&, SW_SHOW)
End Sub
Sub GoTo_WebPage(Address As String)
'Thanks to digi
'Ex: Call GoTo_WebPage("www.deadbyte.com")
    Dim BuddyList As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    If BuddyList& <> 0& Then
    GoTo Start
    Else
    Exit Sub
    End If
Start:
    Dim STWbox As Long, SetAdd As Long
    Dim GoButtin As Long, Click As Long
    BuddyList& = FindWindow("_Oscar_BuddyListWin", vbNullString)
    STWbox& = FindWindowEx(BuddyList&, 0, "Edit", vbNullString)
    SetAdd& = SendMessageByString(STWbox&, WM_SETTEXT, 0, Address$)
    Pause 0.1
    GoButtin& = FindWindowEx(BuddyList&, 0, "_Oscar_IconBtn", vbNullString)
    Click& = SendMessage(GoButtin&, WM_LBUTTONDOWN, 0, 0&)
    Click& = SendMessage(GoButtin&, WM_LBUTTONUP, 0, 0&)
End Sub
Sub Mass_IM(lis As ListBox, txT As String)
'Ex: Call Mass_Im(List1,"Hi")
    If lis.ListCount = 0 Then
    Exit Sub
    Else
    End If
    Dim Moo
    For Moo = 0 To lis.ListCount - 1
    Call Send_Im(lis.List(Moo), txT, True)
    Next Moo
End Sub
Sub Im_Spammer()
'Ex: Call Im_Spammer(List1,"Hi")
    If lis.ListCount = 0 Then
    Exit Sub
    Else
    End If
    Dim Moo
    For Moo = 0 To lis.ListCount - 1
    Call Send_Im(lis.List(Moo), txT, True)
    Next Moo
End Sub
Sub Save_ListBox(Path As String, Lst As ListBox)
'Ex: Call Sace_ListBox("C:\Windows\Flyman\List.lst",List1)
    Dim Listz As Long
    On Error Resume Next
    Open Path$ For Output As #1
    For Listz& = 0 To Lst.ListCount - 1
    Print #1, Lst.List(Listz&)
    Next Listz&
    Close #1
End Sub
Public Sub Form_Move(TheForm As Form)
'Place in mousedown
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
Sub Macros_Glass() 's
Call Send_Text("    ¸.·²'°'²·.¸_¸.·²'°'²·.¸¸.-·~²°˜¨")
Call Send_Text("    `·.,¸¸,.·´  `·.,¸¸,.·´")
End Sub
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
Sub SendInvite(who$, Message$, Room$)
Dim aimchatinvitesendwnd&
Dim edit&
Dim oscariconbtn&
Call Click0r(ChatInviteButton)
aimchatinvitesendwnd& = FindWindow("aim_chatinvitesendwnd", vbNullString)
edit& = FindWindowEx(aimchatinvitesendwnd&, 0&, "edit", vbNullString)
Call SetText(edit, who$)
edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
Call SetText(edit, Message$)
edit& = FindWindowEx(aimchatinvitesendwnd&, edit&, "edit", vbNullString)
Call SetText(edit, Room$)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Butt& = FindWindowEx(aimchatinvitesendwnd&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call Click0r(Butt&)
End Sub
Sub IM_Respond(what$)
'Ex: Call IM_Respond(Text1)
'or
'Ex: Call IM_Respond("BRB")
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
Sub Voice_Chat(Person$)
'Thanks to blackout
Dim oscarbuddylistwin&
Dim oscartabgroup&
Dim oscariconbtn&
oscarbuddylistwin& = FindWindow("_oscar_buddylistwin", vbNullString)
oscartabgroup& = FindWindowEx(oscarbuddylistwin&, 0&, "_oscar_tabgroup", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Butt& = FindWindowEx(oscartabgroup&, oscariconbtn&, "_oscar_iconbtn", vbNullString)
Call Click0r(Butt&)
Do
DoEvents
talkwin& = FindWindow("#32770", vbNullString)
Butt& = FindWindowEx(talkwin&, 0&, "Button", vbNullString)
stat& = FindWindowEx(talkwin&, 0&, "Static", vbNullString)
Loop Until talkwin <> 0
Call SetText(stat&, Person$)
Call Click0r(Butt&)
End Sub
Sub form_mini()
Form1.WindowState = 1
End Sub
Sub form_max()
Form1.WindowState = 2
End Sub
Sub form_ontop()
Form1.WindowState = 0
End Sub
Sub Form_AlwaysOntop(the As Form)
SetWinOnTop = SetWindowPos(the.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
