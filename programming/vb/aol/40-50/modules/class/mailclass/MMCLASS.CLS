VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MailClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Programmer: Chris (Crisis)
'Email Me At: memorexChris@hotmail.com

'Hello its me Chris (crisis) again. This is my class I made for aol4
'mails. If you get a byval error, try declaring your variables
'as the correct variable type. If you have good programming
'habits, you are probably doing this anyway. To work the
'ScanChat function do it like this.
'In the form declaration do this:
'
'Dim WithEvents MyClass as MailClass
'
'Without the withevents you won't get events
'In the form load up do this
'Set MyClass = New MailClass
'The myclass will appear in the object combobox. The event
'will appear in then procedure combobox. That's all there is
'to it.
Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Const WM_CLEAR = &H303
Private Const BM_SETCHECK = &HF1
Private Const VK_RETURN = &HD
Private Const VK_SPACE = &H20
Private Const WM_CHAR = &H102
Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const LB_GETCOUNT = &H18B
Private Const LB_GETTEXT = &H189
Private Const LB_GETTEXTLEN = &H18A
Private Const LB_SETCURSEL = &H186
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_SETTEXT = &HC
Private Const GW_HWNDPREV = 3
Private Const GW_HWNDLAST = 1
Private Const WM_CLOSE = &H10
Private Const SW_MINIMIZE = 6

Private m_MailList() As String 'The maillist vector
Private m_PeopleList As New Collection 'The people in the list.
Private m_MailSubjectPrefix As String 'Prefix string
Private m_MailText As String 'The body text of a mail
Private m_DeleteFwd As Boolean 'This sets whether to take off the fwd
Private m_BlindCarbonCopied As Boolean 'This sets whether the names are blind carbon copied
Private m_ErrorPeople As New Collection 'The people who were taken off due to errors

Private Const ClassErr As String = "AOL CLASS ERROR DEBUG STATEMENT" 'Debug Constant

Public Event NewChat(Text As String)    'The event used when a new chat line is found

Property Get MailBoxHandle() As Long
'Returns the mailbox handle.
    MailBoxHandle = FindWindowEx(GetMDI, 0, "AOL Child", UserName & "'s Online Mailbox")
End Property

Property Get PersonOnMM(Optional Person As String = Empty, Optional PersonIndex As Integer = Empty) As String
'Retrives a person from the internal collection object.
On Error Resume Next
Dim Counter As Variant, Checker As String

    If PersonIndex <> Empty Then
        MailListPerson = m_PeopleList.Item(PersonIndex)
        If Err.Number <> 0 Then
            Debug.Print Err.Number & ": " & Err.Description
            Err.Clear
            Debug.Print ClassErr
        End If
        Exit Property
    ElseIf Person <> Empty Then
        Checker = LCase(Person)
        For Each Counter In m_PeopleList
            If LCase(Counter) = Checker Then
                MailListPerson = m_PeopleList.Item(Counter)
                Exit Property
            End If
        Next Counter
        MailListPerson = ""
        Exit Property
    End If
End Property

Public Function AddPerson(Person As String, Optional CheckIfAlreadyThere As Boolean = True) As Integer
'Adds a person to the collection and returns a integer. If the
'integer is -1 then the operation was successful, if not then
'the integer represents the index of the person string that is
'already in the collection.
Dim Counter As Variant, Checker As String, Num As Integer

    If CheckIfAlreadyThere = True Then
        Checker = LCase(Person)
        For Each Counter In m_PeopleList
            Num = Num + 1
            If LCase(Counter) = Checker Then
                AddPerson = Num
                Exit Function
            End If
        Next Counter
        m_PeopleList.Add Person
        AddPerson = -1
        Exit Function
    Else
        m_PeopleList.Add Person
        AddPerson = -1
        Exit Function
    End If
End Function

Public Function RemovePerson(Person As String) As Integer
'Removes a person from the collection. RemovePerson returns
'a -1 on error. Any other number means the operation was a
'success and the number represents the index of the Person
'string that was in the collection.
Dim Counter As Variant, Checker As String, Num As Integer
On Error Resume Next

    Checker = LCase(Person)
    For Each Counter In m_PeopleList
        Num = Num + 1
        If LCase(Counter) = Checker Then
            m_PeopleList.Remove Num
            RemovePerson = Num
            Exit Function
        End If
    Next Counter
    GoTo ErrHandler
    Exit Function
    
ErrHandler:
    RemovePerson = -1
    Err.Clear
End Function

Property Get NumberOfPeopleOnMM() As Integer
Dim Holder As Long

    Holder = m_PeopleList.Count
    NumberOfPeopleOnMM = CInt(Holder)
End Property

Private Function UserName() As String
'This is a private function that returns
Dim Buffer As String, Win As Long
Dim Length As Long, Ret As Long

    Win = GetWindow(GetMDI, GW_CHILD)
        Do Until Win = 0
            Length = SendMessage(Win, WM_GETTEXTLENGTH, 0, 0)
            Buffer = String(Length, Chr(0))
            Ret = SendMessageByString(Win, WM_GETTEXT, Length + 1, Buffer)
            If Mid$(Buffer, 1, 9) = "Welcome, " Then
                UserName = Mid$(Buffer, 10)
                UserName = Left$(UserName, Len(UserName) - 1)
                Exit Function
            End If
            Win = GetWindow(Win, GW_HWNDNEXT)
        Loop
    UserName = ""
End Function

Private Function GetMDI() As Long
'This is a private function call that returns the aolmdi
'handle.
Dim AOL As Long, mdi As Long
    
    AOL = FindWindow("AOL Frame25", vbNullString)
    mdi = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
    GetMDI = mdi
End Function

Public Sub ClearNames()
'Clears the names from the collection.
Dim Counter As Integer

    For Counter = 1 To m_PeopleList.Count
        m_PeopleList.Remove 1
    Next Counter
End Sub

Public Sub LoadMailList()
'Loads the mail from the listbox. Resets all BlockMail calls.
Dim Length As Long, Counter As Long, Tree As Long
Dim Number As Long, Buffer As String
On Error GoTo ErrHandler

    Tree = GetTree
    If Tree = 0 Then Exit Sub
        Number = SendMessage(Tree, LB_GETCOUNT, 0, 0)
        ReDim m_MailList(Number)
            For Counter = 0 To Number - 1
                Length = SendMessage(Tree, LB_GETTEXTLEN, Counter, 0)
                Buffer = String$(Length, vbNullChar)
                SendMessageByString Tree, LB_GETTEXT, Counter, Buffer
                m_MailList(Counter) = Buffer
            Next Counter
            Exit Sub

ErrHandler:
Err.Clear
Debug.Print ClassErr
End Sub

Public Function RetrieveMailString(Index As Integer, Optional WithDate As Boolean = False, Optional WithSender As Boolean = False) As String
'Retrieves the mail string from the internal array of the mail
'list based on the index provided.
Dim Holder As String, NewPos As Integer, SaveDate As String
Dim NewIndex As Integer
    
    Holder = m_MailList(Index)
    If WithDate = False Then
        NewPos = InStr(1, Holder, vbTab)
        Holder = Mid$(Holder, NewPos + 1)
            If WithSender = False Then
                NewPos = InStr(1, Holder, vbTab)
                Holder = Mid$(Holder, NewPos + 1)
            End If
    Else
        If WithSender = False Then
            NewPos = InStr(1, Holder, vbTab)
            SaveDate = Left$(Holder, NewPos)
            NewPos = InStr(NewPos, Holder, vbTab)
            Holder = Mid$(Holder, NewPos + 1)
            Holder = SaveDate & vbTab & Holder
        End If
    End If
    RetrieveMailString = Holder
End Function

Property Get MailCount() As Integer
'This returns the mail count of the mail that was last loaded.
'This may or not be the actual mail count.
On Error Resume Next

    MailCount = UBound(m_MailList)
End Property

Property Let MailSubjectPrefix(Prefix As String)
'Sets a mail prefix to a mail subject.
'Ex:
'dim Test as new mmclass
'Test.mailsubjectprefix = "[Fire]"
'Then the mail would look like this
'Orginal: Any program
'New Thing: [Fire]Any program
    m_MailSubjectPrefix = Prefix
End Property

Property Get MailSubjectPrefix() As String
'Returns the mail prefix.
    MailSubjectPrefix = m_MailSubjectPrefix
End Property

Property Let MailText(Text As String)
'Sets the text when you mail something.
    m_MailText = Text
End Property

Property Get MailText() As String
'Returns the mail text you last entered in the mailtext call.
    MailText = m_MailText
End Property

Public Sub ForwardMail(ByVal MailIndex As Long, Optional RemoveErrorPeople As Boolean = True, Optional KeepAsNew As Boolean = True)
'Sends a mail based on your index. If RemoveErrorPeople = True
'then the forwardmail will automatically remove them from
'the peoplelist.
Dim Win As Long, ReadIcon As Long, WinRich As Long
Dim WinIcon As Long, WinStatic As Long, WinMail As Long
Dim Counter As Integer, WinDownload As Long, WinCombo As Long
Dim WinCount As Integer, WinStart As Long, Mail As Long
Dim RealMail As Long, AOLModal As Long, AOLModalIcon As Long
Dim AOLErrorBox As Long, AOLErrorView As Long, Buffer As String
Dim Length As Long, PeopleCounter As Variant, KeepNew As Long
Dim WinEdit As Long, WordCount As Long, LastWordCount As Long
Dim RichTrackBar As Long

    ReadIcon = FindWindowEx(MailBoxHandle, 0, "_AOL_Icon", vbNullString)
    KeepNew = FindWindowEx(MailBoxHandle, ReadIcon, "_AOL_Icon", vbNullString)
    KeepNew = FindWindowEx(MailBoxHandle, KeepNew, "_AOL_Icon", vbNullString)
    If ReadIcon = 0 Then GoTo ErrHandler
    SendMessage GetTree, LB_SETCURSEL, MailIndex, 0
    PostClick ReadIcon
        Do
        DoEvents
        Win = FindWindowEx(GetMDI, 0, "AOL Child", vbNullString)
            While Win
Retry:
                WinCount = 0
                WinStart = GetWindow(Win, GW_CHILD)
                WinStart = GetWindow(WinStart, GW_HWNDFIRST)
                Do
                    WinStart = GetWindow(WinStart, GW_HWNDNEXT)
                    WinCount = WinCount + 1
                Loop Until WinStart = 0
                WinDownload = FindWindowEx(Win, 0, "_AOL_Icon", vbNullString)
                WinIcon = WinDownload
                For Counter = 1 To 6
                        WinIcon = FindWindowEx(Win, WinIcon, "_AOL_Icon", vbNullString)
                Next Counter
                If WinDownload = WinIcon Then
                    If WinDownload <> 0 Then
                        Win = GetWindow(Win, GW_HWNDNEXT)
                        GoTo Retry
                    End If
                End If
                WinRich = FindWindowEx(Win, 0, "RICHCNTL", vbNullString)
                WinStatic = FindWindowEx(Win, 0, "_AOL_Static", vbNullString)
                WinCombo = FindWindowEx(Win, 0, "_AOL_Combobox", vbNullString)
                If WinIcon <> 0 And WinRich <> 0 And WinStatic <> 0 And WinCount = 18 Then
                    LockWindowUpdate WinRich
                        Do
                            WordCount = SendMessage(WinRich, WM_GETTEXTLENGTH, 0, 0)
                            Pause 1
                            LastWordCount = SendMessage(WinRich, WM_GETTEXTLENGTH, 0, 0)
                            RichTrackBar = FindWindowEx(WinRich, 0, "RICHTRACKBAR", vbNullString)
                        Loop Until WordCount = LastWordCount And RichTrackBar = 0
                    PostClick WinIcon
                        Do
                            Mail = FindWindowEx(GetMDI, 0, "AOL Child", vbNullString)
                            DoEvents
                                While Mail
                                    If FindWindowEx(Mail, 0, "_AOL_FontCombo", vbNullString) <> 0 Then
                                        RealMail = Mail
                                        Exit Do
                                    End If
                                    Mail = GetWindow(Mail, GW_HWNDNEXT)
                                Wend
                        Loop Until RealMail <> 0
                        Do
                            DoEvents
                            WinRich = FindWindowEx(RealMail, 0, "RICHCNTL", vbNullString)
                            WinEdit = FindWindowEx(RealMail, 0, "_AOL_Edit", vbNullString)
                        Loop Until WinRich <> 0 And WinEdit <> 0
AfterError:
                    SetPeople RealMail
                    SetSubject RealMail
                    SetMailText RealMail
                    WinIcon = GetWindow(RealMail, GW_CHILD)
                    WinIcon = GetWindow(WinIcon, GW_HWNDLAST)
                    WinIcon = GetWindow(WinIcon, GW_HWNDPREV)
                    WinIcon = GetWindow(WinIcon, GW_HWNDPREV)
                    WinIcon = GetWindow(WinIcon, GW_HWNDPREV)
                    WinIcon = GetWindow(WinIcon, GW_HWNDPREV)
                    WinIcon = GetWindow(WinIcon, GW_HWNDPREV)
                    If WinIcon = 0 Then GoTo AfterError
                    PostClick WinIcon
                        Do
                        DoEvents
                        AOLModal = FindWindow("_AOL_Modal", vbNullString)
                        AOLModalIcon = FindWindowEx(AOLModal, 0, "_AOL_Icon", vbNullString)
                        AOLErrorBox = FindWindowEx(GetMDI, 0, "AOL Child", "Error")
                        AOLErrorView = FindWindowEx(AOLErrorBox, 0, "_AOL_View", vbNullString)
                        If AOLErrorView <> 0 Then
                            If RemoveErrorPeople = False Then Exit Sub
                            Length = SendMessage(AOLErrorView, WM_GETTEXTLENGTH, 0, 0)
                            Buffer = String$(Length, vbNullChar)
                            SendMessageByString AOLErrorView, WM_GETTEXT, Length + 1, Buffer
                                For Each PeopleCounter In m_PeopleList
                                    If InStr(1, LCase(Buffer), LCase(PeopleCounter & " -")) > 0 Then
                                        m_ErrorPeople.Add PeopleCounter
                                        RemovePerson CStr(PeopleCounter)
                                    End If
                                Next PeopleCounter
                            SendMessage AOLErrorBox, WM_CLOSE, 0, 0
                            GoTo AfterError
                        End If
                        Loop Until AOLModalIcon <> 0
                            PostClick AOLModalIcon
                            LockWindowUpdate 0
                            PostMessage Win, WM_CLOSE, 0, 0
                                Do
                                DoEvents
                                Loop Until IsWindow(Win) = False
                            PostClick KeepNew
                            Exit Sub
                    Exit Sub
                End If
                Win = GetWindow(Win, GW_HWNDNEXT)
            Wend
        Loop
ErrHandler:
Exit Sub
End Sub

Public Sub OpenMailBox(Optional MinimizeIt As Boolean = True)
'Opens your mailbox.
Dim FirstCount As Long, SecondCount As Long
Dim Mail As Long

    If MailBoxHandle = 0 Then
        PostMessage GetIconOffToolBar, WM_LBUTTONDOWN, 0&, 0&
        PostMessage GetIconOffToolBar, WM_LBUTTONUP, 0&, 0&
            Do
                FirstCount = SendMessage(GetTree, LB_GETCOUNT, 0, 0)
                Pause 2
                SecondCount = SendMessage(GetTree, LB_GETCOUNT, 0, 0)
            Loop Until FirstCount = SecondCount And GetTree <> 0
            If MinimizeIt = True Then ShowWindow MailBoxHandle, SW_MINIMIZE
    End If
End Sub

Private Function GetIconOffToolBar() As Long
'Private function that gets the toolbar icon handle.
Dim Tool As Long, AOLTool As Long, AOL As Long, AOLIcon As Long
    
    AOL = FindWindow("AOL Frame25", vbNullString)
    Tool = FindWindowEx(AOL, 0, "AOL Toolbar", vbNullString)
    AOLTool = FindWindowEx(Tool, 0, "_AOL_Toolbar", vbNullString)
    AOLIcon = FindWindowEx(AOLTool, 0, "_AOL_Icon", vbNullString)
    GetIconOffToolBar = AOLIcon
End Function

Private Function GetTree() As Long
'Private function that gets the mail tree handle.
Dim TabPage As Long, TabControl As Long, Tree As Long
Dim hBox As Long

    hBox = MailBoxHandle
    TabControl = FindWindowEx(hBox, 0, "_AOL_TabControl", vbNullString)
    TabPage = FindWindowEx(TabControl, 0, "_AOL_TabPage", vbNullString)
    Tree = FindWindowEx(TabPage, 0, "_AOL_Tree", vbNullString)
    GetTree = Tree
End Function

Property Let DeleteFwd(YesOrNo As Boolean)
'This sets whether to remove the fwd: setting when sending
'mail.
    m_DeleteFwd = YesOrNo
End Property

Property Get DeleteFwd() As Boolean
'This returns the internal value stating if the Fwd is placed
'on outgoing mails.

    DeleteFwd = m_DeleteFwd
End Property

Private Sub PostClick(Handle As Long)
'Private sub that clicks a handle using postmessage.

    PostMessage Handle, WM_LBUTTONDOWN, 0&, 0&
    PostMessage Handle, WM_LBUTTONUP, 0&, 0&
End Sub

Private Sub SetSubject(MailHandle As Long)
'Private sub that sets the subject of a mail.
Dim MSubject As Long, MString As String, MLength As Long

    MSubject = FindWindowEx(MailHandle, 0, "_AOL_Edit", vbNullString)
    MSubject = FindWindowEx(MailHandle, MSubject, "_AOL_Edit", vbNullString)
    MSubject = FindWindowEx(MailHandle, MSubject, "_AOL_Edit", vbNullString)
    MLength = SendMessage(MSubject, WM_GETTEXTLENGTH, 0, 0)
    MString = String$(MLength, vbNullChar)
    SendMessageByString MSubject, WM_GETTEXT, MLength + 1, MString
        If m_DeleteFwd = True Then
            If Left$(MString, 4) = "Fwd:" Then
                MString = Mid$(MString, 6)
            End If
        End If
    SendMessageByString MSubject, WM_SETTEXT, 0, MailSubjectPrefix & MString
End Sub

Private Sub SetPeople(MailHandle As Long)
'Private sub that places the people into the Mail to slot.
Dim MPeople As Long, Counter As Variant, Final As String
    
    
    MPeople = FindWindowEx(MailHandle, 0, "_AOL_Edit", vbNullString)
    For Each Counter In m_PeopleList
        If BlindCarbonCopied = True Then
            Final = Final & "(" & Counter & ")"
        Else
            Final = Final & "," & Counter
            If Final = "," Then Final = ""
        End If
    Next Counter
    SendMessageByString MPeople, WM_SETTEXT, 0, Final
End Sub

Public Sub SetMailPrefs()
'You need to call this at least one time, so that the
'mail will send correctly.
Dim Combo As Long, AOL As Long, Tool As Long, AOLTool As Long
Dim ToolEdit As Long, Win As Long, WinIcon As Long
Dim WinImage As Long, WinPref As Long, WinPrefCheck As Long
Dim WinPrefIcon As Long

    AOL = FindWindow("AOL Frame25", vbNullString)
    Tool = FindWindowEx(AOL, 0, "AOL Toolbar", vbNullString)
    AOLTool = FindWindowEx(Tool, 0, "_AOL_Toolbar", vbNullString)
    Combo = FindWindowEx(AOLTool, 0, "_AOL_Combobox", vbNullString)
    ToolEdit = FindWindowEx(Combo, 0, "Edit", vbNullString)
    SendMessageByString ToolEdit, WM_SETTEXT, 0, "Mail Features"
    SendMessageByString ToolEdit, WM_CHAR, VK_SPACE, 0
    SendMessageByString ToolEdit, WM_CHAR, VK_RETURN, 0
        Do
            DoEvents
            Win = FindWindowEx(GetMDI, 0, "AOL Child", " Mailbox")
            WinImage = FindWindowEx(Win, 0, "_AOL_Image", vbNullString)
            WinIcon = FindWindowEx(Win, 0, "_AOL_Icon", vbNullString)
            WinIcon = FindWindowEx(Win, WinIcon, "_AOL_Icon", vbNullString)
            WinIcon = FindWindowEx(Win, WinIcon, "_AOL_Icon", vbNullString)
            WinIcon = FindWindowEx(Win, WinIcon, "_AOL_Icon", vbNullString)
            WinIcon = FindWindowEx(Win, WinIcon, "_AOL_Icon", vbNullString)
        Loop Until WinIcon <> 0 And WinImage <> 0
            Do
                DoEvents
                PostClick WinIcon
                Win = FindWindowEx(GetMDI, 0, "AOL Child", " Mail Preferences")
                WinImage = FindWindowEx(Win, 0, "_AOL_Image", vbNullString)
                If WinImage <> 0 Then SendMessage WinImage, WM_CLOSE, 0, 0
            Loop Until Win <> 0
                SendMessageByString Win, WM_SETTEXT, 0, "Setting Prefs..."
                WinIcon = GetWindow(Win, GW_CHILD)
                WinIcon = GetWindow(WinIcon, GW_HWNDLAST)
                WinIcon = GetWindow(WinIcon, GW_HWNDPREV)
                    Do
                    DoEvents
                    PostClick WinIcon
                    WinPref = FindWindow("_AOL_Modal", "Mail Preferences")
                    WinPrefCheck = FindWindowEx(WinPref, 0, "_AOL_CheckBox", vbNullString)
                    WinPrefIcon = FindWindowEx(WinPref, 0, "_AOL_Icon", vbNullString)
                    Loop Until WinPrefCheck <> 0 And WinPrefIcon <> 0
                        SendMessage WinPrefCheck, BM_SETCHECK, 1, 0
                        WinPrefCheck = GetWindow(WinPrefCheck, GW_HWNDNEXT)
                        SendMessage WinPrefCheck, BM_SETCHECK, 1, 0
                        WinPrefCheck = GetWindow(WinPrefCheck, GW_HWNDNEXT)
                        WinPrefCheck = GetWindow(WinPrefCheck, GW_HWNDNEXT)
                        WinPrefCheck = GetWindow(WinPrefCheck, GW_HWNDNEXT)
                        WinPrefCheck = GetWindow(WinPrefCheck, GW_HWNDNEXT)
                        SendMessage WinPrefCheck, BM_SETCHECK, 0, 0
                        PostClick WinPrefIcon
                        SendMessage Win, WM_CLOSE, 0, 0
End Sub

Property Let BlindCarbonCopied(YesOrNo As Boolean)
'This sets whether the names will be blind carbon copied.
    
    m_BlindCarbonCopied = YesOrNo
End Property

Property Get BlindCarbonCopied() As Boolean
'This returns the m_BlindCarbonCopied setting.
    
    BlindCarbonCopied = m_BlindCarbonCopied
End Property

Private Sub SetMailText(MailHandle As Long)
'Private sub that sets the mailtext into the mail.
Dim RICH As Long
    
    RICH = FindWindowEx(MailHandle, 0, "RICHCNTL", vbNullString)
    If m_MailText = "" Then Exit Sub 'Important*
    SendMessageByString RICH, WM_SETTEXT, 0, m_MailText
End Sub

Property Get NumberOfErroredPeople() As Integer
'Returns the number of errored people who have been
'taken off due to a error.

    NumberOfErroredPeople = m_ErrorPeople.Count
End Property

Public Sub ClearErrorPeople()
'This clears the people on the error collection.
Dim Counter As Variant

    For Each Counter In m_ErrorPeople
        m_ErrorPeople.Remove 1
    Next Counter
End Sub

Public Function RetrieveErrorPerson(Index As Integer) As String
'This returns a person from the errorpeople collection.
On Error Resume Next

    RetrieveErrorPerson = m_ErrorPeople.Item(Index)
End Function

Private Sub Pause(Seconds As Integer)
'This is a private sub that pauses for a duration.
Dim StartTime As Double

    StartTime = Timer
    Do
    DoEvents
    Loop Until Timer - Seconds >= StartTime
End Sub

Public Sub ScanChat()
'This scans the chat.
Static ChatString As String
Dim Length As Long, Buffer As String, Room As Long
Dim RICH As Long, Pos As Long, NewLine As String
Dim ParcedBuffer As String, ScreenName As String, Text As String
Dim NewPos As Long
Static OldChat As String

Retry:
    Room = FindRoom 'returns the room handle
    RICH = FindWindowEx(Room, 0, "RICHCNTL", vbNullString) 'returns the richcntl (chat area) handle
    If RICH = 0 Then Exit Sub 'if the room chat isn't there, then exit the function.
        Length = SendMessage(RICH, WM_GETTEXTLENGTH, 0, 0) 'returns the length of the chat
        Buffer = String$(Length, vbNullChar) 'Presets length of buffer.
        SendMessageByString RICH, WM_GETTEXT, Length + 1, Buffer 'retrieves chat in the string: buffer
        If ChatString <> Buffer Then 'if the static chatstring is different than the new string then chat has changed
            If ChatString = "" Then 'this is the 1st call.
                ChatString = Buffer 'makes the static = the current chat text
                Pause 0.5 'waits half a second.
                ScanChat 'calls itself.
                Exit Sub 'exits the sub
            End If
            'Something was different.
            ParcedBuffer = Mid$(Buffer, Len(ChatString) + 2) 'retrives the entire string after the static variable stopped. This will grab all new chat lines.
            Pos = InStr(1, ParcedBuffer, Chr(13)) 'this checks to see if the linebreak character is in the new text.
            If Pos = 0 Then 'if it isn't, that means that only 1 line has been added.
                If Right(ParcedBuffer, 1) = vbTab Then GoTo Retry 'if the chat didn't update all the way, retry it.
                If ParcedBuffer = "" Then 'this should probably be above, but this checks to see if any error happens.
                    Pause 1 'wait 1 second
                    ChatString = Buffer & Chr(13) 'the static variable
                    ScanChat 'recursion
                    Exit Sub 'exit the sub
                End If '***If ParcedBuffer = "" Then***
                    NewLine = ParcedBuffer 'makes the variable newline into the parced buffer since there is only 1 line anyway.
                    RaiseEvent NewChat(NewLine) 'raises the NewChat event
                    ChatString = Buffer 'chatstring static now equals th current chat text
                    ScanChat 'recursion
                    Exit Sub 'exit sub
            End If
Retry2:
            On Error Resume Next 'on a error resume next statment
            NewPos = InStr(1, ParcedBuffer, Chr(13)) 'checks for the new position of the line break
            NewLine = Mid$(ParcedBuffer, 1, NewPos - 1) 'parces the parcedbuffer string for the next line
            If NewLine = "" Then 'if the newline that was extracted was "" then do this:
                ChatString = ChatString & Chr(13) & NewLine 'updates chatstring
                ScanChat 'recursion
                Exit Sub 'exit sub
            End If
            If Right(NewLine, 1) = vbTab Then Exit Sub 'this is a error handler in case the chat does not fully show text
            RaiseEvent NewChat(NewLine) 'raises event
            ChatString = ChatString & Chr(13) & NewLine 'updates chatstring
            ScanChat 'recursion
            Exit Sub 'exit sub
        End If
End Sub

Private Function FindRoom() As Long
'This private function finds the aol chat room.
Dim List As Long, RICH As Long, Edit As Long, Icon As Long
Dim Room As Long, Combo As Long
    
    Room = FindWindowEx(GetMDI, 0, "AOL Child", vbNullString)
    While Room
        List = FindWindowEx(Room, 0, "_AOL_ListBox", vbNullString)
        Combo = FindWindowEx(Room, 0, "_AOL_ComboBox", vbNullString)
        RICH = FindWindowEx(Room, 0, "RICHCNTL", vbNullString)
        Edit = FindWindowEx(Room, RICH, "RICHCNTL", vbNullString)
        Icon = FindWindowEx(Room, 0, "_AOL_icon", vbNullString)
        If List <> 0 And Combo <> 0 And RICH <> 0 And Edit <> 0 And Icon <> 0 Then
            FindRoom = Room
            Exit Function
        End If
        Room = GetWindow(Room, GW_HWNDNEXT)
    Wend
    FindRoom = 0
End Function

Public Sub SendChat(Text As String)
'This is a sub that sends text to the chat room.
Dim Holder As Long

    Holder = FindWindowEx(FindRoom, 0, "RICHCNTL", vbNullString)
    Holder = FindWindowEx(FindRoom, Holder, "RICHCNTL", vbNullString)
    SendMessageByString Holder, WM_SETTEXT, 0, Text
    SendMessage Holder, WM_CHAR, VK_RETURN, 0
End Sub

Public Sub SendMail(Person As String, Subject As String, Text As String)
'Sends a mail on aol4.
Dim WriteIcon As Long, Win As Long, WinRich As Long
Dim WinEdit As Long, WinIcon As Long, X As Integer
Dim AOLModal As Long, Button As Long, Error As Long

    WriteIcon = GetIconOffToolBar
    WriteIcon = GetWindow(WriteIcon, GW_HWNDNEXT)
    PostClick WriteIcon
        Do
            DoEvents
            Win = FindWindowEx(GetMDI, 0, "AOL Child", "Write Mail")
            WinRich = FindWindowEx(Win, 0, "RICHCNTL", vbNullString)
            WinIcon = FindWindowEx(Win, 0, "_AOL_Icon", vbNullString)
            WinEdit = FindWindowEx(Win, 0, "_AOL_Edit", vbNullString)
        Loop Until WinRich <> 0 And WinIcon <> 0 And WinEdit <> 0
        Pause 1
        For X = 1 To 13
            WinIcon = FindWindowEx(Win, WinIcon, "_AOL_Icon", vbNullString)
        Next X
            If Person <> "" Then
                SendMessageByString WinEdit, WM_SETTEXT, 0, Person
            End If
            If Subject <> "" Then
                WinEdit = FindWindowEx(Win, WinEdit, "_AOL_Edit", vbNullString)
                WinEdit = FindWindowEx(Win, WinEdit, "_AOL_Edit", vbNullString)
                SendMessageByString WinEdit, WM_SETTEXT, 0, Subject
            End If
            SendMessageByString WinRich, WM_SETTEXT, 0, Text
                PostClick WinIcon
                    Do
                    AOLModal = FindWindow("_AOL_Modal", vbNullString)
                    Button = FindWindowEx(AOLModal, 0, "_AOL_Icon", vbNullString)
                    Error = FindWindowEx(GetMDI, 0, "AOL Child", "Error")
                    If Error <> 0 Then
                        SendMessage Error, WM_CLOSE, 0, 0
                        SendMessage Win, WM_CLOSE, 0, 0
                        Exit Sub
                    End If
                    Loop Until Button <> 0
                        PostClick Button
End Sub

Public Sub MinimizeMail()
'This minimizes the mail. It was requested.
    ShowWindow MailBoxHandle, SW_MINIMIZE
End Sub

Public Sub LoadListWithMailList(ListToAddTo As ListBox, Optional WithDate As Boolean = False, Optional WithSender As Boolean = False)
'This loads a list with the mailbox list. Requested.
Dim NextMail As Integer, MailString As String
    
    For NextMail = 0 To Me.MailCount - 1
        MailString = Me.RetrieveMailString(NextMail, WithDate, WithSender)
        ListToAddTo.AddItem MailString
    Next NextMail
End Sub

Public Sub LoadComboWithMailList(ComboToAddTo As ComboBox, Optional WithDate As Boolean = False, Optional WithSender As Boolean = False)
'This loads a combobox with the mailbox list. Requested.
Dim NextMail As Integer, MailString As String
    
    For NextMail = 0 To Me.MailCount - 1
        MailString = Me.RetrieveMailString(NextMail, WithDate, WithSender)
        ComboToAddTo.AddItem MailString
    Next NextMail
End Sub

Public Sub CloseMailBox()
'Closes the mail box.
    SendMessage Me.MailBoxHandle, WM_CLOSE, 0, 0
End Sub

