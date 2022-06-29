Attribute VB_Name = "Server"
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Sub MailToListFlash(TheList As ListBox)
    Dim AOL2 As Long, MDI2 As Long, fmail As Long, fList As Long
    Dim Count As Long, MyString As String, AddMails As Long
    Dim Aoltabcontrol As Long, aoltabpage As Long, aoltree As Long
    Dim sLength As Long, Spot As Long
    Dim donna As String
    AOL2& = FindWindow("AOL Frame25", vbNullString)
    MDI2& = FindWindowEx(AOL2&, 0&, "MDICLIENT", vbNullString)
    fmail& = FindWindowEx(MDI2&, 0&, "AOL Child", UserSN & "'s Filing Cabinet")
    If fmail& = 0& Then MsgBox "yoyoy"
    Aoltabcontrol& = FindWindowEx(fmail&, 0&, "_AOL_TabControl", vbNullString)
    aoltabpage& = FindWindowEx(Aoltabcontrol&, 0&, "_AOL_TabPage", vbNullString)
    fList& = FindWindowEx(aoltabpage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    MyString$ = String(255, 0)
    donna$ = 0 - ServerForm.List5.ListCount
    For AddMails& = 0 To Count& - 1
        DoEvents
        sLength& = SendMessage(fList&, LB_GETTEXTLEN, AddMails&, 0&)
        MyString$ = String(sLength& + 1, 0)
        Call SendMessageByString(fList&, LB_GETTEXT, AddMails&, MyString$)
        Spot& = InStr(MyString$, Chr(9))
        Spot& = InStr(Spot& + 1, MyString$, Chr(9))
        MyString$ = Right(MyString$, Len(MyString$) - Spot&)
        MyString$ = ReplaceString(MyString$, Chr(0), "")
        If MyString$ = "Mail Waiting To Be Sent" Then
            Exit Sub
        End If
        If MyString$ = "Mail" Then
            GoTo mea
        End If
        If MyString$ = "Incoming/Saved Mail" Then
            GoTo mea
        End If
        TheList.AddItem donna$ & "---)" & MyString$
        donna$ = donna$ + 1
mea:
    Next AddMails&
End Sub
Public Sub MailOpenEmailFlash(Index As Long)
Dim aol As Long, mdi As Long, fmail As Long, fList As Long, icon As Long
    Dim Count As Long, MyString As String, AddMails As Long
    Dim Aoltabcontrol As Long, aoltabpage As Long, aoltree As Long
    Dim sLength As Long, Spot As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDICLIENT", vbNullString)
    fmail& = FindWindowEx(mdi&, 0&, "AOL Child", UserSN & "'s Filing Cabinet")
    If fmail& = 0& Then Exit Sub
    Aoltabcontrol& = FindWindowEx(fmail&, 0&, "_AOL_TabControl", vbNullString)
    aoltabpage& = FindWindowEx(Aoltabcontrol&, 0&, "_AOL_TabPage", vbNullString)
    fList& = FindWindowEx(aoltabpage&, 0&, "_AOL_Tree", vbNullString)
    Count& = SendMessage(fList&, LB_GETCOUNT, 0&, 0&)
    If Count& < Index& Then Exit Sub
    Call SendMessage(fList&, LB_SETCURSEL, Index& + 2, 0&)
    icon& = Findopenbutton(fmail&)
    ServerForm.Timer6 = True
    Do:
        Call PostMessage(fList&, WM_KEYDOWN, VK_RETURN, 0&)
        StopIncomingText
        killexplorer
    Loop Until FindEmail <> 0&
    End Sub

Function WinTxt(ByVal hwnd As Integer)
Dim X As Integer
Dim Y As String
Dim z As Integer
X = SendMessage(hwnd, &HE, 0&, 0&)
Y = String(X + 1, " ")
z = SendMessageByString(hwnd, &HD, X + 1, Y)
WinTxt = Left(Y, X)
End Function
Public Function getit() As String
Dim AOLFrame25 As Long, MDIClient As Long, AOLChild As Long, RICHCNTL As Long
AOLFrame25& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame25&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Dim TheText As String, TL As Long
TL& = SendMessageLong(RICHCNTL&, WM_GETTEXTLENGTH, 0&, 0&)
TheText$ = String$(TL& + 1, " ")
Call SendMessageByString(RICHCNTL&, WM_GETTEXT, TL + 1, TheText$)
GetText$ = Left(TheText$, TL&)
End Function
Public Sub ForwardFlashMail(screennames As String, message As String, mailIndex As Long, DeleteFwd As Boolean)
Dim icon As Long, opensend As Long, EditTo As Long, EditCC As Long, EditSubject As Long, opensend2 As Long
Dim Rich As Long, Combo As Long, current As Variant, fCombo As Long, TempSubject As String
Dim aoerror As Long, aomodal As Long, current2 As Variant, killit As Variant, clicked As String
Dim aol As Long
killit = Timer
ServerForm.Label8.Caption = GetCaption(FindChatRoom)
Opendisshitup:
checkdeadmsg
checkflashmail
MailOpenEmailFlash (mailIndex&)
opensend2& = FindEmail
current = Timer
current2 = Timer
tryagain:
checkaol
checkflashmail
reply& = FindReplyButton(opensend2&)
current = Timer
Do While reply& = 0
    checkaol
    If FindEmail& = 0 Then Exit Sub
    StopIncomingText
Loop
Call ShowWindow(reply&, SW_HIDE)
clicked = 0
ServerForm.Timer6 = False
current2 = Timer
Do While Findsendbutton(FindForward) = 0
opendafuckinlist:
    checkaol
    checkdeadmsg
    'reply& = FindReplyButton(opensend2&&)
    'Call ShowWindow(reply&, SW_HIDE)
    If FindEmail& = 0 Then Exit Sub
    If Findsendbutton(FindForward) <> 0& Then GoTo vernon
    ServerForm.chattime.Caption = "1"
    icon& = FindForwardButton(opensend2&)
    Call PostMessage(icon&, WM_KEYDOWN, VK_SHIFT, 0&)
    Call PostMessage(icon&, WM_KEYDOWN, VK_RETURN, 0&)
    checkdeadmsg
    If Findsendbutton(FindForward) <> 0& Then GoTo vernon
    current = Timer
    Dim AOL2 As Long, MDI2 As Long, Status As Long, ourhandle As Long
    Do While Findsendbutton(FindForward) = 0&
        checkaol
        checkdeadmsg
        If FindEmail& = 0 Then Exit Sub
        If Timer - current > 5 Then Restartaol2
        If Findsendbutton(FindForward) <> 0& Then GoTo vernon
        AOL2& = FindWindow("AOL Frame25", vbNullString)  'locates the aol window
        MDI2& = FindWindowEx(AOL2&, 0, "MDIClient", vbNullString)
        Status& = FindWindowEx(MDI2&, 0&, "AOL Child", "Status")
        ourhandle& = FindWindowEx(MyHandle&, 0, "_AOL_View", vbNullString)
        Call SetWindowPos(Status&, 0, 0, 0, 0, 0, SWP_NOMOVE)
        If Findsendbutton(FindForward) <> 0& Then GoTo vernon
        GoTo opendafuckinlist
    Loop
Loop
vernon:
current = Timer
Do:
    aol& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
    If aol& = 0 Then Exit Sub
    checkaol
    checkdeadmsg
    checkdeadmsg
    DoEvents
    opensend& = FindForward
    EditTo& = FindWindowEx(opensend&, 0&, "_AOL_Edit", vbNullString)
    EditCC& = FindWindowEx(opensend&, EditTo&, "_AOL_Edit", vbNullString)
    EditSubject& = FindWindowEx(opensend&, EditCC&, "_AOL_Edit", vbNullString)
    Rich& = FindWindowEx(opensend&, 0&, "RICHCNTL", vbNullString)
    Combo& = FindWindowEx(opensend&, 0&, "_AOL_Combobox", vbNullString)
    fCombo& = FindWindowEx(opensend&, 0&, "_AOL_Fontcombo", vbNullString)
Loop Until EditTo& <> 0& Or FindEmail = 0
If FindEmail = 0 Then Exit Sub
Call SendMessageByString(EditTo&, WM_SETTEXT, 0&, screennames$)
TempSubject$ = Gettext2(EditSubject&)
TempSubject$ = Right(TempSubject$, Len(TempSubject$) - 5)
Call SendMessageByString(EditSubject&, WM_SETTEXT, 0&, TempSubject$)
Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
mike = 0
Do:
    checkaol
    checkdeadmsg
    icon& = Findsendbutton(FindForward)
    Call PostMessage(icon&, WM_KEYDOWN, VK_CONTROL, 0&)
    Call PostMessage(icon&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(icon&, WM_KEYUP, VK_CONTROL, 0&)
    Call PostMessage(icon&, WM_KEYUP, VK_RETURN, 0&)
    mike = mike + 1
Loop Until mike = 3
shithole:
Do:
    checkdeadmsg
    checkaol
    checkdeadmsg
    Call SendMessage(FindForward(), WM_DESTROY, 0&, 0&)
    Call SendMessage(FindForward(), WM_CLOSE, 0&, 0&)
    Call SendMessage(FindEmail&(), WM_DESTROY, 0&, 0&)
    Call SendMessage(FindEmail&(), WM_CLOSE, 0&, 0&)
Loop Until FindForward = 0 And FindEmail& = 0
checkaol
checkdeadmsg
killwait
checkaol
checkdeadmsg
ServerForm.chattime.Caption = "0"
Do: DoEvents
    AOL2& = FindWindow("AOL Frame25", vbNullString)  'locates the aol window
    MDI2& = FindWindowEx(AOL2&, 0, "MDIClient", vbNullString)
    aoerror& = FindWindowEx(MDI2&, 0&, "AOL Child", "Error")
    aomodal& = FindWindow("_AOL_Modal", vbNullString)
    If aomodal& <> 0 Then
        'killwin opensend2&2
        Exit Sub
    End If
    If aoerror& <> 0 Then
        killwin aoerror&
        Ban.BanName.AddItem screennames$
        'killwin opensend2&2
        Exit Do
    End If
Loop Until aoerror = 0&
End Sub
Public Function FindForwardButton(hwnd As Long) As Long
Dim aol As Long, mdi As Long, child As Long, icon As Long
Dim X
    For X = 0 To 7
        icon& = FindWindowEx(hwnd&, icon&, "_AOL_Icon", vbNullString)
    Next X
    FindForwardButton = icon&
End Function
Public Function FindReplyButton(hwnd As Long) As Long
Dim aol As Long, mdi As Long, child As Long, icon As Long
Dim X
    For X = 0 To 6
        icon& = FindWindowEx(hwnd&, icon&, "_AOL_Icon", vbNullString)
    Next X
    FindReplyButton = icon&
End Function
Public Function Findsendbutton(hwnd As Long) As Long
Dim aol As Long, mdi As Long, child As Long, icon As Long
Dim X
    For X = 0 To 15
        icon& = FindWindowEx(hwnd&, icon&, "_AOL_Icon", vbNullString)
    Next X
Findsendbutton = icon&
End Function
Public Function Findsendnow(hwnd As Long) As Long
Dim aol As Long, mdi As Long, child As Long, icon As Long
Dim X
    For X = 0 To 17
        icon& = FindWindowEx(hwnd&, icon&, "_AOL_Icon", vbNullString)
    Next X
Findsendnow = icon&
End Function
Public Function Findopenbutton(hwnd As Long) As Long
Dim aol As Long, mdi As Long, child As Long, icon As Long
Dim X
    For X = 0 To 2
        icon& = FindWindowEx(hwnd&, icon&, "_AOL_Icon", vbNullString)
    Next X
Findopenbutton = icon&
End Function
Public Sub RunMenuByString(SearchString As String)
    Dim aol As Long, aMenu As Long, mCount As Long
    Dim LookFor As Long, sMenu As Long, sCount As Long
    Dim LookSub As Long, sID As Long, sString As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    aMenu& = GetMenu(aol&)
    mCount& = GetMenuItemCount(aMenu&)
    For LookFor& = 0& To mCount& - 1
        sMenu& = GetSubMenu(aMenu&, LookFor&)
        sCount& = GetMenuItemCount(sMenu&)
        For LookSub& = 0 To sCount& - 1
            sID& = GetMenuItemID(sMenu&, LookSub&)
            sString$ = String$(100, " ")
            Call GetMenuString(sMenu&, sID&, sString$, 100&, 1&)
            If InStr(LCase(sString$), LCase(SearchString$)) Then
                Call SendMessageLong(aol&, WM_COMMAND, sID&, 0&)
                Exit Sub
            End If
        Next LookSub&
    Next LookFor&
End Sub
Public Sub killwait()
'kills the aol hourglass.
Dim aol As Long, aolmodal As Long, AOLGlyph As Long
Dim AOLStatic As Long, AOLIcon As Long, AolInstance As Long, current As Variant
aol& = FindWindowEx(0, 0&, "AOL Frame25", vbNullString)
'AOLInst = GetWindowWord(aol&, GWL_HINSTANCE)
'call createcursor(aolinst,
'Call SetCursor(vbArrow)
Call RunMenuByString("&About America Online")
current = Timer
Do: DoEvents
If Timer - current > 5 Then Exit Sub
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
AOLGlyph& = FindWindowEx(aolmodal&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(aolmodal&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(aolmodal&, 0&, "_AOL_Icon", vbNullString)
Loop Until aolmodal& <> 0& And AOLGlyph <> 0& And AOLStatic& <> 0& And AOLIcon& <> 0& '
Do: DoEvents
aolmodal& = FindWindowEx(0, 0&, "_AOL_Modal", vbNullString)
Call PostMessage(aolmodal&, WM_CLOSE, 0, 0&)
Loop Until aolmodal& = 0&
StopIncomingText
End Sub

Public Sub StopIncomingText()
Dim AOL2 As Long
AOL2& = FindWindow("AOL Frame25", vbNullString)
Call PostMessage(AOL2&, WM_KEYDOWN, VK_ESCAPE, 0&)
Call PostMessage(AOL2&, WM_KEYUP, VK_ESCAPE, 0&)
End Sub
Public Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
Public Sub Pause2(interval)



Dim CurrentTime As Long
CurrentTime& = Timer
Do While Timer - CurrentTime& < Val(interval): DoEvents
Loop
End Sub
Public Function AOLRoomView() As Long
    AOLRoomView = FindWindowEx(FindChatRoom, 0, "RICHCNTL", vbNullString)
End Function
Public Function AOLRoomEdit() As Long
    AOLRoomEdit = FindWindowEx(FindChatRoom, AOLRoomView, "RICHCNTL", vbNullString)
End Function
Public Function ConvertListToString(list As ListBox) As String



Dim i As Long, strString As String
For i& = 0& To list.ListCount - 1&
    If i& = (list.ListCount - 1&) Then
        strString$ = strString$ & list.list(i&)
    Else
        strString$ = strString$ & list.list(i&) & Chr(13) & Chr(10)
    End If
Next i&
ConvertListToString$ = strString$
End Function
Sub killwin(Windo)
   Call SendMessageLong(Windo, WM_DESTROY, 0&, 0&)
   Call SendMessageLong(Windo, WM_CLOSE, 0&, 0&)
End Sub
Public Sub FormNotOnTop(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Sub List_CopyListToList(ListA As ListBox, ListB As ListBox)
'revised in ver 3 for speed. it was bad before :)
On Error Resume Next
ListB.Clear
Dim X As Integer
For X = 0 To ListA.ListCount - 1
    ListB.AddItem ListA.list(X)
Next X
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
Public Function click(hwnd As Long)
Dim Size As RECT
Dim CurPos As POINTAPI
Call GetWindowRect(hwnd, Size)
Call GetCursorPos(CurPos)
Call SetCursorPos(Size.Left + (Size.Right - Size.Left) / 2, Size.Top + (Size.Bottom - Size.Top) / 2)
Call SendMessageByNum(hwnd, WM_LBUTTONDOWN, 0, 0)
Call SendMessageByNum(hwnd, WM_LBUTTONUP, 0, 0)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Function
Public Sub Restartaol2()
    Dim roomname As String, current As Variant, AOL2 As Long, clicked As Long, Restartdirectory As String, mdi As Long, signonscreen As Long, icon1 As Long, icon2 As Long, icon3 As Long, signonbut As Long, fmail As Long, Windo As Long, child As Long, X As Long, donna As String
    'Form1.Label7.Caption = "Restarting America Online"
    'Form1.Label7.Caption = "Scrolling off!"
    ServerForm.Timer1 = False
    ServerForm.Timer2 = False
    ServerForm.Timer3 = False
    ServerForm.Timer4 = False
    ServerForm.Timer5 = False
    ServerForm.Timer6 = False
    ServerForm.Timer8 = False
    roomname = GetCaption(FindChatRoom)
    SendChat lascii & " Restarting AOL BRB " & rascii
    Pause2 0.5
    Do: DoEvents
    AOL2& = FindWindow("AOL Frame25", vbNullString)
mea:
    killwin AOL2&
    Loop Until AOL2 = 0
    Pause2 3
openaolagain:
    'Form1.Label7.Caption = "Opening America Online"
    Shell "c:\America Online 6.0\aol.exe", vbNormal
    Do:
        AOL2& = FindWindow("AOL Frame25", vbNullString)
        Pause2 0.1
    Loop Until AOL2& <> 0&
    ServerForm.Timer1 = True
    ServerForm.Timer2 = True
    ServerForm.Timer3 = True
    ServerForm.Timer4 = True
    ServerForm.Timer5 = True
    ServerForm.Timer6 = True
    ServerForm.Timer8 = True
End Sub
Public Sub Restartaol()
    Dim roomname As String, current As Variant, AOL2 As Long, clicked As Long, Restartdirectory As String, mdi As Long, signonscreen As Long, icon1 As Long, icon2 As Long, icon3 As Long, signonbut As Long, fmail As Long, Windo As Long, child As Long, X As Long, donna As String
    'Form1.Label7.Caption = "Restarting America Online"
    'Form1.Label7.Caption = "Scrolling off!"
    ServerForm.Timer1 = False
    ServerForm.Timer2 = False
    ServerForm.Timer3 = False
    ServerForm.Timer4 = False
    ServerForm.Timer5 = False
    ServerForm.Timer6 = False
    ServerForm.Timer8 = False
openaolagain:
   'Form1.Label7.Caption = "Opening America Online"
    Shell "c:\America Online 6.0\aol.exe", vbNormal
    Do:
        AOL2& = FindWindow("AOL Frame25", vbNullString)
        Pause2 0.1
    Loop Until AOL2& <> 0&
End Sub
Sub checkrestart()

End Sub
Sub checkflashmail()
Dim aol As Long, mdi As Long, fmail As Long
Do
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
    fmail& = FindWindowEx(mdi&, 0&, "AOL Child", UserSN & "'s Filing Cabinet")
    If fmail& = 0 Then
            aol& = FindWindow("AOL Frame25", vbNullString)
            mdi& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
            fmail& = FindWindowEx(mdi&, 0&, "AOL Child", UserSN & "'s Filing Cabinet")
            Call RunMenuByString("Filing &Cabinet")
            Pause2 1
            aol& = FindWindow("AOL Frame25", vbNullString)
            mdi& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
            fmail& = FindWindowEx(mdi&, 0&, "AOL Child", UserSN & "'s Filing Cabinet")
            Call ShowWindow(fmail&, SW_MINIMIZE)
    End If
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0, "MDIClient", vbNullString)
    fmail& = FindWindowEx(mdi&, 0&, "AOL Child", UserSN & "'s Filing Cabinet")
Loop Until fmail& <> 0&
End Sub
Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
