Attribute VB_Name = "Module1"
Option Explicit
Declare Function SendMessageString Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As String) As Long
Declare Function SendMessageNumber Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal LParam As Long) As Long
Declare Function IsWindow Lib "User32" (ByVal hwnd As Long) As Long
Public Sub ChatSend(Chat As String)
    If Len(Chat) = 0 Then Exit Sub  'checks the length of the chat you are trying to send, it's its 0 there's no point in sending it
    Dim sCurText As String, dcTime As Variant
    Dim lRoom As Long
    If AOLChatRoom = 0 Then Exit Sub    'checks to see if a chatroom is open
    lRoom = AOLRoomEdit 'uses the sub to find the text box to type in
    sCurText = GetText(lRoom)   'gets the current text in the text box
    If Len(sCurText) Then   'gets the length of that text
        dcTime = Timer  'same as current = timer from before.....(check main form)
        Do Until Len(GetText(lRoom)) = 0    'begins a loop that will end when there is no text in the edit box
            If Timer - dcTime > 0.4 Then Exit Do    'if 4 tenths of a second has elapsed it skips it
            SetText lRoom, ""   'uses a sub to clear the edit box
        Loop    'loops
        DoEvents    'see main form
    End If  'ends the if
    SetText lRoom, Chat 'uses the sub to set the edit box to the text trying to be sent
    dcTime = Timer  'see above
    Do Until InStr(1, GetText(lRoom), Chat) 'begins a loop that loops until the the box is set to the text needed
        If Timer - dcTime > 0.4 Then    'if 4 tenths of a second elapses, it exits
            SetText lRoom, Chat 'uses the sub to set the edit box to the chat needed
            Exit Do 'exits the do
        End If  'ends the if
    Loop    'loops
    DoEvents    'see above
    SendChar lRoom, 13  'uses a sub to send enter the edit box
    dcTime = Timer  'see above
    Do Until InStr(1, GetText(lRoom), Chat) = 0 'begins a loop that will end when the text is gone
        If Timer - dcTime > 0.4 Then    'if 4 thenthes of second has elapsed then....
            SendChar lRoom, 13  'see above
            Exit Do 'exits the do loop
        End If  'ends the if
    Loop    'loop
    DoEvents    'see above
    If Len(sCurText) Then   'checks the variable set in the begining has a length, if so then that means there was chat in the text box before
        SetText lRoom, sCurText 'sets it back to the text that was there before
        dcTime = Timer  'see above
        Do Until InStr(1, GetText(lRoom), sCurText) 'begins a loop that will end when the text is put back
            If Timer - dcTime > 0.4 Then    'ya ya ya the 4/10 of second again
                SetText lRoom, sCurText 'see above
                Exit Do 'exits the do
            End If  'ends the if
        Loop    'loops
        DoEvents    'see above
    End If  'ends the if
End Sub
Public Function AOLHandle() As Long
    AOLHandle = FindWindowEx(0, 0, "AOL Frame25", vbNullString) 'locates aol
End Function
Public Function MDIHandle() As Long
    MDIHandle = FindWindowEx(AOLHandle, 0, "MDIClient", vbNullString)   'locates the mdiclient
End Function
Public Function AOLChatRoom() As Long
    Dim lMdi As Long, lRoom As Long
    Dim lAolRich As Long, lAolCombo As Long, lAolIcon As Long
    Dim lAolImage As Long, lAolStatic As Long, lAolStatic2 As Long, lAolGlyph As Long
    Dim lAolList As Long
    lMdi = MDIHandle    'calls the sub to find the mdiclient window
    lRoom = FindWindowEx(lMdi, 0, "AOL Child", vbNullString)    'locates an aol child
    Do Until lRoom = 0  'begins a loop
        lAolRich = FindWindowEx(lRoom, 0, "RICHCNTL", vbNullString) 'finds the specidfied window
        lAolCombo = FindWindowEx(lRoom, 0, "_AOL_Combobox", vbNullString) 'finds the specidfied window
        lAolIcon = FindWindowEx(lRoom, 0, "_AOL_Icon", vbNullString) 'finds the specidfied window
        lAolStatic = FindWindowEx(lRoom, 0, "_AOL_Static", vbNullString) 'finds the specidfied window
        lAolStatic2 = FindWindowEx(lRoom, lAolStatic, "_AOL_Static", vbNullString) 'finds the specidfied window
        lAolList = FindWindowEx(lRoom, 0, "_AOL_Listbox", vbNullString) 'finds the specidfied window
        If lAolRich <> 0 And _
        lAolCombo <> 0 And _
        lAolIcon <> 0 And _
        lAolStatic <> 0 And _
        lAolStatic2 <> 0 And _
        lAolList <> 0 Then  'this entire if statment is to make sure the room has all of the following to make sure it's a room that the sub has found and not another window
            AOLChatRoom = lRoom 'if they're all there then it knows it found the chatroom
            Exit Function   'exits the function
        Else    'otherwise
            lRoom = FindWindowEx(lMdi, lRoom, "AOL Child", vbNullString)    'check the next child
        End If  'ends the if
    Loop    'loop
End Function
Public Function AOLRoomView() As Long
    AOLRoomView = FindWindowEx(AOLChatRoom, 0, "RICHCNTL", vbNullString)    'finds the view on the room
End Function
Public Function AOLRoomEdit() As Long
    AOLRoomEdit = FindWindowEx(AOLChatRoom, AOLRoomView, "RICHCNTL", vbNullString)  'finds the edit of the room
End Function
Public Function GetText(hwnd As Long) As String
    Dim bResult As Long, lLength As Long
    Dim sBuffer As String
    bResult = IsWindow(hwnd)    'checks to see if the hwnd specified by the user is a window
    If bResult = 0 Then Exit Function   'if its' not a window it doesn't continue
    lLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0) 'gets the length of the hwnd
    sBuffer = String(lLength, 0)    'sets a string with the same length
    Call SendMessageString(hwnd, WM_GETTEXT, lLength + 1, sBuffer)  'replaces the string with the text of the window
    GetText = sBuffer   'sets the function = to the string
End Function
Public Sub SetText(hwnd As Long, Text As String)
    Dim bResult As Long
    bResult = IsWindow(hwnd)    'see above
    If bResult = 0 Then Exit Sub    'see above
    Call SendMessageString(hwnd, WM_SETTEXT, 0, Text)   'sets the windows text to the text specified
End Sub
Public Sub SendChar(hwnd As Long, Char As Byte)
    Dim bResult As Long
    bResult = IsWindow(hwnd)    'see above
    If bResult = 0 Then Exit Sub    'see above
    Call SendMessageNumber(hwnd, WM_CHAR, Char, 0)  'sends the specified character to the specified hwnd
End Sub
