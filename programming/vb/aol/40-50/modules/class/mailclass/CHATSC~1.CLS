VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChatScanClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'America Online 4.0 ScanChat Module
'By: Chris (Crisis)
'Email Me At MemorexChris@hotmail.com
'This was orginally part of my massmailer class
'but was by far the most popular thing.

Option Explicit
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2

Public Event NewChat(Text As String)

Public Sub ScanChat()
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

Private Function GetMDI() As Long
'This is a private function call that returns the aolmdi
'handle.
Dim AOL As Long, mdi As Long
    
    AOL = FindWindow("AOL Frame25", vbNullString)
    mdi = FindWindowEx(AOL, 0, "MDIClient", vbNullString)
    GetMDI = mdi
End Function

Private Sub Pause(Seconds As Integer)
'This is a private sub that pauses for a duration.
Dim StartTime As Double

    StartTime = Timer
    Do
        DoEvents
    Loop Until Timer - Seconds >= StartTime
End Sub

