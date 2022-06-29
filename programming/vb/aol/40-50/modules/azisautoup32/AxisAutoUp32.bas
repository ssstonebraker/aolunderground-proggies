Attribute VB_Name = "AxisAutoUp32"
''''''''''''''''''''''''''''''''''''''
'    AxisAutoUp32.bas made simple by '
'                 axis               '
'*************************************
'   mail me:axis@asiansonly.net      *
'   www.asiansimplicity.cjb.net      *
'   visit my site and vote!!!!!      *
''''''''''''''''''''''''''''''''''''''
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_CHAR = &H102

Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9

Function AutoUpChat(per As String)
'look at the sub carefully there is a % sign so the % is
'included and per equals what number in percents you want
'it to minimize so i would call like this to give you an
'example alright here it goes
'call autoupchat(1)
'that will minimize on 1% when upwindow is shown
'heh so easy lata-axis
    Dim UpWindow As Long
    UpWindow& = FindWindow("_AOL_MODAL", "File Transfer - " + per$ + "%") 'u can change it to any percent
    If UpWindow& = 0 Then
        Exit Function
    End If
    Call ShowWindow(UpWindow&, SW_HIDE)
    Call ShowWindow(UpWindow&, SW_MINIMIZE)
'you can change the bottom to what timer you want
Form1.Timer1.Enabled = False
Form1.Timer2.Enabled = True

End Function

Sub UnUpchat()

    Dim UpWindow As Long
    UpWindow& = FindWindow("_AOL_MODAL", vbNullString)
    If UpWindow& = 0 Then
    Exit Sub
    End If
    Call ShowWindow(UpWindow&, SW_HIDE)
    Call ShowWindow(UpWindow&, SW_RESTORE)
End Sub
Sub About()
'sup sup i just made this bas for a friend who wanted to
'learn how to make an autoupchat which i never made be4
'so i just wanted to see if i can make one so i could
'well its pretty easy heh well yeah thats about it
'and you know what peepz deeznuts in ya mouth bish :D
'lata axis
End Sub


Public Function GetText(WinHandle As Long) As String
    Dim blah As String, TextLen As Long
   
   Let TextLen& = SendMessage(WinHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    blah$ = String(TextLen&, 0&)
    Call SendMessageByString(WinHandle&, WM_GETTEXT, TextLen& + 1, blah$)
    GetText$ = blah$
End Function







Public Sub ChatSend(txt As String)
    Dim chat As Long, ChatRich As Long, RichText As String
    If txt = "" Then Exit Sub
    chat& = FindRoom&
    ChatRich& = ChatSendBox&
    If chat& = 0& Then Exit Sub
    If ChatRich& = 0& Then Exit Sub
    RichText = GetText(ChatSendBox&)
    If RichText = "" Then
        Call SendMessageByString(ChatRich&, WM_SETTEXT, 0&, txt)
    Else
        Call SendMessageByString(ChatRich&, WM_SETTEXT, 0&, "")
        Call SendMessageByString(ChatRich&, WM_SETTEXT, 0&, txt)
    End If
sendtext:
    Call SendMessageByString(ChatRich&, WM_CHAR, 13, 0&)
    If Not GetText(ChatSendBox&) = "" Then
        DoEvents
        GoTo sendtext
    End If
    Call SendMessageByString(ChatRich&, WM_SETTEXT, 0&, RichText)
End Sub

Function ChatSendBox() As Long
Dim Room As Long
    Dim Rich As Long
    Room& = FindRoom&
    Rich& = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
    ChatSendBox = FindWindowEx(Room, Rich, "RICHCNTL", vbNullString)
End Function

Public Function FindRoom() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            AOLIcon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And AOLIcon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindRoom& = child&
End Function
Sub Status()
Dim lchild2 As Long, lchild1 As Long, lchild3 As Long

lchild1 = FindWindowEx(lParent, 0, "_AOL_Modal", vbNullString)
uppercent$ = GetText(lchild1&)
'ChatSend (uppercent$)

lchild2 = FindWindowEx(lchild1, 0, "_AOL_Static", vbNullString)
upstat$ = GetText(lchild2&)
'ChatSend (upstat$)

lchild3 = FindWindowEx(lchild1, lchild2, "_AOL_Static", vbNullString)
upstat1$ = GetText(lchild3&)
'ChatSend (upstat1$)





End Sub
