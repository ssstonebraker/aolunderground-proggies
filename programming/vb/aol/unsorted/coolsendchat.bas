Attribute VB_Name = "coolsendchat"
Sub SendChat(Text As String)
If aol_findroom <> 0 Then

Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim richcntl As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, Text)

aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
    Call SendMessageLong(richcntl, WM_CHAR, 13, 0&)
    ClickIt aolicon
    EnterKey aolicon
    Call SendMessageLong(richcntl, WM_CHAR, 13, 0&)
    ClickIt aolicon
    EnterKey aolicon
    End If
End Sub
Sub Chatsend(Text As String)
If aol_findroom <> 0 Then

Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim richcntl As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, Text)

aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
    Call SendMessageLong(richcntl, WM_CHAR, 13, 0&)
    ClickIt aolicon
    EnterKey aolicon
    Call SendMessageLong(richcntl, WM_CHAR, 13, 0&)
    ClickIt aolicon
    EnterKey aolicon
    End If
End Sub
Sub sendchat40(Text As String)
Dim aolframe As Long
Dim mdiclient As Long
Dim aolchild As Long
Dim richcntl As Long
Dim aolicon As Long
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)


aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
richcntl = FindWindowEx(aolchild, 0&, "richcntl", vbNullString)
richcntl = FindWindowEx(aolchild, richcntl, "richcntl", vbNullString)
Call SendMessageByString(richcntl, WM_SETTEXT, 0&, Text)

    Call SendMessageLong(richcntl, WM_CHAR, 13, 0&)
    ClickIt aolicon
    EnterKey aolicon
End Sub
Public Function FindRoomx() As Long
    Dim aol As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim aolicon As Long, aolstatic As Long
    On Error Resume Next
    aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    aolstatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And aolstatic& <> 0& Then
        FindRoomx& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            aolstatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And aolstatic& <> 0& Then
                FindRoomx& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindRoomx& = child&
End Function
