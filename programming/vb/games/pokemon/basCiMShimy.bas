Attribute VB_Name = "Module1"
Public Function FindRoom() As Long
    Dim Aol As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim AOLIcon As Long, AOLStatic As Long
    Aol& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(Aol&, 0&, "MDIClient", vbNullString)
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
Public Sub ChatSend(strText As String)
    Dim intRoom As Long, intRich1 As Long, intRich2 As Long
    intRoom& = FindRoom&
    intRich1& = FindWindowEx(intRoom&, 0, "RICHCNTL", vbNullString)
    intRich2& = FindWindowEx(intRoom&, intRich1&, "RICHCNTL", vbNullString)
    Call SendMessageByString(intRich2, WM_SETTEXT, 0&, strText)
    Call SendMessageLong(intRich2, WM_CHAR, ENTER_KEY, 0&)
    Call SendMessageLong(intRich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
