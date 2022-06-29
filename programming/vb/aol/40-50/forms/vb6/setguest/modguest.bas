Attribute VB_Name = "modguest"
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Const Cb_GetCount& = &H146
Public Const Cb_SetCursel& = &H14E
Public Function CmbCount(CmbBox As Long) As Long
    CmbCount& = SendMessageLong(CmbBox&, Cb_GetCount, 0&, 0&)
End Function
Public Sub GuestSetToGuest()
    Dim CmbBox As Long
    CmbBox& = FindWindowEx(FindSignOn&, 0&, "_AOL_ComboBox", vbNullString)
    Call PostMessage(CmbBox&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(CmbBox&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessageLong(CmbBox&, Cb_SetCursel, CmbCount(CmbBox&) - 1, 0&)
    Call PostMessage(CmbBox&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(CmbBox&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function GetCaption(WinHandle As Long) As String
    Dim buffer As String, TextLen As Long
    TextLen& = GetWindowTextLength(WinHandle&)
    buffer$ = String(TextLen&, 0&)
    Call GetWindowText(WinHandle&, buffer$, TextLen& + 1)
    GetCaption$ = buffer$
End Function
Public Function FindSignOn() As Long
Dim Counter As Long, AOLIcon As Long, AOLCombobox2 As Long, i As Long
Dim AOLStatic2 As Long, AOLEdit As Long, AOLCombobox As Long, AOLStatic As Long, RICHCNTL As Long
Dim AOLChild As Long, MDIClient As Long, AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
Next i&
AOLCombobox2& = FindWindowEx(AOLChild&, AOLCombobox&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (RICHCNTL& = 0& Or AOLStatic& = 0& Or AOLCombobox& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLCombobox2& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    Next i&
    AOLCombobox2& = FindWindowEx(AOLChild&, AOLCombobox&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If RICHCNTL& And AOLStatic& And AOLCombobox& And AOLEdit& And AOLStatic2& And AOLCombobox2& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindSignOn& = AOLChild&
    Exit Function
End If
End Function

