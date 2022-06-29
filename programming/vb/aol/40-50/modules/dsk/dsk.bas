Attribute VB_Name = "dsk"
' ******** dsk.bas ********

' aol version made on: 5.0

' i have used some sendkeys and appactivate,
' but to me, that savs alot of time, now
' i didn't use it in all the functions, i only
' used it like about 8 times out of all of them.

' this bas is just a taste of what i can do,
' like a beta. the next bas i make, will be never
' before seen options, some options in here are
' never before seen! and i bet people will
' copy them.

' NOTE: some things might not work yet due to the
' demands of this bas. the ghosting, and mabye the
' aolkeyword2 is a little glitchy, the buddy room
' invite still works, but that's a little glitchy
' also, and the signonguest that will be done in my
' next bas. and a few other things are glitchy at the moment.

' peace,
' dsk

' dsk info

' web page: http://dskrealm.cjb.net
' comment: still working on dsk's realm
' e-mail: dsk@dskrealm.cjb.net
' aim: i am dsk







Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const LB_SETCURSEL = &H186
Public Const WM_MOVE = &HF012
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_CHAR = &H102
Public Const VK_SPACE = &H20
Public Const VK_RETURN = &HD
Public Const WM_CLOSE = &H10
Public Const SW_HIDE = 0
Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_SYSCOMMAND = &H112
Public Const WM_LBUTTONUP = &H202
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1
Public Function FindWWW() As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "World Wide Web")
End Function
Public Function FindWindowsToolBar() As Long
Dim SysTabControl As Long
Dim MSTaskSwWClass As Long
Dim ShellTrayWnd As Long
ShellTrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
MSTaskSwWClass& = FindWindowEx(ShellTrayWnd&, 0&, "MSTaskSwWClass", vbNullString)
SysTabControl& = FindWindowEx(MSTaskSwWClass&, 0&, "SysTabControl32", vbNullString)
End Function
Public Function FindAOL() As Long
Dim Counter As Long
Dim AOLMMI As Long
Dim AOLToolbar As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLMMI& = FindWindowEx(AOLFrame&, 0&, "_AOL_MMI", vbNullString)
Do While (Counter& <> 100&) And (MDIClient& = 0& Or AOLToolbar& = 0& Or AOLMMI& = 0&): DoEvents
    AOLFrame& = FindWindowEx(AOLFrame&, AOLFrame&, "AOL Frame25", vbNullString)
    MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
    AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
    AOLMMI& = FindWindowEx(AOLFrame&, 0&, "_AOL_MMI", vbNullString)
    If MDIClient& And AOLToolbar& And AOLMMI& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOL& = AOLFrame&
    Exit Function
End If
End Function
Public Function FindSignOn() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLCombobox2 As Long
Dim i As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLCombobox As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
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
Do While (Counter& <> 100&) And (AOLCombobox& = 0& Or AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLCombobox2& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
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
    If AOLCombobox& And AOLStatic& And AOLEdit& And AOLStatic2& And AOLCombobox2& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindSignOn& = AOLChild&
    Exit Function
End If
End Function
Public Function FindWelcome() As Long
Dim Counter As Long
Dim AOLIcon9 As Long
Dim RICHCNTL9 As Long
Dim AOLIcon8 As Long
Dim RICHCNTL8 As Long
Dim AOLIcon7 As Long
Dim RICHCNTL7 As Long
Dim AOLIcon6 As Long
Dim RICHCNTL6 As Long
Dim AOLIcon5 As Long
Dim RICHCNTL5 As Long
Dim AOLIcon4 As Long
Dim RICHCNTL4 As Long
Dim AOLIcon3 As Long
Dim RICHCNTL3 As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLIcon As Long
Dim AOLButton As Long
Dim RICHCNTL As Long
Dim AOLListbox As Long
Dim JGAOLSNP As Long
Dim i As Long
Dim AOLCheckbox As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 18&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
JGAOLSNP& = FindWindowEx(AOLChild&, 0&, "JGAOLSNP", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLButton& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
RICHCNTL3& = FindWindowEx(AOLChild&, RICHCNTL2&, "RICHCNTL", vbNullString)
RICHCNTL3& = FindWindowEx(AOLChild&, RICHCNTL3&, "RICHCNTL", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
RICHCNTL4& = FindWindowEx(AOLChild&, RICHCNTL3&, "RICHCNTL", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
RICHCNTL5& = FindWindowEx(AOLChild&, RICHCNTL4&, "RICHCNTL", vbNullString)
AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
RICHCNTL6& = FindWindowEx(AOLChild&, RICHCNTL5&, "RICHCNTL", vbNullString)
AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
RICHCNTL7& = FindWindowEx(AOLChild&, RICHCNTL6&, "RICHCNTL", vbNullString)
AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon7&, "_AOL_Icon", vbNullString)
RICHCNTL8& = FindWindowEx(AOLChild&, RICHCNTL7&, "RICHCNTL", vbNullString)
AOLIcon8& = FindWindowEx(AOLChild&, AOLIcon7&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon8& = FindWindowEx(AOLChild&, AOLIcon8&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL9& = FindWindowEx(AOLChild&, RICHCNTL8&, "RICHCNTL", vbNullString)
For i& = 1& To 4&
    RICHCNTL9& = FindWindowEx(AOLChild&, RICHCNTL9&, "RICHCNTL", vbNullString)
Next i&
AOLIcon9& = FindWindowEx(AOLChild&, AOLIcon8&, "_AOL_Icon", vbNullString)
For i& = 1& To 5&
    AOLIcon9& = FindWindowEx(AOLChild&, AOLIcon9&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLCheckbox& = 0& Or JGAOLSNP& = 0& Or AOLListbox& = 0& Or RICHCNTL& = 0& Or AOLButton& = 0& Or AOLIcon& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or RICHCNTL3& = 0& Or AOLIcon3& = 0& Or RICHCNTL4& = 0& Or AOLIcon4& = 0& Or RICHCNTL5& = 0& Or AOLIcon5& = 0& Or RICHCNTL6& = 0& Or AOLIcon6& = 0& Or RICHCNTL7& = 0& Or AOLIcon7& = 0& Or RICHCNTL8& = 0& Or AOLIcon8& = 0& Or RICHCNTL9& = 0& Or AOLIcon9& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
    For i& = 1& To 18&
        AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    Next i&
    JGAOLSNP& = FindWindowEx(AOLChild&, 0&, "JGAOLSNP", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLButton& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 4&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    RICHCNTL3& = FindWindowEx(AOLChild&, RICHCNTL2&, "RICHCNTL", vbNullString)
    RICHCNTL3& = FindWindowEx(AOLChild&, RICHCNTL3&, "RICHCNTL", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    RICHCNTL4& = FindWindowEx(AOLChild&, RICHCNTL3&, "RICHCNTL", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    RICHCNTL5& = FindWindowEx(AOLChild&, RICHCNTL4&, "RICHCNTL", vbNullString)
    AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
    RICHCNTL6& = FindWindowEx(AOLChild&, RICHCNTL5&, "RICHCNTL", vbNullString)
    AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
    AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
    RICHCNTL7& = FindWindowEx(AOLChild&, RICHCNTL6&, "RICHCNTL", vbNullString)
    AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
    AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon7&, "_AOL_Icon", vbNullString)
    RICHCNTL8& = FindWindowEx(AOLChild&, RICHCNTL7&, "RICHCNTL", vbNullString)
    AOLIcon8& = FindWindowEx(AOLChild&, AOLIcon7&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon8& = FindWindowEx(AOLChild&, AOLIcon8&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL9& = FindWindowEx(AOLChild&, RICHCNTL8&, "RICHCNTL", vbNullString)
    For i& = 1& To 4&
        RICHCNTL9& = FindWindowEx(AOLChild&, RICHCNTL9&, "RICHCNTL", vbNullString)
    Next i&
    AOLIcon9& = FindWindowEx(AOLChild&, AOLIcon8&, "_AOL_Icon", vbNullString)
    For i& = 1& To 5&
        AOLIcon9& = FindWindowEx(AOLChild&, AOLIcon9&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLCheckbox& And JGAOLSNP& And AOLListbox& And RICHCNTL& And AOLButton& And AOLIcon& And RICHCNTL2& And AOLIcon2& And RICHCNTL3& And AOLIcon3& And RICHCNTL4& And AOLIcon4& And RICHCNTL5& And AOLIcon5& And RICHCNTL6& And AOLIcon6& And RICHCNTL7& And AOLIcon7& And RICHCNTL8& And AOLIcon8& And RICHCNTL9& And AOLIcon9& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindWelcome& = AOLChild&
    Exit Function
End If
End Function
Public Function FindPreparing2SwitchSNs() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindPreparingToS& = AOLModal&
    Exit Function
End If
End Function
Public Function FindBuddyList() As Long
Dim Counter As Long
Dim AOLStatic6 As Long
Dim AOLIcon4 As Long
Dim AOLStatic5 As Long
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim AOLIcon2 As Long
Dim AOLStatic3 As Long
Dim AOLIcon As Long
Dim AOLStatic2 As Long
Dim AOLListbox As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLListbox& = 0& Or AOLStatic2& = 0& Or AOLIcon& = 0& Or AOLStatic3& = 0& Or AOLIcon2& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0& Or AOLStatic5& = 0& Or AOLIcon4& = 0& Or AOLStatic6& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
    AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
    If AOLStatic& And AOLListbox& And AOLStatic2& And AOLIcon& And AOLStatic3& And AOLIcon2& And AOLStatic4& And AOLIcon3& And AOLStatic5& And AOLIcon4& And AOLStatic6& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindBuddyList& = AOLChild&
    Exit Function
End If
End Function
Public Function FindSignOff() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLCombobox2 As Long
Dim i As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLCombobox As Long
Dim AOLStatic As Long
Dim RICHCNTL As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
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
    FindSignOff& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAOLConnect() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLGlyph2 As Long
Dim AOLStatic As Long
Dim i As Long
Dim AOLGlyph As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
For i& = 1& To 3&
    AOLGlyph& = FindWindowEx(AOLModal&, AOLGlyph&, "_AOL_Glyph", vbNullString)
Next i&
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLGlyph2& = FindWindowEx(AOLModal&, AOLGlyph&, "_AOL_Glyph", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLGlyph& = 0& Or AOLStatic& = 0& Or AOLGlyph2& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
    For i& = 1& To 3&
        AOLGlyph& = FindWindowEx(AOLModal&, AOLGlyph&, "_AOL_Glyph", vbNullString)
    Next i&
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph2& = FindWindowEx(AOLModal&, AOLGlyph&, "_AOL_Glyph", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLGlyph& And AOLStatic& And AOLGlyph2& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindSignOnConnect& = AOLModal&
    Exit Function
End If
End Function
Public Function GetText(hwnd As Long) As String
Dim TextLen As Long
Dim hWndTxt As String
TextLen& = SendMessage(hwnd&, WM_GETTEXTLENGTH, 0&, 0&)
hWndTxt$ = String(TextLen&, 0&)
Call SendMessageByString(hwnd&, WM_GETTEXT, TextLen& + 1&, hWndTxt$)
GetText$ = hWndTxt$
End Function
Public Function GetWindowCaption(hwnd As Long) As String
Dim CaptionLen As Long
Dim WndCaption As String
CaptionLen& = SendMessage(hwnd&, WM_GETTEXTLENGTH, 0&, 0&)
WndCaption$ = String(CaptionLen&, 0&)
Call SendMessageByString(hwnd&, WM_GETTEXT, CaptionLen& + 1&, WndCaption$)
GetWindowCaption$ = WndCaption$
End Function
Public Function FindChat() As Long
Dim Counter As Long
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim AOLListbox As Long
Dim AOLStatic3 As Long
Dim AOLImage As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim RICHCNTL As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 6&
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or AOLImage& = 0& Or AOLStatic3& = 0& Or AOLListbox& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 6&
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And RICHCNTL& And AOLCombobox& And AOLIcon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And AOLImage& And AOLStatic3& And AOLListbox& And AOLStatic4& And AOLIcon3& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindChat& = AOLChild&
    Exit Function
End If
End Function
Public Function SendToRoom(What As String) As Long
Dim i As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "")
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, What)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Function SendToRoomAIM(What As String) As Long

' this third sendtoroom chat will place a time
' before everything you say.
' example.

' Dafil:        ( 3:49:45 PM ): HEHE
' Dafil:        ( 3:59:24 PM ): Got To Go

' it will always place the time in parentheses
' before what ever you say.

Dim i As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "")
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "</b></i></u>( <b>" & Time & "</b></i></u> ): " & What)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Function SendToRoomAIM2(What As String) As Long

' this third sendtoroom chat will place a date
' before everything you say.
' example.

' Dafil:        ( 10/17/99 ): HEHE
' Dafil:        ( 10/17/99 ): Got To Go

' it will always place the date in parentheses
' before what ever you say.

Dim i As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "")
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "</b></i></u>( <b>" & Date & "</b></i></u> ): " & What)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Function SendToRoomAIM3(What As String) As Long

' this third sendtoroom chat will place a time
' and date before everything you say.
' example.

' Dafil:        ( 3:42:00 PM : 10/17/99 ): HEHE
' Dafil:        ( 3:54:52 PM : 10/17/99 ): Got To Go

' it will always place the time and date in
' parentheses before what ever you say.

Dim i As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "")
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "</b></i></u>( <b>" & Time & " : " & Date & "</b></i></u> ): " & What)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Function TimeNow(Interval) As Long
Dim CurrentTime As Long
CurrentTime& = Timer
Do While Timer - CurrentTime& < Val(Interval): DoEvents
Loop
End Function
Public Function ChangeCaption(hwnd As Long, NCaption As String)
' with this function, u can change any of the find window functions
' in this bas

Call SendMessageByString(hwnd&, WM_SETTEXT, 0&, NCaption$)
End Function
Public Function CloseWindow(Window As String)
Call SendMessage(Window, WM_CLOSE, 0&, 0&)
End Function
Public Function HideWin(Window As String)
Call ShowWindow(Window, SW_HIDE)
End Function
Public Function MaximizeWin(Window As String)
Call ShowWindow(Window, SW_MAXIMIZE)
End Function
Public Function MinimizeWin(Window As String)
Call ShowWindow(Window, SW_MINIMIZE)
End Function
Public Function NormalShowSizeWin(Window As String)
Call ShowWindow(Window, SW_SHOWNORMAL)
End Function
Public Function ShowWin(Window As String)
Call ShowWindow(Window, SW_SHOW)
End Function
Public Function SetFocusWin(Window As String)
Call SetFocusAPI(Window)
End Function
Public Function FindMailAttatch() As Long
Dim Counter As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLTree As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLTree& = FindWindowEx(AOLModal&, 0&, "_AOL_Tree", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLTree& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLTree& = FindWindowEx(AOLModal&, 0&, "_AOL_Tree", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 4&
        AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And AOLTree& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindMailAttatch& = AOLModal&
    Exit Function
End If
End Function
Public Function SendToRoom2(What As String)

' this second sendtoroom function will copy
' the text you were typing, sends what it
' needs to send, and then paste back what
' you was typing to continue.

Dim Room As Long
Dim ChatBox As Long
Dim Oldtext As String
Dim NewText As String
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Room = FindChat
If Room Then
ChatBox = FindWindowEx(Room, 0&, "RICHCNTL", vbNullString)
ChatBox = FindWindowEx(Room, ChatBox, "RICHCNTL", vbNullString)
Oldtext = GetText(ChatBox)
Call SendMessageByString(ChatBox, &HC, 0, "")
Call SendMessageByString(ChatBox, &HC, 0, What)
Do
DoEvents
NewText = GetText(ChatBox)
Loop Until NewText = What
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
NewText = GetText(ChatBox)
Loop Until NewText = ""
Call SendMessageByString(ChatBox, &HC, 0, Oldtext)
End If
End Function
Public Function ChatTextBox() As Long
Dim ChatBox As Long
ChatBox = FindWindowEx(FindChat, 0&, "RICHCNTL", vbNullString)
ChatBox = FindWindowEx(FindChat, ChatBox, "RICHCNTL", vbNullString)
ChatTextBox = ChatBox
End Function
Public Function FindMail() As Long
Dim Counter As Long
Dim AOLStatic10 As Long
Dim AOLIcon7 As Long
Dim AOLStatic9 As Long
Dim AOLIcon6 As Long
Dim AOLStatic8 As Long
Dim AOLIcon5 As Long
Dim AOLStatic7 As Long
Dim AOLIcon4 As Long
Dim AOLStatic6 As Long
Dim AOLIcon3 As Long
Dim AOLCheckbox As Long
Dim AOLGlyph As Long
Dim AOLStatic5 As Long
Dim AOLIcon2 As Long
Dim RICHCNTL As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim AOLStatic4 As Long
Dim AOLFontCombo As Long
Dim AOLEdit3 As Long
Dim AOLStatic3 As Long
Dim AOLEdit2 As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLEdit3& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
AOLFontCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_FontCombo", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 11&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLStatic7& = FindWindowEx(AOLChild&, AOLStatic6&, "_AOL_Static", vbNullString)
AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
AOLStatic8& = FindWindowEx(AOLChild&, AOLStatic7&, "_AOL_Static", vbNullString)
AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
AOLStatic9& = FindWindowEx(AOLChild&, AOLStatic8&, "_AOL_Static", vbNullString)
AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
AOLStatic10& = FindWindowEx(AOLChild&, AOLStatic9&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLEdit2& = 0& Or AOLStatic3& = 0& Or AOLEdit3& = 0& Or AOLFontCombo& = 0& Or AOLStatic4& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or RICHCNTL& = 0& Or AOLIcon2& = 0& Or AOLStatic5& = 0& Or AOLGlyph& = 0& Or AOLCheckbox& = 0& Or AOLIcon3& = 0& Or AOLStatic6& = 0& Or AOLIcon4& = 0& Or AOLStatic7& = 0& Or AOLIcon5& = 0& Or AOLStatic8& = 0& Or AOLIcon6& = 0& Or AOLStatic9& = 0& Or AOLIcon7& = 0& Or AOLStatic10& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLEdit3& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
    AOLFontCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_FontCombo", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 11&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLStatic7& = FindWindowEx(AOLChild&, AOLStatic6&, "_AOL_Static", vbNullString)
    AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
    AOLStatic8& = FindWindowEx(AOLChild&, AOLStatic7&, "_AOL_Static", vbNullString)
    AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
    AOLStatic9& = FindWindowEx(AOLChild&, AOLStatic8&, "_AOL_Static", vbNullString)
    AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
    AOLStatic10& = FindWindowEx(AOLChild&, AOLStatic9&, "_AOL_Static", vbNullString)
    If AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLStatic3& And AOLEdit3& And AOLFontCombo& And AOLStatic4& And AOLCombobox& And AOLIcon& And RICHCNTL& And AOLIcon2& And AOLStatic5& And AOLGlyph& And AOLCheckbox& And AOLIcon3& And AOLStatic6& And AOLIcon4& And AOLStatic7& And AOLIcon5& And AOLStatic8& And AOLIcon6& And AOLStatic9& And AOLIcon7& And AOLStatic10& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindMail& = AOLChild&
    Exit Function
End If
End Function
Public Function MailSend(Screennames As String, subject As String, message As String) As Long
Dim AOLIcon As Long
Dim AOLToolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
Dim RICHCNTL As Long
Dim i As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Do
TimeNow 0.4
Loop Until FindMail <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindMail))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Screennames)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindMail))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, subject)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindMail))
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, message)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindMail))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 15&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
DoEvents
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindMail = 0 Or FindNotFilled <> 0
If FindNotFilled <> 0 Then
Call CloseWindow(FindNotFilled)
MsgBox "Could not send due to in-complete ""mail"" form.", 64
Call CloseWindow(FindMail)
Do
DoEvents
Loop
End If
End Function
Public Function WWWPopUpKiller() As Long
' put this in a timer

If FindWWW > 1 Then
Do
Call SetFocusAPI(FindWWW)
Call SendMessage(FindWWW, WM_CLOSE, 0&, 0&)
Loop Until FindWWW = 1
End If
End Function
Public Function GetRidWAOL() As Long
Do
TimeNow 0.4
Loop Until FindWAOL <> 0
Call DisableWin(FindWAOL)
Call MinimizeWin(FindWAOL)
Call EnableWin(FindAOL)
Call SetFocusWin(FindAOL)
End Function
Public Function MakeDirectory(name As String)
Call MkDir(name)
End Function
Public Function DeleteDirectory(name As String)
Call RmDir(name)
End Function
Public Function DeleteFile(File As String)
On Error GoTo oops
Call Kill(File)
oops:
Exit Function
End Function
Public Function IMOff()
Call InstantMessage("$Im_Off", "dsk.bas")
End Function
Public Function IMOn()
Call InstantMessage("$Im_On", "dsk.bas")
End Function
Public Function FileExists(File As String) As Boolean
    If Len(File$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(File$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function
Public Function FindWAOL() As Long
Dim Counter As Long
Dim StaticW As Long
Dim Edit As Long
Dim i As Long
Dim Button As Long
Dim child As Long
child& = FindWindow("#32770", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
For i& = 1& To 2&
    Button& = FindWindowEx(child&, Button&, "Button", vbNullString)
Next i&
Edit& = FindWindowEx(child&, 0&, "Edit", vbNullString)
StaticW& = FindWindowEx(child&, 0&, "Static", vbNullString)
For i& = 1& To 2&
    StaticW& = FindWindowEx(child&, StaticW&, "Static", vbNullString)
Next i&
Do While (Counter& <> 100&) And (Button& = 0& Or Edit& = 0& Or StaticW& = 0&): DoEvents
    child& = FindWindowEx(child&, child&, "#32770", vbNullString)
    Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
    For i& = 1& To 2&
        Button& = FindWindowEx(child&, Button&, "Button", vbNullString)
    Next i&
    Edit& = FindWindowEx(child&, 0&, "Edit", vbNullString)
    StaticW& = FindWindowEx(child&, 0&, "Static", vbNullString)
    For i& = 1& To 2&
        StaticW& = FindWindowEx(child&, StaticW&, "Static", vbNullString)
    Next i&
    If Button& And Edit& And StaticW& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    Find_WAOL& = child&
    Exit Function
End If
End Function
Public Function AOLKeyword(Place As String) As Long
Dim AOL As Long
Dim tool As Long
Dim Toolbar As Long
Dim Combo As Long
Dim EditWin As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, Place)
Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Function
Public Function BuddyRoomInvite(Screennames As String, message As String, Room As String) As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim AOLCheckbox As Long
Dim AOLEdit As Long
If FindBuddyList = False Then
Call AOLKeyword("BuddyView")
Do
TimeNow 0.4
Loop Until FindBuddyList <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyList))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyInvite <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Screennames)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, message)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Room)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyInviteAccept <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End If
If FindBuddyList <> 0 Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyList))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyInvite <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Screennames)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, message)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Room)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindBuddyInvite))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyInviteAccept <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End If
End Function
Public Function FindBuddyInvite() As Long
Dim Counter As Long
Dim AOLCheckbox As Long
Dim AOLIcon As Long
Dim AOLEdit2 As Long
Dim i As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
For i& = 1& To 3&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
Next i&
AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLEdit2& = 0& Or AOLIcon& = 0& Or AOLCheckbox& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    For i& = 1& To 3&
        AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    Next i&
    AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    If AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLIcon& And AOLCheckbox& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindBuddyInvite& = AOLChild&
    Exit Function
End If
End Function
Public Function FindBuddyInviteAccept() As Long
Dim Counter As Long
Dim AOLView2 As Long
Dim AOLIcon As Long
Dim AOLStatic2 As Long
Dim AOLView As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLView2& = FindWindowEx(AOLChild&, AOLView&, "_AOL_View", vbNullString)
AOLView2& = FindWindowEx(AOLChild&, AOLView2&, "_AOL_View", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLView& = 0& Or AOLStatic2& = 0& Or AOLIcon& = 0& Or AOLView2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLView2& = FindWindowEx(AOLChild&, AOLView&, "_AOL_View", vbNullString)
    AOLView2& = FindWindowEx(AOLChild&, AOLView2&, "_AOL_View", vbNullString)
    If AOLStatic& And AOLView& And AOLStatic2& And AOLIcon& And AOLView2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindBuddyInviteGO& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAttatchWindow() As Long
Dim Counter As Long
Dim Child4 As Long
Dim ToolbarWindow As Long
Dim i As Long
Dim Button As Long
Dim ComboBox2 As Long
Dim Static4 As Long
Dim Edit As Long
Dim Static3 As Long
Dim SHELLDLLDefView As Long
Dim ListBox As Long
Dim Static2 As Long
Dim ComboBox As Long
Dim StaticA As Long
Dim Child3 As Long
Dim Child2 As Long
Dim child As Long
child& = FindWindow("#32770", vbNullString)
Child2& = FindWindowEx(child&, 0&, "#32770", vbNullString)
Child3& = FindWindowEx(Child2&, 0&, "#32770", vbNullString)
StaticA& = FindWindowEx(child&, 0&, "Static", vbNullString)
ComboBox& = FindWindowEx(child&, 0&, "ComboBox", vbNullString)
Static2& = FindWindowEx(child&, StaticA&, "Static", vbNullString)
ListBox& = FindWindowEx(child&, 0&, "ListBox", vbNullString)
SHELLDLLDefView& = FindWindowEx(child&, 0&, "SHELLDLL_DefView", vbNullString)
Static3& = FindWindowEx(child&, Static2&, "Static", vbNullString)
Edit& = FindWindowEx(child&, 0&, "Edit", vbNullString)
Static4& = FindWindowEx(child&, Static3&, "Static", vbNullString)
ComboBox2& = FindWindowEx(child&, ComboBox&, "ComboBox", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
For i& = 1& To 3&
    Button& = FindWindowEx(child&, Button&, "Button", vbNullString)
Next i&
ToolbarWindow& = FindWindowEx(child&, 0&, "ToolbarWindow32", vbNullString)
Child4& = FindWindowEx(child&, Child3&, "#32770", vbNullString)
Do While (Counter& <> 100&) And (StaticA& = 0& Or ComboBox& = 0& Or Static2& = 0& Or ListBox& = 0& Or SHELLDLLDefView& = 0& Or Static3& = 0& Or Edit& = 0& Or Static4& = 0& Or ComboBox2& = 0& Or Button& = 0& Or ToolbarWindow& = 0& Or Child4& = 0&): DoEvents
    child& = FindWindowEx(child&, child&, "#32770", vbNullString)
    StaticA& = FindWindowEx(child&, 0&, "Static", vbNullString)
    ComboBox& = FindWindowEx(child&, 0&, "ComboBox", vbNullString)
    Static2& = FindWindowEx(child&, StaticA&, "Static", vbNullString)
    ListBox& = FindWindowEx(child&, 0&, "ListBox", vbNullString)
    SHELLDLLDefView& = FindWindowEx(child&, 0&, "SHELLDLL_DefView", vbNullString)
    Static3& = FindWindowEx(child&, Static2&, "Static", vbNullString)
    Edit& = FindWindowEx(child&, 0&, "Edit", vbNullString)
    Static4& = FindWindowEx(child&, Static3&, "Static", vbNullString)
    ComboBox2& = FindWindowEx(child&, ComboBox&, "ComboBox", vbNullString)
    Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
    For i& = 1& To 3&
        Button& = FindWindowEx(child&, Button&, "Button", vbNullString)
    Next i&
    ToolbarWindow& = FindWindowEx(child&, 0&, "ToolbarWindow32", vbNullString)
    Child4& = FindWindowEx(child&, Child3&, "#32770", vbNullString)
    If StaticA& And ComboBox& And Static2& And ListBox& And SHELLDLLDefView& And Static3& And Edit& And Static4& And ComboBox2& And Button& And ToolbarWindow& And Child4& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAttatchWindow& = child&
    Exit Function
End If
End Function
Public Function FindProfileWarn() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLCheckbox As Long
Dim AOLStatic As Long
Dim AOLGlyph As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLGlyph& = 0& Or AOLStatic& = 0& Or AOLCheckbox& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    If AOLGlyph& And AOLStatic& And AOLCheckbox& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindProfileWarn& = AOLModal&
    Exit Function
End If
End Function
Public Function EditProfile(name As String, CityOrStateOrCountry As String, Birthday As String, Gender As String, MaritalStatus As String, Hobbies As String, ComputersUsed As String, Occupation As String, PersonalQuote As String, IncludeLink As Boolean) As Long

' this function will only work if you still have
' that annoying pop-up waring, and if you don't
' have it, then you will have to copy what u had
' put before, then delete it, or just delete it
' or you can use the "DeleteProfile" function in
' this bas. AND THIS FUNCTION ONLY WORKS IF
' GENDER IS A "MALE". i don't know why, but i'am
' trying to figure this out, so you can chose
' anyone, but until then......only male.

' Gender Options
' when choosing a gender, after the BirthDay
' Type M or m = Male
' Type F or f = Female ( working on.... )
' Type NR or nr = No Response ( working on.... )

Dim AOLIcon As Long
Dim AOLCheckbox As Long
Dim i As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim AOLModal As Long
Dim Button As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%ay" ' i know appactivate and sendkeys suck but, in some cases, they really save time.
Do
TimeNow 0.4
Loop Until FindProfileWarn <> 0
If FindProfileWarn = 0 Then
GoTo NoWarn
Else
If FindProfileWarn <> 0 Then
'AOLModal& = FindWindow("_AOL_Modal", vbNullString)
'AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
'Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
'Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindProfileWarn = 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, name)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, CityOrStateOrCountry)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Birthday)
If Gender = LCase("M") Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 3&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, MaritalStatus)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Hobbies)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 5&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ComputersUsed)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 6&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Occupation)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 7&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, PersonalQuote)
If IncludeLink = True Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf IncludeLink = False Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf Gender = LCase("F") Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 3&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, MaritalStatus)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Hobbies)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 5&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ComputersUsed)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 6&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Occupation)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 7&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, PersonalQuote)
If IncludeLink = True Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf IncludeLink = False Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf Gender = LCase("NR") Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 2&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 3&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, MaritalStatus)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Hobbies)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 5&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ComputersUsed)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 6&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Occupation)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 7&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, PersonalQuote)
If IncludeLink = True Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf IncludeLink = False Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
NoWarn:
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, name)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, CityOrStateOrCountry)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Birthday)
If Gender = LCase("M") Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 3&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, MaritalStatus)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Hobbies)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 5&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ComputersUsed)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 6&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Occupation)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 7&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, PersonalQuote)
If IncludeLink = True Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf IncludeLink = False Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf Gender = LCase("F") Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 3&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, MaritalStatus)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Hobbies)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 5&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ComputersUsed)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 6&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Occupation)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 7&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, PersonalQuote)
If IncludeLink = True Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf IncludeLink = False Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Call PostMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Button&, WM_LBUTTONUP, 0&, 0&)
ElseIf Gender = LCase("NR") Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 2&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 3&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, MaritalStatus)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Hobbies)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 5&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ComputersUsed)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 6&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Occupation)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 7&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, PersonalQuote)
If IncludeLink = "True" Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
ElseIf IncludeLink = False Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindEditProfile))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End Function
Public Function FindEditProfile() As Long
Dim Counter As Long
Dim AOLIcon2 As Long
Dim AOLStatic3 As Long
Dim AOLCheckbox2 As Long
Dim AOLEdit2 As Long
Dim AOLCheckbox As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLIcon As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 8&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 2&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
For i& = 1& To 4&
    AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
Next i&
AOLCheckbox2& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLCheckbox& = 0& Or AOLEdit2& = 0& Or AOLCheckbox2& = 0& Or AOLStatic3& = 0& Or AOLIcon2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 8&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    For i& = 1& To 2&
        AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
    For i& = 1& To 2&
        AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    Next i&
    AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
    For i& = 1& To 4&
        AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
    Next i&
    AOLCheckbox2& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 4&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And AOLIcon& And AOLEdit& And AOLStatic2& And AOLCheckbox& And AOLEdit2& And AOLCheckbox2& And AOLStatic3& And AOLIcon2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindEditProfile& = AOLChild&
    Exit Function
End If
End Function
Public Function FindProfileUpdate() As Long
Dim Counter As Long
Dim StaticP As Long
Dim Button As Long
Dim Child3 As Long
Dim Child2 As Long
Dim child As Long
child& = FindWindow("#32770", vbNullString)
Child2& = FindWindowEx(child&, 0&, "#32770", vbNullString)
Child3& = FindWindowEx(Child2&, 0&, "#32770", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
StaticP& = FindWindowEx(child&, 0&, "Static", vbNullString)
StaticP& = FindWindowEx(child&, StaticP&, "Static", vbNullString)
Do While (Counter& <> 100&) And (Button& = 0& Or StaticP& = 0&): DoEvents
    child& = FindWindowEx(child&, child&, "#32770", vbNullString)
    Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
    StaticP& = FindWindowEx(child&, 0&, "Static", vbNullString)
    StaticP& = FindWindowEx(child&, StaticP&, "Static", vbNullString)
    If Button& And StaticP& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindProfileUpdate& = child&
    Exit Function
End If
End Function
Public Function DeleteProfile() As Long

' DO NOT!!!!! use this if you have nothing in your
' profile. but if you do, then it is ok.

Dim AOLIcon As Long
Dim AOLCheckbox As Long
Dim AOLModal As Long
Dim i As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim child As Long
Dim Button As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%ay" ' i know appactivate and sendkeys suck but, in some cases, they really save time.
Do
TimeNow 0.4
Loop Until FindProfileWarn <> 0
If FindProfileWarn = 0 Then
GoTo NoWarn
Else
If FindProfileWarn <> 0 Then
'AOLModal& = FindWindow("_AOL_Modal", vbNullString)
'AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
'Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
'Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindProfileWarn = 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Edit Your Online Profile")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindDeleteProfile <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
NoWarn:
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Edit Your Online Profile")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindDeleteProfile <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End If
End If
End Function
Public Function FindDeleteProfile() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindDeleteProfile& = AOLModal&
    Exit Function
End If
End Function
Public Function FindWindows() As Long
Dim Counter As Long
Dim SysHeader As Long
Dim SysListView As Long
Dim SHELLDLLDefView As Long
Dim Progman As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
SysHeader& = FindWindowEx(SysListView&, 0&, "SysHeader32", vbNullString)
Do While (Counter& <> 100&) And (SysHeader& = 0&): DoEvents
    SysListView& = FindWindowEx(SHELLDLLDefView&, SysListView&, "SysListView32", vbNullString)
    SysHeader& = FindWindowEx(SysListView&, 0&, "SysHeader32", vbNullString)
    If SysHeader& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindWindows& = SysListView&
    Exit Function
End If
End Function
Public Function ShutDownWindows()
AppActivate GetWindowCaption(FindWindows)
SendKeys "%{f4}sy"
End Function
Public Function RestartWindows()
AppActivate GetWindowCaption(FindWindows)
SendKeys "%{f4}ry"
End Function
Public Function FindNotFilled() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindNotFilled& = AOLModal&
    Exit Function
End If
End Function
Public Function SetText(TextName As String, WhatToSet As String) As Long

'example
' if you have
'RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
'RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
' as text box, and you wanted to set some
' words in it. and instead of doing
'RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
'RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
'Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, WhatToSendToBox)
'all you have to do is
'RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
'RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
'Call Settext(RICHCNTL,Text1)
'something like that, it will save the time.

Call SendMessageByString(TextName, WM_SETTEXT, 0&, WhatToSet)
End Function
Public Function AOLUser() As String
Dim AOL As Long
Dim MDI As Long
Dim welcome As Long
Dim child As Long
Dim UserString As String
AOL& = FindWindow("AOL Frame25", vbNullString)
MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
UserString$ = GetWindowCaption(child&)
If InStr(UserString$, "Welcome, ") = 1 Then
UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
AOLUser$ = UserString$
Exit Function
Else
Do
child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
UserString$ = GetWindowCaption(child&)
If InStr(UserString$, "Welcome, ") = 1 Then
UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
AOLUser$ = UserString$
Exit Function
End If
Loop Until child& = 0&
End If
AOLUser$ = ""
End Function
Public Function FindBuddyEdit() As Long
Dim Counter As Long
Dim RICHCNTL As Long
Dim AOLIcon3 As Long
Dim AOLStatic3 As Long
Dim AOLIcon2 As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLListbox As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 5&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLListbox& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or AOLIcon2& = 0& Or AOLStatic3& = 0& Or AOLIcon3& = 0& Or RICHCNTL& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 5&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    If AOLStatic& And AOLListbox& And AOLIcon& And AOLStatic2& And AOLIcon2& And AOLStatic3& And AOLIcon3& And RICHCNTL& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindBuddyEdit& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAddBuddy() As Long
Dim Counter As Long
Dim i As Long
Dim AOLIcon3 As Long
Dim AOLGlyph2 As Long
Dim AOLIcon2 As Long
Dim AOLListbox As Long
Dim AOLStatic4 As Long
Dim AOLIcon As Long
Dim AOLEdit2 As Long
Dim AOLStatic3 As Long
Dim AOLGlyph As Long
Dim AOLEdit As Long
Dim AOLStatic2 As Long
Dim AOLCombobox As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLGlyph2& = FindWindowEx(AOLChild&, AOLGlyph&, "_AOL_Glyph", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLCombobox& = 0& Or AOLStatic2& = 0& Or AOLEdit& = 0& Or AOLGlyph& = 0& Or AOLStatic3& = 0& Or AOLEdit2& = 0& Or AOLIcon& = 0& Or AOLStatic4& = 0& Or AOLListbox& = 0& Or AOLIcon2& = 0& Or AOLGlyph2& = 0& Or AOLIcon3& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLGlyph2& = FindWindowEx(AOLChild&, AOLGlyph&, "_AOL_Glyph", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And AOLCombobox& And AOLStatic2& And AOLEdit& And AOLGlyph& And AOLStatic3& And AOLEdit2& And AOLIcon& And AOLStatic4& And AOLListbox& And AOLIcon2& And AOLGlyph2& And AOLIcon3& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAddBuddy& = AOLChild&
    Exit Function
End If
End Function
Public Function FindBuddyUpdate() As Long
Dim Counter As Long
Dim StaticU As Long
Dim Button As Long
Dim Child3 As Long
Dim Child2 As Long
Dim child As Long
child& = FindWindow("#32770", vbNullString)
Child2& = FindWindowEx(child&, 0&, "#32770", vbNullString)
Child3& = FindWindowEx(Child2&, 0&, "#32770", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
StaticU& = FindWindowEx(child&, 0&, "Static", vbNullString)
StaticU& = FindWindowEx(child&, StaticU&, "Static", vbNullString)
Do While (Counter& <> 100&) And (Button& = 0& Or StaticU& = 0&): DoEvents
    child& = FindWindowEx(child&, child&, "#32770", vbNullString)
    Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
    StaticU& = FindWindowEx(child&, 0&, "Static", vbNullString)
    StaticU& = FindWindowEx(child&, StaticU&, "Static", vbNullString)
    If Button& And StaticU& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindBuddyUpdate& = child&
    Exit Function
End If
End Function
Public Function Click(ButtonOrIcon As String)
Call PostMessage(ButtonOrIcon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ButtonOrIcon, WM_LBUTTONUP, 0&, 0&)
End Function
Public Function RoomCount() As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
RoomCount = GetText(AOLStatic&)
End Function
Public Function InstantMessage(ScreenName As String, message As String) As Long
Dim i As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Call AOLKeyword("aol://9293:" & ScreenName)
Do
TimeNow 0.4
Loop Until FindIMWindow <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindIMWindow))
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, message)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Send Instant Message")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 8&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
DoEvents
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindIMWindow = 0 Or FindNotFilled <> 0
If FindNotFilled <> 0 Then
Call CloseWindow(FindNotFilled)
MsgBox "Could not send due to in-complete ""instant message"" form.", 64
Call CloseWindow(FindIMWindow)
Do
DoEvents
Loop
End If
End Function
Public Function FindIMWindow() As Long
Dim Counter As Long
Dim AOLIcon2 As Long
Dim RICHCNTL As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLIcon& = 0& Or RICHCNTL& = 0& Or AOLIcon2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLEdit& And AOLIcon& And RICHCNTL& And AOLIcon2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindIMWindow& = AOLChild&
    Exit Function
End If
End Function
Public Function FindIMError() As Long
Dim Counter As Long
Dim StaticO As Long
Dim Button As Long
Dim Child2 As Long
Dim child As Long
child& = FindWindow("#32770", vbNullString)
Child2& = FindWindowEx(child&, 0&, "#32770", vbNullString)
Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
StaticO& = FindWindowEx(child&, 0&, "Static", vbNullString)
StaticO& = FindWindowEx(child&, StaticO&, "Static", vbNullString)
Do While (Counter& <> 100&) And (Button& = 0& Or StaticO& = 0&): DoEvents
    child& = FindWindowEx(child&, child&, "#32770", vbNullString)
    Button& = FindWindowEx(child&, 0&, "Button", vbNullString)
    StaticO& = FindWindowEx(child&, 0&, "Static", vbNullString)
    StaticO& = FindWindowEx(child&, StaticO&, "Static", vbNullString)
    If Button& And StaticO& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindIMError& = child&
    Exit Function
End If
End Function
Public Function CloseAOL()
Call CloseWindow(FindAOL)
End Function
Public Function FindUpchat() As Long
Dim Counter As Long
Dim AOLButton As Long
Dim AOLCheckbox As Long
Dim AOLStatic2 As Long
Dim AOLGauge As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLGauge& = FindWindowEx(AOLModal&, 0&, "_AOL_Gauge", vbNullString)
AOLGauge& = FindWindowEx(AOLModal&, AOLGauge&, "_AOL_Gauge", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
AOLButton& = FindWindowEx(AOLModal&, 0&, "_AOL_Button", vbNullString)
AOLButton& = FindWindowEx(AOLModal&, AOLButton&, "_AOL_Button", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLGauge& = 0& Or AOLStatic2& = 0& Or AOLCheckbox& = 0& Or AOLButton& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLGauge& = FindWindowEx(AOLModal&, 0&, "_AOL_Gauge", vbNullString)
    AOLGauge& = FindWindowEx(AOLModal&, AOLGauge&, "_AOL_Gauge", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
    AOLButton& = FindWindowEx(AOLModal&, 0&, "_AOL_Button", vbNullString)
    AOLButton& = FindWindowEx(AOLModal&, AOLButton&, "_AOL_Button", vbNullString)
    If AOLStatic& And AOLGauge& And AOLStatic2& And AOLCheckbox& And AOLButton& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindUpchat& = AOLModal&
    Exit Function
End If
End Function
Public Function UpChat()
Do
TimeNow 0.4
Loop Until FindUpchat <> 0
Call DisableWin(FindUpchat)
Call MinimizeWin(FindUpchat)
Call EnableWin(FindAOL)
Call SetFocusWin(FindAOL)
End Function
Public Function UnUpChat()
Do
TimeNow 0.4
Loop Until FindUpchat <> 0
Call DisableWin(FindAOL)
Call NormalShowSizeWin(FindUpchat)
Call EnableWin(FindUpchat)
Call SetFocusWin(FindUpchat)
End Function
Public Function EnableWin(Window As Long)
Call EnableWindow(Window, True)
End Function
Public Function DisableWin(Window As Long)
Call EnableWindow(Window, False)
End Function
Public Function LoadText(txt2Load As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    Open Path$ For Input As #1
    TextString$ = Input(LOF(1), #1)
    Close #1
    txt2Load.Text = TextString$
End Function
Public Function LoadLabel(Lbl2Load As Label, Path As String)
    Dim LabelString As String
    On Error Resume Next
    Open Path$ For Input As #1
    LabelString$ = Input(LOF(1), #1)
    Close #1
    Lbl2Load.Caption = LabelString$
End Function
Public Function LoadComboBox(Cmb2Load As ComboBox, ByVal Path As String)
Dim MyString As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        Cmb2Load.AddItem MyString$
    Wend
    Close #1
    End Function
    Public Function LoadList(Lst2Load As ListBox, ByVal Path As String)
Dim MyString As String
    On Error Resume Next
    Open Path$ For Input As #1
    While Not EOF(1)
        Input #1, MyString$
        DoEvents
        Lst2Load.AddItem MyString$
    Wend
    Close #1
    End Function
Public Function SaveText(txt2Save As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txt2Save.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Function
Public Function SaveLabel(Lbl2Save As Label, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = Lbl2Save.Caption
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Function
Public Function SaveComboBox(Cmb2Save As ComboBox, ByVal Path As String)
    Dim SaveCombo As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveCombo& = 0 To Cmb2Save.ListCount - 1
        Print #1, Cmb2Save.List(SaveCombo&)
    Next SaveCombo&
    Close #1
End Function
Public Function SaveList(Lst2Save As ListBox, ByVal Path As String)
       Dim SaveLis As Long
    On Error Resume Next
    Open Path$ For Output As #1
    For SaveLis& = 0 To Lst2Save.ListCount - 1
        Print #1, Lst2Save.List(SaveLis&)
    Next SaveLis&
    Close #1
End Function
Public Function SignOnGuest(ScreenName As String, Password As String) As Long
' not done yet

' will be in the next version

Dim AOLEdit As Long
Dim AOLModal As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
If FindAOL <> 0 Then
AppActivate GetWindowCaption(FindAOL)
SendKeys "%ss"
Do
TimeNow 0.4
Loop Until FindSignOff <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Goodbye from America Online!")
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
Call PostMessage(AOLCombobox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCombobox&, WM_LBUTTONUP, 0&, 0&)
SendKeys "{PgDn}"
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindSignOff))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Do
TimeNow 0.4
Loop Until FindAOLGuest <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ScreenName)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Password)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Else
MsgBox "America Online not loaded"
End If
End Function
Public Function FindAOLGuest() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLStatic3 As Long
Dim AOLEdit2 As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLStatic3& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLEdit2& = 0& Or AOLStatic3& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLStatic3& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLStatic3& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLGuest& = AOLModal&
    Exit Function
End If
End Function
Public Function Ghost() As Long
Dim AOLCheckbox As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Window")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyEdit <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Da Real dsk's Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyGhost <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 4&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 6&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyUpdate <> 0
Do
Call CloseWindow(FindBuddyUpdate)
TimeNow 0.4
Loop Until FindBuddyEdit <> 0
Call CloseWindow(FindBuddyEdit)
End Function
Public Function UnGhost() As Long
Dim AOLCheckbox As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Window")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyEdit <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Da Real dsk's Buddy List")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyGhost <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 5&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Privacy Preferences")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindBuddyUpdate <> 0
Do
Call CloseWindow(FindBuddyUpdate)
TimeNow 0.4
Loop Until FindBuddyEdit <> 0
Call CloseWindow(FindBuddyEdit)
End Function
Public Function FindBuddyGhost() As Long
Dim Counter As Long
Dim AOLIcon4 As Long
Dim AOLStatic6 As Long
Dim AOLIcon3 As Long
Dim AOLListbox As Long
Dim AOLIcon2 As Long
Dim AOLStatic5 As Long
Dim AOLEdit As Long
Dim AOLStatic4 As Long
Dim AOLCheckbox2 As Long
Dim AOLStatic3 As Long
Dim AOLCheckbox As Long
Dim AOLStatic2 As Long
Dim AOLIcon As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 4&
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLCheckbox2& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
AOLCheckbox2& = FindWindowEx(AOLChild&, AOLCheckbox2&, "_AOL_Checkbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or AOLCheckbox& = 0& Or AOLStatic3& = 0& Or AOLCheckbox2& = 0& Or AOLStatic4& = 0& Or AOLEdit& = 0& Or AOLStatic5& = 0& Or AOLIcon2& = 0& Or AOLListbox& = 0& Or AOLIcon3& = 0& Or AOLStatic6& = 0& Or AOLIcon4& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
    For i& = 1& To 4&
        AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    Next i&
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLCheckbox2& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    AOLCheckbox2& = FindWindowEx(AOLChild&, AOLCheckbox2&, "_AOL_Checkbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLIcon& And AOLStatic2& And AOLCheckbox& And AOLStatic3& And AOLCheckbox2& And AOLStatic4& And AOLEdit& And AOLStatic5& And AOLIcon2& And AOLListbox& And AOLIcon3& And AOLStatic6& And AOLIcon4& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindBuddyGhost& = AOLChild&
    Exit Function
End If
End Function
Public Function FindKeyword() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And AOLEdit& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindKeyword& = AOLChild&
    Exit Function
End If
End Function
Public Function AOLKeyword2(Place As String) As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLToolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 21&
    AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindKeyword <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindKeyword))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Place)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindKeyword))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindKeyword = 0 Or FindKeyword <> 0
If FindKeyword <> 0 Then
TimeNow 0.6
Do
Call CloseWindow(FindKeyword)
Loop Until FindKeyword = 0
End If
End Function
Public Function FindAOLNumbers() As Long
Dim Counter As Long
Dim RICHCNTL As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0& Or RICHCNTL& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    If AOLStatic& And AOLIcon& And RICHCNTL& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLNumbers& = AOLChild&
    Exit Function
End If
End Function
Public Function CloseAOLNewANumbers() As Long
Call CloseWindow(FindAOLNumbers)
End Function
Public Function FindFavorites() As Long
Dim Counter As Long
Dim AOLStatic5 As Long
Dim AOLIcon5 As Long
Dim AOLStatic4 As Long
Dim AOLIcon4 As Long
Dim AOLStatic3 As Long
Dim AOLIcon3 As Long
Dim AOLStatic2 As Long
Dim AOLIcon2 As Long
Dim AOLStatic As Long
Dim AOLIcon As Long
Dim AOLTree As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLTree& = 0& Or AOLIcon& = 0& Or AOLStatic& = 0& Or AOLIcon2& = 0& Or AOLStatic2& = 0& Or AOLIcon3& = 0& Or AOLStatic3& = 0& Or AOLIcon4& = 0& Or AOLStatic4& = 0& Or AOLIcon5& = 0& Or AOLStatic5& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    If AOLTree& And AOLIcon& And AOLStatic& And AOLIcon2& And AOLStatic2& And AOLIcon3& And AOLStatic3& And AOLIcon4& And AOLStatic4& And AOLIcon5& And AOLStatic5& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindFavorites& = AOLChild&
    Exit Function
End If
End Function
Public Function FindFavoritesAdd() As Long
Dim Counter As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLEdit3 As Long
Dim AOLStatic3 As Long
Dim AOLEdit2 As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLCheckbox As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLEdit3& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLCheckbox& = 0& Or AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLEdit2& = 0& Or AOLStatic3& = 0& Or AOLEdit3& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit2& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLEdit3& = FindWindowEx(AOLChild&, AOLEdit2&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLCheckbox& And AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLStatic3& And AOLEdit3& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindFavoritesAdd& = AOLChild&
    Exit Function
End If
End Function
Public Function AddToFavorites(Place As String, LinkOrUrl As String) As Long
Dim i As Long
Dim AOLEdit As Long
Dim AOLCheckbox As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%vf"
Do
TimeNow 0.4
Loop Until FindFavorites <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindFavorites))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindFavoritesAdd <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindFavoritesAdd))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindFavoritesAdd))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Place)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindFavoritesAdd))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLChild&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, LinkOrUrl)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindFavoritesAdd))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindFavoritesAdd = 0
If FindFavorites <> 0 Then
Do
Call CloseWindow(FindFavorites)
Loop Until FindFavorites = 0
Else
Exit Function
End If
End Function
Public Function FindAOLPrefs() As Long
Dim Counter As Long
Dim AOLIcon18 As Long
Dim AOLStatic18 As Long
Dim AOLIcon17 As Long
Dim AOLStatic17 As Long
Dim AOLIcon16 As Long
Dim AOLStatic16 As Long
Dim AOLIcon15 As Long
Dim AOLStatic15 As Long
Dim AOLIcon14 As Long
Dim AOLStatic14 As Long
Dim AOLIcon13 As Long
Dim AOLStatic13 As Long
Dim AOLIcon12 As Long
Dim AOLStatic12 As Long
Dim AOLIcon11 As Long
Dim AOLStatic11 As Long
Dim AOLIcon10 As Long
Dim AOLStatic10 As Long
Dim AOLIcon9 As Long
Dim AOLStatic9 As Long
Dim AOLIcon8 As Long
Dim AOLStatic8 As Long
Dim AOLIcon7 As Long
Dim AOLStatic7 As Long
Dim AOLIcon6 As Long
Dim AOLStatic6 As Long
Dim AOLIcon5 As Long
Dim AOLStatic5 As Long
Dim AOLIcon4 As Long
Dim AOLStatic4 As Long
Dim AOLIcon3 As Long
Dim AOLStatic3 As Long
Dim AOLIcon2 As Long
Dim AOLStatic2 As Long
Dim AOLIcon As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
AOLStatic7& = FindWindowEx(AOLChild&, AOLStatic6&, "_AOL_Static", vbNullString)
AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
AOLStatic8& = FindWindowEx(AOLChild&, AOLStatic7&, "_AOL_Static", vbNullString)
AOLIcon8& = FindWindowEx(AOLChild&, AOLIcon7&, "_AOL_Icon", vbNullString)
AOLStatic9& = FindWindowEx(AOLChild&, AOLStatic8&, "_AOL_Static", vbNullString)
AOLIcon9& = FindWindowEx(AOLChild&, AOLIcon8&, "_AOL_Icon", vbNullString)
AOLStatic10& = FindWindowEx(AOLChild&, AOLStatic9&, "_AOL_Static", vbNullString)
AOLIcon10& = FindWindowEx(AOLChild&, AOLIcon9&, "_AOL_Icon", vbNullString)
AOLStatic11& = FindWindowEx(AOLChild&, AOLStatic10&, "_AOL_Static", vbNullString)
AOLIcon11& = FindWindowEx(AOLChild&, AOLIcon10&, "_AOL_Icon", vbNullString)
AOLStatic12& = FindWindowEx(AOLChild&, AOLStatic11&, "_AOL_Static", vbNullString)
AOLIcon12& = FindWindowEx(AOLChild&, AOLIcon11&, "_AOL_Icon", vbNullString)
AOLStatic13& = FindWindowEx(AOLChild&, AOLStatic12&, "_AOL_Static", vbNullString)
AOLIcon13& = FindWindowEx(AOLChild&, AOLIcon12&, "_AOL_Icon", vbNullString)
AOLStatic14& = FindWindowEx(AOLChild&, AOLStatic13&, "_AOL_Static", vbNullString)
AOLIcon14& = FindWindowEx(AOLChild&, AOLIcon13&, "_AOL_Icon", vbNullString)
AOLStatic15& = FindWindowEx(AOLChild&, AOLStatic14&, "_AOL_Static", vbNullString)
AOLIcon15& = FindWindowEx(AOLChild&, AOLIcon14&, "_AOL_Icon", vbNullString)
AOLStatic16& = FindWindowEx(AOLChild&, AOLStatic15&, "_AOL_Static", vbNullString)
AOLIcon16& = FindWindowEx(AOLChild&, AOLIcon15&, "_AOL_Icon", vbNullString)
AOLStatic17& = FindWindowEx(AOLChild&, AOLStatic16&, "_AOL_Static", vbNullString)
AOLIcon17& = FindWindowEx(AOLChild&, AOLIcon16&, "_AOL_Icon", vbNullString)
AOLStatic18& = FindWindowEx(AOLChild&, AOLStatic17&, "_AOL_Static", vbNullString)
AOLIcon18& = FindWindowEx(AOLChild&, AOLIcon17&, "_AOL_Icon", vbNullString)
AOLIcon18& = FindWindowEx(AOLChild&, AOLIcon18&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or AOLIcon2& = 0& Or AOLStatic3& = 0& Or AOLIcon3& = 0& Or AOLStatic4& = 0& Or AOLIcon4& = 0& Or AOLStatic5& = 0& Or AOLIcon5& = 0& Or AOLStatic6& = 0& Or AOLIcon6& = 0& Or AOLStatic7& = 0& Or AOLIcon7& = 0& Or AOLStatic8& = 0& Or AOLIcon8& = 0& Or AOLStatic9& = 0& Or AOLIcon9& = 0& Or AOLStatic10& = 0& Or AOLIcon10& = 0& Or AOLStatic11& = 0& Or AOLIcon11& = 0& Or AOLStatic12& = 0& Or AOLIcon12& = 0& Or AOLStatic13& = 0& Or AOLIcon13& = 0& Or AOLStatic14& = 0& Or AOLIcon14& = 0& Or AOLStatic15& = 0& Or AOLIcon15& = 0& Or AOLStatic16& = 0& Or AOLIcon16& = 0& Or AOLStatic17& = 0& Or AOLIcon17& = 0& Or AOLStatic18& = 0& Or AOLIcon18& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
    AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
    AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
    AOLStatic7& = FindWindowEx(AOLChild&, AOLStatic6&, "_AOL_Static", vbNullString)
    AOLIcon7& = FindWindowEx(AOLChild&, AOLIcon6&, "_AOL_Icon", vbNullString)
    AOLStatic8& = FindWindowEx(AOLChild&, AOLStatic7&, "_AOL_Static", vbNullString)
    AOLIcon8& = FindWindowEx(AOLChild&, AOLIcon7&, "_AOL_Icon", vbNullString)
    AOLStatic9& = FindWindowEx(AOLChild&, AOLStatic8&, "_AOL_Static", vbNullString)
    AOLIcon9& = FindWindowEx(AOLChild&, AOLIcon8&, "_AOL_Icon", vbNullString)
    AOLStatic10& = FindWindowEx(AOLChild&, AOLStatic9&, "_AOL_Static", vbNullString)
    AOLIcon10& = FindWindowEx(AOLChild&, AOLIcon9&, "_AOL_Icon", vbNullString)
    AOLStatic11& = FindWindowEx(AOLChild&, AOLStatic10&, "_AOL_Static", vbNullString)
    AOLIcon11& = FindWindowEx(AOLChild&, AOLIcon10&, "_AOL_Icon", vbNullString)
    AOLStatic12& = FindWindowEx(AOLChild&, AOLStatic11&, "_AOL_Static", vbNullString)
    AOLIcon12& = FindWindowEx(AOLChild&, AOLIcon11&, "_AOL_Icon", vbNullString)
    AOLStatic13& = FindWindowEx(AOLChild&, AOLStatic12&, "_AOL_Static", vbNullString)
    AOLIcon13& = FindWindowEx(AOLChild&, AOLIcon12&, "_AOL_Icon", vbNullString)
    AOLStatic14& = FindWindowEx(AOLChild&, AOLStatic13&, "_AOL_Static", vbNullString)
    AOLIcon14& = FindWindowEx(AOLChild&, AOLIcon13&, "_AOL_Icon", vbNullString)
    AOLStatic15& = FindWindowEx(AOLChild&, AOLStatic14&, "_AOL_Static", vbNullString)
    AOLIcon15& = FindWindowEx(AOLChild&, AOLIcon14&, "_AOL_Icon", vbNullString)
    AOLStatic16& = FindWindowEx(AOLChild&, AOLStatic15&, "_AOL_Static", vbNullString)
    AOLIcon16& = FindWindowEx(AOLChild&, AOLIcon15&, "_AOL_Icon", vbNullString)
    AOLStatic17& = FindWindowEx(AOLChild&, AOLStatic16&, "_AOL_Static", vbNullString)
    AOLIcon17& = FindWindowEx(AOLChild&, AOLIcon16&, "_AOL_Icon", vbNullString)
    AOLStatic18& = FindWindowEx(AOLChild&, AOLStatic17&, "_AOL_Static", vbNullString)
    AOLIcon18& = FindWindowEx(AOLChild&, AOLIcon17&, "_AOL_Icon", vbNullString)
    AOLIcon18& = FindWindowEx(AOLChild&, AOLIcon18&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLIcon& And AOLStatic2& And AOLIcon2& And AOLStatic3& And AOLIcon3& And AOLStatic4& And AOLIcon4& And AOLStatic5& And AOLIcon5& And AOLStatic6& And AOLIcon6& And AOLStatic7& And AOLIcon7& And AOLStatic8& And AOLIcon8& And AOLStatic9& And AOLIcon9& And AOLStatic10& And AOLIcon10& And AOLStatic11& And AOLIcon11& And AOLStatic12& And AOLIcon12& And AOLStatic13& And AOLIcon13& And AOLStatic14& And AOLIcon14& And AOLStatic15& And AOLIcon15& And AOLStatic16& And AOLIcon16& And AOLStatic17& And AOLIcon17& And AOLStatic18& And AOLIcon18& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLPrefs& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAOLToolBarPrefs() As Long
Dim Counter As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLStatic5 As Long
Dim AOLCheckbox4 As Long
Dim AOLStatic4 As Long
Dim AOLCheckbox3 As Long
Dim AOLStatic3 As Long
Dim AOLCheckbox2 As Long
Dim AOLStatic2 As Long
Dim AOLCheckbox As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox2& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
AOLCheckbox2& = FindWindowEx(AOLModal&, AOLCheckbox2&, "_AOL_Checkbox", vbNullString)
AOLStatic3& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLCheckbox3& = FindWindowEx(AOLModal&, AOLCheckbox2&, "_AOL_Checkbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLModal&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLStatic4& = FindWindowEx(AOLModal&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLCheckbox4& = FindWindowEx(AOLModal&, AOLCheckbox3&, "_AOL_Checkbox", vbNullString)
AOLStatic5& = FindWindowEx(AOLModal&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLCheckbox& = 0& Or AOLStatic2& = 0& Or AOLCheckbox2& = 0& Or AOLStatic3& = 0& Or AOLCheckbox3& = 0& Or AOLStatic4& = 0& Or AOLCheckbox4& = 0& Or AOLStatic5& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCheckbox2& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    AOLCheckbox2& = FindWindowEx(AOLModal&, AOLCheckbox2&, "_AOL_Checkbox", vbNullString)
    AOLStatic3& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLCheckbox3& = FindWindowEx(AOLModal&, AOLCheckbox2&, "_AOL_Checkbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLModal&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLStatic4& = FindWindowEx(AOLModal&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLCheckbox4& = FindWindowEx(AOLModal&, AOLCheckbox3&, "_AOL_Checkbox", vbNullString)
    AOLStatic5& = FindWindowEx(AOLModal&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And AOLCheckbox& And AOLStatic2& And AOLCheckbox2& And AOLStatic3& And AOLCheckbox3& And AOLStatic4& And AOLCheckbox4& And AOLStatic5& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLToolBarPrefs& = AOLModal&
    Exit Function
End If
End Function
Public Function FindAOLToolBarClearQ() As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
End Function
Public Function ClearHistory() As Long
' this don't work yet.... but the sendkeys version does

Dim AOLModal As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%ap"
Do
TimeNow 0.4
Loop Until FindAOLPrefs <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindAOLPrefs))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLToolBarPrefs <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Do
TimeNow 0.4
Loop Until FindAOLToolBarClearQ <> 0
Do
AOLIcon& = FindWindowEx(FindAOLToolBarClearQ, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLToolBarClearQ = 0
Call CloseWindow(FindAOLToolBarClearQ)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLToolBarPrefs = 0
If FindAOLPrefs <> 0 Then
Do
Call CloseWindow(FindAOLPrefs)
Loop Until FindAOLPrefs = 0
Else
Exit Function
End If
End Function
Public Function SendKeysClearHistory()
' to me this is quicker then the api version of clearhistory
' and again, i think sometimes sendkeys is alot quicker
' and saves time instead of writing 10 lines of code
' for clicking a button twice.

AppActivate GetWindowCaption(FindAOL)
SendKeys "%ap{tab}{ }{tab}{tab}{tab}{tab}{ }{ }{tab}{ }^{f4}"
End Function
Public Function FindHotMail() As Long
Dim Counter As Long
Dim AOLGlyph As Long
Dim AOLGauge As Long
Dim AOLStatic As Long
Dim AOLWWWView As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLWWWView& = FindWindowEx(AOLChild&, 0&, "_AOL_WWWView", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLGauge& = FindWindowEx(AOLChild&, 0&, "_AOL_Gauge", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
Do While (Counter& <> 100&) And (AOLWWWView& = 0& Or AOLStatic& = 0& Or AOLGauge& = 0& Or AOLGlyph& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLWWWView& = FindWindowEx(AOLChild&, 0&, "_AOL_WWWView", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLGauge& = FindWindowEx(AOLChild&, 0&, "_AOL_Gauge", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    If AOLWWWView& And AOLStatic& And AOLGauge& And AOLGlyph& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindHotMail& = AOLChild&
    Exit Function
End If
End Function
Public Function HotMailLogin(User As String, Password As String)
' this works but, i will fix it better later

Call AOLKeyword("http://www.hotmail.com")
Do
TimeNow 0.4
Loop Until GetWindowCaption(FindHotMail) = "Hotmail - The World's FREE Web-based E-mail"
SendKeys User + "{tab}{tab}" + Password + "{tab}{ }"
End Function
Public Function CopySetText(What As String)
Clipboard.Clear
Clipboard.SetText What
End Function
Public Function CopyGetText(What As String)
What = Clipboard.GetText
End Function
Public Function FindIMNotify() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLEdit& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindIMNotify& = AOLModal&
    Exit Function
End If
End Function
Public Function FindIM() As Long
Dim Counter As Long
Dim AOLStatic2 As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLView As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
AOLView& = FindWindowEx(AOLChild&, AOLView&, "_AOL_View", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 5&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLView& = 0& Or AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLIcon& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or AOLStatic2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
    AOLView& = FindWindowEx(AOLChild&, AOLView&, "_AOL_View", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 5&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    If AOLView& And AOLStatic& And RICHCNTL& And AOLIcon& And RICHCNTL2& And AOLIcon2& And AOLStatic2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindIM& = AOLChild&
    Exit Function
End If
End Function
Public Function IMNotify(message As String) As Long
Dim AOLFrame As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLModal As Long
If FindIM <> 0 Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 13&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindIMNotify <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, message)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindIMNotify = 0
Else
MsgBox "There is no im window open"
Exit Function
End If
End Function
Public Function CenterForm(Form As Form)
Form.Top = (Screen.Height * 0.85) / 2 - Form.Height / 2
Form.Left = Screen.Width / 2 - Form.Width / 2
End Function
Public Function DragForm(Form As Form)
    Call ReleaseCapture
    Call SendMessage(Form.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Function
Public Function OnTopForm(Form As Form)
    Call SetWindowPos(Form.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Function
Public Function NotOnTopForm(Form As Form)
    Call SetWindowPos(Form.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Function
Public Function GetLineCount(Text)
Dim TheChar
Dim findchar
For findchar = 1 To Len(Text)
    TheChar = Mid(Text, findchar, 1)
    If TheChar = Chr(13) Then linenum = linenum + 1
Next findchar
If Mid(Text, Len(Text), 1) = Chr(13) Then
    GetLineCount = linenum
Else
    GetLineCount = linenum + 1
End If
End Function
Public Function LineFromText(Text, theline)
Dim thetext
Dim TempNum
Dim TheChars
Dim TheChar
Dim findchar
For findchar = 1 To Len(Text)
    TheChar = Mid(Text, findchar, 1)
    TheChars = TheChars & TheChar
    If TheChar = Chr(13) Then
        TempNum = TempNum + 1
        thetext = Mid(TheChars, 1, Len(TheChars) - 1)
        If theline = TempNum Then GoTo SkipIt
        TheChars = ""
    End If
Next findchar
Exit Function
SkipIt:
thetext = ReplaceText(thetext, Chr(13), "")
LineFromText = thetext
End Function
Public Function ReplaceText(Text, Find, Replace)
' ex. text1 = replacetext(text1, "hate","luv")
' then that will replace all the words with hate to luv

a = InStr(Text, Find)
If a = 0 Then
    ReplaceText = Text
    Exit Function
End If
Do: DoEvents
    c = Left(Text, a - 1)
    d = Mid(Text, a + Len(Find))
    e = c & Replace & d
    Text = e
    a = InStr(Text, Find)
Loop Until a = 0
ReplaceText = Text
End Function
Public Function ReverseText(Text As String) As String
Dim TempString As String
Dim StringLength As Long
Dim Count As Long
Dim NextChr As String
Dim NewString As String
TempString$ = Text$
StringLength& = Len(TempString$)
Do While Count& <= StringLength&
Count& = Count& + 1
NextChr$ = Mid$(TempString$, Count&, 1)
NewString$ = NextChr$ & NewString$
Loop
ReverseText$ = NewString$
End Function
Public Function ChatClear()
Call SendMessageByString(FindWindowEx(FindChat(), 0&, "RICHCNTL", vbNullString), WM_SETTEXT, 0&, vbNullChar)
End Function
Public Function CompactPFC() As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim AOLModal As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%yp"
Do
TimeNow 0.4
Loop Until FindPFC <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Da Real dsk's Filing Cabinet")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 6&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindPFCConfirm <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Function FindPFC() As Long
Dim Counter As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLTree As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLTree& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLTree& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindPFC& = AOLChild&
    Exit Function
End If
End Function
Public Function FindPFCConfirm() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindPFCConfirm& = AOLModal&
    Exit Function
End If
End Function
Public Function NetZeroBanner() As Long
Dim AwtWindow As Long
Dim AwtFrame As Long
AwtFrame& = FindWindow("AwtFrame", vbNullString)
AwtWindow& = FindWindowEx(AwtFrame&, 0&, "AwtWindow", vbNullString)
AwtWindow& = FindWindowEx(AwtFrame&, AwtWindow&, "AwtWindow", vbNullString)
End Function
Public Function FindSetUpSignature() As Long
Dim Counter As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLTree As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 4&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLTree& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLTree& = FindWindowEx(AOLChild&, 0&, "_AOL_Tree", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 4&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLStatic& And AOLTree& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindSetUpSignature& = AOLChild&
    Exit Function
End If
End Function
Public Function FindCreateSignature() As Long
Dim Counter As Long
Dim AOLIcon2 As Long
Dim RICHCNTL As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim AOLStatic3 As Long
Dim AOLFontCombo As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLFontCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_FontCombo", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 9&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLFontCombo& = 0& Or AOLStatic3& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or RICHCNTL& = 0& Or AOLIcon2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLFontCombo& = FindWindowEx(AOLChild&, 0&, "_AOL_FontCombo", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 9&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLEdit& And AOLStatic2& And AOLFontCombo& And AOLStatic3& And AOLCombobox& And AOLIcon& And RICHCNTL& And AOLIcon2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindCreateSignature& = AOLChild&
    Exit Function
End If
End Function
Public Function CreateSignature(name As String, Signature As String) As Long
Dim i As Long
Dim RICHCNTL As Long
Dim AOLEdit As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%mg"
Do
TimeNow 0.4
Loop Until FindSetUpSignature <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindSetUpSignature))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindCreateSignature <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindCreateSignature))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, name)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindCreateSignature))
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, "")
Call SendMessageByString(RICHCNTL&, WM_SETTEXT, 0&, Signature)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindCreateSignature))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 10&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindCreateSignature = 0
If FindSetUpSignature <> 0 Then
Do
Call CloseWindow(FindSetUpSignature)
Loop Until FindSetUpSignature = 0
End If
End Function
Public Function FindFTP() As Long
Dim Counter As Long
Dim AOLImage As Long
Dim AOLIcon2 As Long
Dim AOLStatic3 As Long
Dim AOLListbox As Long
Dim AOLIcon As Long
Dim AOLStatic2 As Long
Dim RICHCNTL As Long
Dim AOLStatic As Long
Dim AOLGlyph As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
Do While (Counter& <> 100&) And (AOLGlyph& = 0& Or AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLStatic2& = 0& Or AOLIcon& = 0& Or AOLListbox& = 0& Or AOLStatic3& = 0& Or AOLIcon2& = 0& Or AOLImage& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    If AOLGlyph& And AOLStatic& And RICHCNTL& And AOLStatic2& And AOLIcon& And AOLListbox& And AOLStatic3& And AOLIcon2& And AOLImage& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindFTP& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAFTP() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLListbox As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLGlyph As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 4&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do While (Counter& <> 100&) And (AOLGlyph& = 0& Or AOLStatic& = 0& Or AOLListbox& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 4&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    If AOLGlyph& And AOLStatic& And AOLListbox& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAFTP& = AOLChild&
    Exit Function
End If
End Function
Public Function findoftp() As Long
Dim Counter As Long
Dim AOLStatic2 As Long
Dim AOLImage As Long
Dim AOLListbox As Long
Dim AOLIcon As Long
Dim AOLCheckbox As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 4&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLCheckbox& = 0& Or AOLIcon& = 0& Or AOLListbox& = 0& Or AOLImage& = 0& Or AOLStatic2& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 4&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    If AOLStatic& And AOLEdit& And AOLCheckbox& And AOLIcon& And AOLListbox& And AOLImage& And AOLStatic2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    findoftp& = AOLChild&
    Exit Function
End If
End Function
Public Function FindFTPLogin() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLEdit2 As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim AOLStatic As Long
Dim AOLGlyph As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLGlyph& = 0& Or AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLEdit2& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLGlyph& And AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindFTPLogin& = AOLModal&
    Exit Function
End If
End Function
Public Function FTPLogin(FTPStieURL As String, UserName As String, Password As String) As Long
' this doesn't work yet. next bas, v³ it will

Dim AOLModal As Long
Dim AOLCheckbox As Long
Dim AOLEdit As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%if"
Do
TimeNow 0.4
Loop Until FindFTP <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAFTP <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindAFTP))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAFTP <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(findoftp))
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, FTPSiteURL)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(findoftp))
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(findoftp))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindFTPLogin <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, UserName)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Password)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindFTPLogin = 0
Do
Call CloseWindow(FindFTP)
Loop Until FindFTP = 0
Do
Call CloseWindow(FindAFTP)
Loop Until FindAFTP = 0
Do
Call CloseWindow(findoftp)
Loop Until findoftp = 0
End Function
Public Function FindWordPad() As Long
Dim Counter As Long
Dim RICHEDIT As Long
Dim i As Long
Dim AfxControlBar As Long
Dim WordPadClass As Long
WordPadClass& = FindWindow("WordPadClass", vbNullString)
AfxControlBar& = FindWindowEx(WordPadClass&, 0&, "AfxControlBar", vbNullString)
For i& = 1& To 5&
    AfxControlBar& = FindWindowEx(WordPadClass&, AfxControlBar&, "AfxControlBar", vbNullString)
Next i&
RICHEDIT& = FindWindowEx(WordPadClass&, 0&, "RICHEDIT", vbNullString)
Do While (Counter& <> 100&) And (AfxControlBar& = 0& Or RICHEDIT& = 0&): DoEvents
    WordPadClass& = FindWindowEx(WordPadClass&, WordPadClass&, "WordPadClass", vbNullString)
    AfxControlBar& = FindWindowEx(WordPadClass&, 0&, "AfxControlBar", vbNullString)
    For i& = 1& To 5&
        AfxControlBar& = FindWindowEx(WordPadClass&, AfxControlBar&, "AfxControlBar", vbNullString)
    Next i&
    RICHEDIT& = FindWindowEx(WordPadClass&, 0&, "RICHEDIT", vbNullString)
    If AfxControlBar& And RICHEDIT& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindWordPad& = WordPadClass&
    Exit Function
End If
End Function
Public Function OpenAndSetTextToWordPad(Text As String) As Long
Dim RICHEDIT As Long
Dim WordPadClass As Long
Dim OpenWP As Long
OpenWP& = Shell("C:\windows\write.exe", 1)
Do
TimeNow 0.4
Loop Until FindWordPad <> 0
WordPadClass& = FindWindow("WordPadClass", vbNullString)
RICHEDIT& = FindWindowEx(WordPadClass&, 0&, "RICHEDIT", vbNullString)
Call SendMessageByString(RICHEDIT&, WM_SETTEXT, 0&, "")
Call SendMessageByString(RICHEDIT&, WM_SETTEXT, 0&, Text)
End Function
Public Function FindSignOnAFriend() As Long
Dim Counter As Long
Dim AOLIcon3 As Long
Dim AOLStatic2 As Long
Dim AOLIcon2 As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 3&
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLIcon& = 0& Or AOLStatic& = 0& Or AOLIcon2& = 0& Or AOLStatic2& = 0& Or AOLIcon3& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 3&
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    If AOLIcon& And AOLStatic& And AOLIcon2& And AOLStatic2& And AOLIcon3& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindSignOnAFriend& = AOLChild&
    Exit Function
End If
End Function
Public Function FindSignOnAFriendInfo() As Long
Dim Counter As Long
Dim AOLStatic13 As Long
Dim AOLButton As Long
Dim AOLStatic12 As Long
Dim AOLEdit10 As Long
Dim AOLStatic11 As Long
Dim AOLEdit9 As Long
Dim AOLStatic10 As Long
Dim AOLCombobox As Long
Dim AOLStatic9 As Long
Dim AOLEdit8 As Long
Dim AOLStatic8 As Long
Dim AOLEdit7 As Long
Dim AOLStatic7 As Long
Dim AOLEdit6 As Long
Dim AOLStatic6 As Long
Dim AOLEdit5 As Long
Dim AOLStatic5 As Long
Dim AOLEdit4 As Long
Dim AOLStatic4 As Long
Dim AOLEdit3 As Long
Dim AOLStatic3 As Long
Dim AOLEdit2 As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLStatic3& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLEdit3& = FindWindowEx(AOLModal&, AOLEdit2&, "_AOL_Edit", vbNullString)
AOLStatic4& = FindWindowEx(AOLModal&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLEdit4& = FindWindowEx(AOLModal&, AOLEdit3&, "_AOL_Edit", vbNullString)
AOLStatic5& = FindWindowEx(AOLModal&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLEdit5& = FindWindowEx(AOLModal&, AOLEdit4&, "_AOL_Edit", vbNullString)
AOLStatic6& = FindWindowEx(AOLModal&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLEdit6& = FindWindowEx(AOLModal&, AOLEdit5&, "_AOL_Edit", vbNullString)
AOLStatic7& = FindWindowEx(AOLModal&, AOLStatic6&, "_AOL_Static", vbNullString)
AOLEdit7& = FindWindowEx(AOLModal&, AOLEdit6&, "_AOL_Edit", vbNullString)
AOLStatic8& = FindWindowEx(AOLModal&, AOLStatic7&, "_AOL_Static", vbNullString)
AOLEdit8& = FindWindowEx(AOLModal&, AOLEdit7&, "_AOL_Edit", vbNullString)
AOLStatic9& = FindWindowEx(AOLModal&, AOLStatic8&, "_AOL_Static", vbNullString)
AOLCombobox& = FindWindowEx(AOLModal&, 0&, "_AOL_Combobox", vbNullString)
AOLStatic10& = FindWindowEx(AOLModal&, AOLStatic9&, "_AOL_Static", vbNullString)
AOLEdit9& = FindWindowEx(AOLModal&, AOLEdit8&, "_AOL_Edit", vbNullString)
AOLStatic11& = FindWindowEx(AOLModal&, AOLStatic10&, "_AOL_Static", vbNullString)
AOLEdit10& = FindWindowEx(AOLModal&, AOLEdit9&, "_AOL_Edit", vbNullString)
AOLStatic12& = FindWindowEx(AOLModal&, AOLStatic11&, "_AOL_Static", vbNullString)
AOLButton& = FindWindowEx(AOLModal&, 0&, "_AOL_Button", vbNullString)
AOLButton& = FindWindowEx(AOLModal&, AOLButton&, "_AOL_Button", vbNullString)
AOLStatic13& = FindWindowEx(AOLModal&, AOLStatic12&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLEdit2& = 0& Or AOLStatic3& = 0& Or AOLEdit3& = 0& Or AOLStatic4& = 0& Or AOLEdit4& = 0& Or AOLStatic5& = 0& Or AOLEdit5& = 0& Or AOLStatic6& = 0& Or AOLEdit6& = 0& Or AOLStatic7& = 0& Or AOLEdit7& = 0& Or AOLStatic8& = 0& Or AOLEdit8& = 0& Or AOLStatic9& = 0& Or AOLCombobox& = 0& Or AOLStatic10& = 0& Or AOLEdit9& = 0& Or AOLStatic11& = 0& Or AOLEdit10& = 0& Or AOLStatic12& = 0& Or AOLButton& = 0& Or AOLStatic13& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLStatic3& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLEdit3& = FindWindowEx(AOLModal&, AOLEdit2&, "_AOL_Edit", vbNullString)
    AOLStatic4& = FindWindowEx(AOLModal&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLEdit4& = FindWindowEx(AOLModal&, AOLEdit3&, "_AOL_Edit", vbNullString)
    AOLStatic5& = FindWindowEx(AOLModal&, AOLStatic4&, "_AOL_Static", vbNullString)
    AOLEdit5& = FindWindowEx(AOLModal&, AOLEdit4&, "_AOL_Edit", vbNullString)
    AOLStatic6& = FindWindowEx(AOLModal&, AOLStatic5&, "_AOL_Static", vbNullString)
    AOLEdit6& = FindWindowEx(AOLModal&, AOLEdit5&, "_AOL_Edit", vbNullString)
    AOLStatic7& = FindWindowEx(AOLModal&, AOLStatic6&, "_AOL_Static", vbNullString)
    AOLEdit7& = FindWindowEx(AOLModal&, AOLEdit6&, "_AOL_Edit", vbNullString)
    AOLStatic8& = FindWindowEx(AOLModal&, AOLStatic7&, "_AOL_Static", vbNullString)
    AOLEdit8& = FindWindowEx(AOLModal&, AOLEdit7&, "_AOL_Edit", vbNullString)
    AOLStatic9& = FindWindowEx(AOLModal&, AOLStatic8&, "_AOL_Static", vbNullString)
    AOLCombobox& = FindWindowEx(AOLModal&, 0&, "_AOL_Combobox", vbNullString)
    AOLStatic10& = FindWindowEx(AOLModal&, AOLStatic9&, "_AOL_Static", vbNullString)
    AOLEdit9& = FindWindowEx(AOLModal&, AOLEdit8&, "_AOL_Edit", vbNullString)
    AOLStatic11& = FindWindowEx(AOLModal&, AOLStatic10&, "_AOL_Static", vbNullString)
    AOLEdit10& = FindWindowEx(AOLModal&, AOLEdit9&, "_AOL_Edit", vbNullString)
    AOLStatic12& = FindWindowEx(AOLModal&, AOLStatic11&, "_AOL_Static", vbNullString)
    AOLButton& = FindWindowEx(AOLModal&, 0&, "_AOL_Button", vbNullString)
    AOLButton& = FindWindowEx(AOLModal&, AOLButton&, "_AOL_Button", vbNullString)
    AOLStatic13& = FindWindowEx(AOLModal&, AOLStatic12&, "_AOL_Static", vbNullString)
    If AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLStatic3& And AOLEdit3& And AOLStatic4& And AOLEdit4& And AOLStatic5& And AOLEdit5& And AOLStatic6& And AOLEdit6& And AOLStatic7& And AOLEdit7& And AOLStatic8& And AOLEdit8& And AOLStatic9& And AOLCombobox& And AOLStatic10& And AOLEdit9& And AOLStatic11& And AOLEdit10& And AOLStatic12& And AOLButton& And AOLStatic13& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindSignOnAFriendInfo& = AOLModal&
    Exit Function
End If
End Function
Public Function SignOnAFriend(FirstName As String, LastName As String, Street1 As String, Street2 As String, City As String, State As String, ZipCode As String, DayTimePhone As String, EveningPhone As String) As Long
Dim AOLButton As Long
Dim i As Long
Dim AOLEdit As Long
Dim AOLModal As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AppActivate GetWindowCaption(FindAOL)
SendKeys "%po"
Do
TimeNow 0.4
Loop Until FindSignOnAFriend <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindSignOnAFriend))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindSignOnAFriendInfo <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, FirstName)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, LastName)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Street1)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 3&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Street2)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 5&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, City)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 6&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, State)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 7&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ZipCode)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 8&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, DayTimePhone)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 9&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, EveningPhone)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLButton& = FindWindowEx(AOLModal&, 0&, "_AOL_Button", vbNullString)
Do
Call PostMessage(AOLButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLButton&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindSignOnAFriendInfo = 0
If FindSignOnAFriendConfirmation <> 0 Then
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindSignOnAFriendConfirmation))
AOLButton& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
Do
Call PostMessage(AOLButton&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLButton&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindSignOnAFriendConfirmation = 0
End If
If FindSignOnAFriend <> 0 Then
Do
Call CloseWindow(FindSignOnAFriend)
Loop Until FindSignOnAFriend = 0
End If
End Function
Public Function FindSignOnAFriendConfirmation() As Long
Dim Counter As Long
Dim AOLButton As Long
Dim AOLStatic2 As Long
Dim AOLView As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLButton& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLView& = 0& Or AOLStatic2& = 0& Or AOLButton& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLButton& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
    If AOLStatic& And AOLView& And AOLStatic2& And AOLButton& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindSignOnAFriendConfirmation& = AOLChild&
    Exit Function
End If
End Function
Public Function KillDupesListBox(List As ListBox) As Long
Dim DontKill As Long
Dim Kill As Long
For DontKill = 0 To List.ListCount - 1
For Kill = 0 To List.ListCount - 1
If LCase(List.List(DontKill)) Like LCase(List.List(Kill)) And DontKill <> Kill Then
List.RemoveItem (Kill)
End If
Next Kill
Next DontKill
End Function
Public Function FindChatPrefs() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim i As Long
Dim AOLCheckbox As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 4&
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLCheckbox& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
    For i& = 1& To 4&
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    Next i&
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLCheckbox& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindChatPrefs& = AOLModal&
    Exit Function
End If
End Function
Public Function ChatWelcomeOn() As Long
Dim CbVal As Long
Dim AOLCheckbox As Long
Dim AOLModal As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 12&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindChatPrefs <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
CbVal& = SendMessage(AOLCheckbox&, BM_GETCHECK, 0&, vbNullString)
If CbVal& = "0" Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindChatPrefs = 0
End If
If CbVal& = "1" Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindChatPrefs = 0
End If
End Function
Public Function ChatWelcomeOff() As Long
Dim CbVal As Long
Dim AOLCheckbox As Long
Dim AOLModal As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindChat))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 12&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindChatPrefs <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
CbVal& = SendMessage(AOLCheckbox&, BM_GETCHECK, 0&, vbNullString)
If CbVal& = "1" Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindChatPrefs = 0
End If
If CbVal& = "0" Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindChatPrefs = 0
End If
End Function
Public Function MoveMouse(Position1 As String, Position2 As String) As Long
Dim dsk As Long
Dim dsk2 As Long
Dim dsk3 As Long
dsk& = (Position1)
dsk2& = (Position2)
dsk3& = SetCursorPos(dsk, dsk2)
End Function
Public Function FindAngelFirePopUp() As Long
Dim Counter As Long
Dim AOLGlyph As Long
Dim AOLGauge As Long
Dim AOLStatic As Long
Dim AOLWWWView As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLWWWView& = FindWindowEx(AOLChild&, 0&, "_AOL_WWWView", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLGauge& = FindWindowEx(AOLChild&, 0&, "_AOL_Gauge", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
Do While (Counter& <> 100&) And (AOLWWWView& = 0& Or AOLStatic& = 0& Or AOLGauge& = 0& Or AOLGlyph& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLWWWView& = FindWindowEx(AOLChild&, 0&, "_AOL_WWWView", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLGauge& = FindWindowEx(AOLChild&, 0&, "_AOL_Gauge", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    If AOLWWWView& And AOLStatic& And AOLGauge& And AOLGlyph& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAngelFirePopUp& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAngelFirePopUp2() As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Welcome to an Angelfire Member Page!")
End Function
Public Function RunPopupMenu(X As Long, Y As Long, SubMenu As Boolean, Z As Long)
Dim AOLFrame As Long, TextLen As Long, AOLToolbar As Long
Dim AOLToolbar2 As Long, PopMenu As Long, PopMenuVis As Long, i As Long
Dim AOLFrameTxt As String
Dim CursorPos As POINTAPI
Call GetCursorPos(CursorPos)
Call SetCursorPos(Screen.Width, Screen.Height)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
TextLen& = SendMessage(AOLFrame&, WM_GETTEXTLENGTH, 0&, 0&)
AOLFrameTxt$ = String(TextLen&, 0&)
Call SendMessageByString(AOLFrame&, WM_GETTEXT, TextLen& + 1&, AOLFrameTxt$)
AppActivate AOLFrameTxt$
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
If X& = 1& Then
    AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
Else
    AOLIcon& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1 To X& - 1&
        AOLIcon& = FindWindowEx(AOLToolbar2&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
End If
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Do: DoEvents
    PopMenu& = FindWindow("#32768", vbNullString)
    PopMenuVis& = IsWindowVisible(PopMenu&)
Loop Until PopMenuVis& = 1&
For i& = 1& To Y&
    Call PostMessage(PopMenu&, WM_KEYDOWN, VK_DOWN, 0&)
    Call PostMessage(PopMenu&, WM_KEYUP, VK_DOWN, 0&)
Next i&
If SubMenu = True Then
    Call PostMessage(PopMenu&, WM_KEYDOWN, VK_RIGHT, 0&)
    Call PostMessage(PopMenu&, WM_KEYUP, VK_RIGHT, 0&)
    For i& = 1& To Z& - 1&
        Call PostMessage(PopMenu&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(PopMenu&, WM_KEYUP, VK_DOWN, 0&)
    Next i&
End If
Call PostMessage(PopMenu&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(PopMenu&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CursorPos.X, CursorPos.Y)
End Function
Public Function WebSearch(Text As String, Summaries As Boolean)
If Summaries = True Then
Call AOLKeyword("http://www.webcrawler.com/cgi-bin/WebQuery?search=" & Text & "&showSummary=true&start=0&perPage=25")
End If
If Summaries = False Then
Call AOLKeyword("http://www.webcrawler.com/cgi-bin/WebQuery?search=" & Text & "&src=wc_results&showSummary=false")
End If
End Function
Public Function FreeInternet(UserID As String, Password As String)
' this is like netzero, but it's FreeI.Net

Call AOLKeyword("http://survey.blah.freei.net/ok.asp?EID=" & UserID & "&x1=" & Password & "&market=NY&os=98")
End Function
Public Function FindAOLCASN() As Long
Dim Counter As Long
Dim AOLGlyph6 As Long
Dim AOLIcon4 As Long
Dim AOLGlyph5 As Long
Dim AOLIcon3 As Long
Dim AOLGlyph4 As Long
Dim AOLIcon2 As Long
Dim AOLGlyph3 As Long
Dim i As Long
Dim AOLStatic2 As Long
Dim AOLGlyph2 As Long
Dim AOLStatic As Long
Dim AOLGlyph As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLGlyph2& = FindWindowEx(AOLChild&, AOLGlyph&, "_AOL_Glyph", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
For i& = 1& To 4&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
Next i&
AOLGlyph3& = FindWindowEx(AOLChild&, AOLGlyph2&, "_AOL_Glyph", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLGlyph4& = FindWindowEx(AOLChild&, AOLGlyph3&, "_AOL_Glyph", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLGlyph5& = FindWindowEx(AOLChild&, AOLGlyph4&, "_AOL_Glyph", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
Next i&
AOLGlyph6& = FindWindowEx(AOLChild&, AOLGlyph5&, "_AOL_Glyph", vbNullString)
Do While (Counter& <> 100&) And (AOLIcon& = 0& Or AOLGlyph& = 0& Or AOLStatic& = 0& Or AOLGlyph2& = 0& Or AOLStatic2& = 0& Or AOLGlyph3& = 0& Or AOLIcon2& = 0& Or AOLGlyph4& = 0& Or AOLIcon3& = 0& Or AOLGlyph5& = 0& Or AOLIcon4& = 0& Or AOLGlyph6& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    AOLGlyph2& = FindWindowEx(AOLChild&, AOLGlyph&, "_AOL_Glyph", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    For i& = 1& To 4&
        AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    Next i&
    AOLGlyph3& = FindWindowEx(AOLChild&, AOLGlyph2&, "_AOL_Glyph", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLGlyph4& = FindWindowEx(AOLChild&, AOLGlyph3&, "_AOL_Glyph", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    AOLGlyph5& = FindWindowEx(AOLChild&, AOLGlyph4&, "_AOL_Glyph", vbNullString)
    AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
    Next i&
    AOLGlyph6& = FindWindowEx(AOLChild&, AOLGlyph5&, "_AOL_Glyph", vbNullString)
    If AOLIcon& And AOLGlyph& And AOLStatic& And AOLGlyph2& And AOLStatic2& And AOLGlyph3& And AOLIcon2& And AOLGlyph4& And AOLIcon3& And AOLGlyph5& And AOLIcon4& And AOLGlyph6& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLCASN& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAOLCASNMaster() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLCheckbox As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
RICHCNTL& = FindWindowEx(AOLModal&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLCheckbox& = 0& Or RICHCNTL& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
    RICHCNTL& = FindWindowEx(AOLModal&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLCheckbox& And RICHCNTL& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLCASNMaster& = AOLModal&
    Exit Function
End If
End Function
Public Function FindAOLCASN2() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLIcon& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And RICHCNTL& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLCASN2& = AOLChild&
    Exit Function
End If
End Function
Public Function FindAOLCASNStep1() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLStatic2 As Long
Dim RICHCNTL As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLStatic2& = 0& Or AOLEdit& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And RICHCNTL& And AOLStatic2& And AOLEdit& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLCASNStep1& = AOLModal&
    Exit Function
End If
End Function
Public Function FindAOLCASNStep2() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLStatic2 As Long
Dim RICHCNTL As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLStatic2& = 0& Or AOLEdit& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And RICHCNTL& And AOLStatic2& And AOLEdit& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLCASNStep2& = AOLModal&
    Exit Function
End If
End Function
Public Function FindAOLCASNStep3() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim AOLCheckbox As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 4&
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
For i& = 1& To 3&
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
Next i&
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLCheckbox& = 0& Or AOLStatic2& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
    For i& = 1& To 4&
        AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    For i& = 1& To 3&
        AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic2&, "_AOL_Static", vbNullString)
    Next i&
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLCheckbox& And AOLStatic2& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLCASNStep3& = AOLModal&
    Exit Function
End If
End Function
Public Function FindAOLCASNStep4() As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And RICHCNTL& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLCASNStep4& = AOLModal&
    Exit Function
End If
End Function
Public Function CreateAScreenName(ScreenName As String, Password As String, General As Boolean, Mature As Boolean, YoungTeen As Boolean, KidsOnly As Boolean, Master As Boolean) As Long
' this works well, but if there is a message from aol saying "screen name is used etc..."
' then there will be a problem for this, because i haven't added a code for the look out of that
' so it can click it and continue with a second one. so be aware if that shows up, close down the
' program. or whatever is using this function.

Dim i As Long
Dim AOLCheckbox As Long
Dim AOLEdit As Long
Dim AOLModal As Long
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
If Master = True Then
Call AOLKeyword("Names")
Do
TimeNow 0.4
Loop Until FindAOLCASN <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindAOLCASN))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASN2 <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindAOLCASN2))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep1 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ScreenName)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep2 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Password)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Password)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep3 <> 0
If General = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNMaster <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
If Mature = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
If YoungTeen = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 2&
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
If KidsOnly = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
Call CloseWindow(FindAOLCASN)
If Master = False Then
Call AOLKeyword("Names")
Do
TimeNow 0.4
Loop Until FindAOLCASN <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindAOLCASN))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASN2 <> 0
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", GetWindowCaption(FindAOLCASN2))
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep1 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ScreenName)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep2 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Password)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Password)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep3 <> 0
If General = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNMaster <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
If Mature = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
If YoungTeen = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 2&
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
If KidsOnly = True Then
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLCheckbox& = FindWindowEx(AOLModal&, 0&, "_AOL_Checkbox", vbNullString)
For i& = 1& To 3&
    AOLCheckbox& = FindWindowEx(AOLModal&, AOLCheckbox&, "_AOL_Checkbox", vbNullString)
Next i&
Call PostMessage(AOLCheckbox&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLCheckbox&, WM_LBUTTONUP, 0&, 0&)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Loop Until FindAOLCASNStep4 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLCASNStep4 = 0
End If
Call CloseWindow(FindAOLCASN)
End Function
Public Function FindAOLPasswordStep1() As Long
Dim Counter As Long
Dim AOLStatic2 As Long
Dim AOLIcon As Long
Dim RICHCNTL As Long
Dim AOLStatic As Long
Dim AOLGlyph As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
Do While (Counter& <> 100&) And (AOLGlyph& = 0& Or AOLStatic& = 0& Or RICHCNTL& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLGlyph& = FindWindowEx(AOLModal&, 0&, "_AOL_Glyph", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLModal&, 0&, "RICHCNTL", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    If AOLGlyph& And AOLStatic& And RICHCNTL& And AOLIcon& And AOLStatic2& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLPasswordStep1& = AOLModal&
    Exit Function
End If
End Function
Public Function FindAOLPasswordStep2() As Long
Dim Counter As Long
Dim AOLIcon As Long
Dim AOLEdit2 As Long
Dim AOLStatic2 As Long
Dim AOLEdit As Long
Dim i As Long
Dim AOLStatic As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
For i& = 1& To 2&
    AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
Next i&
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit2&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Do While (Counter& <> 100&) And (AOLStatic& = 0& Or AOLEdit& = 0& Or AOLStatic2& = 0& Or AOLEdit2& = 0& Or AOLIcon& = 0&): DoEvents
    AOLModal& = FindWindowEx(AOLModal&, AOLModal&, "_AOL_Modal", vbNullString)
    AOLStatic& = FindWindowEx(AOLModal&, 0&, "_AOL_Static", vbNullString)
    For i& = 1& To 2&
        AOLStatic& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i&
    AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
    AOLStatic2& = FindWindowEx(AOLModal&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
    AOLEdit2& = FindWindowEx(AOLModal&, AOLEdit2&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
    AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
    If AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLIcon& Then Exit Do
    Counter& = Val(Counter&) + 1&
Loop
If Val(Counter&) < 100& Then
    FindAOLPasswordStep2& = AOLModal&
    Exit Function
End If
End Function
Public Function ChangePassword(OldPassword As String, NewPassword As String) As Long
Dim i As Long
Dim AOLEdit As Long
Dim AOLIcon As Long
Dim AOLModal As Long
Call AOLKeyword("Password")
Do
TimeNow 0.4
Loop Until FindAOLPasswordStep1 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLPasswordStep2 <> 0
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, OldPassword)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, NewPassword)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
For i& = 1& To 2&
    AOLEdit& = FindWindowEx(AOLModal&, AOLEdit&, "_AOL_Edit", vbNullString)
Next i&
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, "")
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, NewPassword)
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
Do
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
TimeNow 0.4
Loop Until FindAOLPasswordStep2 = 0

End Function
