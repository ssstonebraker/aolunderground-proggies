Attribute VB_Name = "TornadoX"
'--------THIS FILE IS NOT DONE AND NEEDS ADJUSTMENTS-----
'DOWNLOAD UTILITIES VERSION
'TornadoX.bas by koko
'This is a nice bas file I made
'With NO fade functions/subs.
'It has subs and functions
'for AOL4.0, AOL3.0, mIRC,
'Windows, AIM, and more.
'E-Mail:    k0ko@hotmail.com
'AIM:       Born In East LA
'AOL:       hye koko
'Phone:     1-818-555-6962
'Pager:     1-818-555-4260
';;;;;;;;    ,;;;;;,    ;;;;;;;;,   ;;,     ;;     ;;;     ;;;;;;;,      ,;;;;;,        ´;;     ;;´
'   ;;      ;;´   ´;;   ;;     ´;;  ;;;;    ;;    ,;´;,    ;;    ´;;,   ;;´   ´;;         ;;   ;;
'   ;;     ;;       ;;  ;;     ,;;  ;;´;;   ;;    ;; ;;    ;;      ;;  ;;       ;;         ;;,;;
'   ;;     ;;       ;;  ;;;;;;;;´   ;; ´;;  ;;   ;;   ;;   ;;      ;;  ;;       ;;          ;;;
'   ;;     ;;       ;;  ;;   ´;;    ;;   ;;,;;  ,;;;;;;;,  ;;      ;;  ;;       ;;         ;;´;;
'   ;;      ;;,   ,;;   ;;     ;;,  ;;    ;;;;  ;;     ;;  ;;     ;;´   ;;,   ,;;         ;;   ;;
'   ;;       ´;;;;;´    ;;      ´;, ;;     ;;; ;;´     ´;; ;;;;;;;´      ´;;;;;´        ,;;     ;;,
'Macro by Hourglass Converter by Sandman.
Option Explicit
Dim mouseIsDown As Boolean
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function SetFocusApi Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const VK_RIGHT = &H27
Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_CLOSE = &H10
Public Const SW_HIDE = 0
Public Const WM_SETTEXT = &HC
Public Const WM_MOVE = &HF012
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONUP = &H202
Public Const WM_CHAR = &H102
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const ENTER_KEY = 13
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Function AOL40FindRoom() As Long
Dim Counter
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim AOLListbox As Long
Dim AOLStatic3 As Long
Dim AOLImage As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim RICHCNTL As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim i As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 3
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i = 1 To 2
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i = 1 To 6
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i
Do While (Counter <> 100) And (AOLStatic& = 0 Or RICHCNTL& = 0 Or AOLCombobox& = 0 Or AOLIcon& = 0 Or AOLStatic2& = 0 Or RICHCNTL2& = 0 Or AOLIcon2& = 0 Or AOLImage& = 0 Or AOLStatic3& = 0 Or AOLListbox& = 0 Or AOLStatic4& = 0 Or AOLIcon3& = 0): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i = 1 To 3
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(AOLChild&, RICHCNTL&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i = 1 To 2
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i = 1 To 6
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i
    If AOLStatic& And RICHCNTL& And AOLCombobox& And AOLIcon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And AOLImage& And AOLStatic3& And AOLListbox& And AOLStatic4& And AOLIcon3& Then Exit Do
    Counter = Val(Counter) + 1
Loop
If Val(Counter) < 100 Then
    AOL40FindRoom& = AOLChild&
    Exit Function
End If
End Function



Public Sub AOL40SendChat(Text As String)
'Sends text in a chat room on AOL4.0
'AOL40SendChat("TornadoX.bas")
Dim Room As Long
Dim Rich As Long
Dim Rich2 As Long
Dim chatroom As Long
chatroom& = AOL40FindRoom&
Rich& = FindWindowEx(chatroom&, 0&, "RICHCNTL", vbNullString)
Rich2& = FindWindowEx(chatroom&, Rich, "RICHCNTL", vbNullString)
Call SendMessageByString(Rich2, WM_SETTEXT, 0&, Text$)
Call SendMessageLong(Rich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub AOL40ScreennameExploit(ScreenName As String)
'It exploits screennames so they will be over 16 characters in the text.
'Call AOL40ScreennameExploit("Tornado X Ownz Me")
Dim AOLEdit As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLEdit& = FindWindowEx(AOLModal&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, ScreenName$)
End Sub
Public Sub AOL40ClickSignOn()
'Clicks on the Sign On button
'Call AOL40ClickSignOn
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim i As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Sign On")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1 To 2
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub WindowsClickStart()
'Clicks the Windows 98 start button
'Call WindowsClickStart
Dim Button As Long
Dim ShellTrayWnd As Long
ShellTrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(ShellTrayWnd&, 0&, "Button", vbNullString)
Call PostMessage(Button&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(Button&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function GetCaption(WindowHandle As Long) As String
    Dim Buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    Buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, Buffer$, TextLength& + 1)
    GetCaption$ = Buffer$
End Function
Public Function GetUser() As String
    Dim AOL As Long, MDI As Long, Welcome As Long
    Dim child As Long, UserString As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    UserString$ = GetCaption(child&)
    If InStr(UserString$, "Welcome, ") = 1 Then
        UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
        GetUser$ = UserString$
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            UserString$ = GetCaption(child&)
            If InStr(UserString$, "Welcome, ") = 1 Then
                UserString$ = Mid$(UserString$, 10, (InStr(UserString$, "!") - 10))
                GetUser$ = UserString$
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    GetUser$ = ""
End Function
Public Function AOL25FindRoom() As Long
Dim Counter
Dim AOLIcon3 As Long
Dim AOLGlyph As Long
Dim AOLIcon2 As Long
Dim AOLStatic2 As Long
Dim AOLListbox As Long
Dim AOLStatic As Long
Dim AOLImage As Long
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLView As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim i As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
For i = 1 To 2
    AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
Next i
AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i = 1 To 3
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i
AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i = 1 To 4
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i
Do While (Counter <> 100) And (AOLView& = 0 Or AOLEdit& = 0 Or AOLIcon& = 0 Or AOLImage& = 0 Or AOLStatic& = 0 Or AOLListbox& = 0 Or AOLStatic2& = 0 Or AOLIcon2& = 0 Or AOLGlyph& = 0 Or AOLIcon3& = 0): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
    AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    For i = 1 To 2
        AOLStatic& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    Next i
    AOLListbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i = 1 To 3
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i
    AOLGlyph& = FindWindowEx(AOLChild&, 0&, "_AOL_Glyph", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i = 1 To 4
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i
    If AOLView& And AOLEdit& And AOLIcon& And AOLImage& And AOLStatic& And AOLListbox& And AOLStatic2& And AOLIcon2& And AOLGlyph& And AOLIcon3& Then Exit Do
    Counter = Val(Counter) + 1
Loop
If Val(Counter) < 100 Then
    AOL25FindRoom& = AOLChild&
    Exit Function
End If
End Function

Public Sub InternetExplorer5GoToSite(URL As String)
'Not done
Dim Edit As Long
Dim ComboBox As Long
Dim ComboBoxEx As Long
Dim ReBarWindow As Long
Dim WorkerA As Long
Dim CabinetWClass As Long
CabinetWClass& = FindWindow("CabinetWClass", vbNullString)
WorkerA& = FindWindowEx(CabinetWClass&, 0&, "WorkerA", vbNullString)
ReBarWindow& = FindWindowEx(WorkerA&, 0&, "ReBarWindow32", vbNullString)
ComboBoxEx& = FindWindowEx(ReBarWindow&, 0&, "ComboBoxEx32", vbNullString)
ComboBox& = FindWindowEx(ComboBoxEx&, 0&, "ComboBox", vbNullString)
Edit& = FindWindowEx(ComboBox&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit&, WM_SETTEXT, 0&, URL$)
Call SendMessageLong(Edit&, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub ListKillDupes(lst As ListBox)
Dim X As Long
Dim i As Long
Dim Newer As String
Dim Current As String
For X& = 0 To lst.ListCount - 1
Current$ = lst.List(X)
For i& = 0 To lst.ListCount - 1
Newer$ = lst.List(i&)
If i& = X Then GoTo dontkill
If Newer$ = Current$ Then lst.RemoveItem (i&)
dontkill:
Next i
Next X
End Sub

Public Sub AOL40Keyword(Keyword As String)
'Goes to a keyword specified
'Call AOL40Keyword("Tornado X")
Dim Edit As Long
Dim AOLCombobox As Long
Dim AOLToolbar2 As Long
Dim AOLToolbar As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
AOLToolbar& = FindWindowEx(AOLFrame&, 0&, "AOL Toolbar", vbNullString)
AOLToolbar2& = FindWindowEx(AOLToolbar&, 0&, "_AOL_Toolbar", vbNullString)
AOLCombobox& = FindWindowEx(AOLToolbar2&, 0&, "_AOL_Combobox", vbNullString)
Edit& = FindWindowEx(AOLCombobox&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit&, WM_SETTEXT, 0&, Keyword$)
Call SendMessageLong(Edit&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(Edit&, WM_CHAR, VK_RETURN, 0&)
End Sub

Public Sub FormAbove(which As Form)
'Keeps the form on top other windows
'Call FormAbove(Me)
Call SetWindowPos(which.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub FormDrag(which As Form)
'Moves the form
'Call FormDrag(Me)
Call ReleaseCapture
Call SendMessage(which.hWnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub

Public Sub AOL30SendChat(Text As String)
'Not done
Dim AOLIcon As Long
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim AOL30FindRoom As String
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, Text$)
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
Call SendMessageLong(AOLEdit&, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub AOL40BuddyListClose()
'Clicks the X and closes the buddylist window.
'Call BuddyListClose
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Buddy List Window")
Call SendMessage(AOLChild&, WM_CLOSE, 0&, 0&)
End Sub

Public Function AOL40MailFlashOpen()
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Incoming/Saved Mail")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Sub AOL40KillDownloadAd()
'Not done
Dim AOLImage As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
Call ShowWindow(AOLImage&, SW_HIDE)
End Sub

Public Sub mIRCSendChat(Text As String)
'Sends text into a mIRC32 chat room
'Call mIRCSendChat("Tornado X")
Dim Edit As Long
Dim channel As Long
Dim MDIClient As Long
Dim mirc As Long
mirc& = FindWindow("mIRC32", vbNullString)
MDIClient& = FindWindowEx(mirc&, 0&, "MDIClient", vbNullString)
channel& = FindWindowEx(MDIClient&, 0&, "channel", vbNullString)
Edit& = FindWindowEx(channel&, 0&, "Edit", vbNullString)
Call SendMessageByString(Edit, WM_SETTEXT, 0&, Text$)
Call SendMessageLong(Edit, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub TimeOut(interval)
'Pauses
'Call TimeOut(.7)
Dim Start As String
Start$ = Timer
Do While Timer - Start$ < interval
DoEvents
Loop
End Sub
Public Sub AOL40ClickCancelSignOn()
'Clicks cancel on signon
'Call AOL40ClickCancelSignOn
Dim AOLIcon As Long
Dim AOLModal As Long
AOLModal& = FindWindow("_AOL_Modal", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLModal&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If
Dim trimspace
For trimspace = 1 To Len(Text)
DoEvents
Dim thechar As String
Dim thechars As String
thechar$ = Mid(Text, trimspace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next trimspace

TrimSpaces = thechars$
End Function
Public Sub AOL40SendLink(URL As Long, Description As Long)
AOL40SendChat ("< A HREF=") + URL + (">") + (Description) + ("</A>")
End Sub
Public Sub AOL40RunMenu(MainMenu As Long, SubMenu As Long)
'Not done
Dim AOLFrame As Long
Dim TpMenu As Long
Dim SbMenu As Long
Dim MenuID As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
TpMenu& = GetMenu(AOLFrame&)
SbMenu& = GetSubMenu(TpMenu&, MainMenu&)
MenuID& = GetMenuItemID(SbMenu&, SubMenu&)
Call SendMessageLong(AOLFrame&, WM_COMMAND, MenuID&, 0&)
End Sub
Public Sub AOL40SignOnWithPW(PW As String)
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Goodbye from America Online!")
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, PW$)
Call SendMessageLong(AOLEdit&, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub AOL40ContinueDownload()
AOL40Keyword ("download manager")
TimeOut (15)
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", "Download Manager")
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Call PostMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)

End Sub
Public Sub AOL40DownloadClickFinishLater()
'Not done
Dim AOLBUTTON As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLBUTTON& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
Call PostMessage(AOLBUTTON&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(AOLBUTTON&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function AOL40DownloadFindWindow() As Long
'Finds the download window
Dim Counter
Dim AOLImage As Long
Dim AOLBUTTON As Long
Dim AOLCheckbox As Long
Dim AOLStatic2 As Long
Dim AOLGauge As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLGauge& = FindWindowEx(AOLChild&, 0&, "_AOL_Gauge", vbNullString)
AOLGauge& = FindWindowEx(AOLChild&, AOLGauge&, "_AOL_Gauge", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLBUTTON& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
AOLBUTTON& = FindWindowEx(AOLChild&, AOLBUTTON&, "_AOL_Button", vbNullString)
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
Do While (Counter <> 100) And (AOLStatic& = 0 Or AOLGauge& = 0 Or AOLStatic2& = 0 Or AOLCheckbox& = 0 Or AOLBUTTON& = 0 Or AOLImage& = 0): DoEvents
AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLGauge& = FindWindowEx(AOLChild&, 0&, "_AOL_Gauge", vbNullString)
AOLGauge& = FindWindowEx(AOLChild&, AOLGauge&, "_AOL_Gauge", vbNullString)
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLBUTTON& = FindWindowEx(AOLChild&, 0&, "_AOL_Button", vbNullString)
AOLBUTTON& = FindWindowEx(AOLChild&, AOLBUTTON&, "_AOL_Button", vbNullString)
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
If AOLStatic& And AOLGauge& And AOLStatic2& And AOLCheckbox& And AOLBUTTON& And AOLImage& Then Exit Do
Counter = Val(Counter) + 1
Loop
If Val(Counter) < 100 Then
AOL40DownloadFindWindow& = AOLChild&
Exit Function
End If
End Function

Public Sub AOL40MailOpenFlash()
Dim AOL As Long
Dim tool As Long
Dim Toolbar As Long
Dim ToolIcon As Long
Dim DoThis As Long
Dim sMod As Long
Dim CurPos As POINTAPI
Dim WinVis As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, 0&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
ToolIcon& = FindWindowEx(Toolbar&, ToolIcon&, "_AOL_Icon", vbNullString)
Call GetCursorPos(CurPos)
Call SetCursorPos(Screen.Width, Screen.Height)
Call PostMessage(ToolIcon&, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(ToolIcon&, WM_LBUTTONUP, 0&, 0&)
Do
sMod& = FindWindow("#32768", vbNullString)
WinVis& = IsWindowVisible(sMod&)
Loop Until WinVis& = 1
Call PostMessage(sMod&, WM_KEYDOWN, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_UP, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RIGHT, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RIGHT, 0&)
Call PostMessage(sMod&, WM_KEYDOWN, VK_RETURN, 0&)
Call PostMessage(sMod&, WM_KEYUP, VK_RETURN, 0&)
Call SetCursorPos(CurPos.X, CurPos.Y)
End Sub
Public Function AOL40MailFlashOpenWindow() As Long
Dim Counter
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
Dim AOLStatic As Long
Dim AOLView As Long
Dim RICHCNTL As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Dim i
For i = 1 To 3
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
Do While (Counter <> 100) And (RICHCNTL& = 0 Or AOLView& = 0 Or AOLStatic& = 0 Or AOLIcon& = 0 Or AOLStatic2& = 0 Or AOLIcon2& = 0 Or AOLStatic3& = 0 Or AOLIcon3& = 0 Or AOLStatic4& = 0 Or AOLIcon4& = 0 Or AOLStatic5& = 0 Or AOLIcon5& = 0 Or AOLStatic6& = 0 Or AOLIcon6& = 0): DoEvents
AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLView& = FindWindowEx(AOLChild&, 0&, "_AOL_View", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i = 1 To 3
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon5& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLIcon6& = FindWindowEx(AOLChild&, AOLIcon5&, "_AOL_Icon", vbNullString)
If RICHCNTL& And AOLView& And AOLStatic& And AOLIcon& And AOLStatic2& And AOLIcon2& And AOLStatic3& And AOLIcon3& And AOLStatic4& And AOLIcon4& And AOLStatic5& And AOLIcon5& And AOLStatic6& And AOLIcon6& Then Exit Do
Counter = Val(Counter) + 1
Loop
If Val(Counter) < 100 Then
AOL40MailFlashOpenWindow& = AOLChild&
Exit Function
End If
End Function
Public Function AOL40MailFlashForwardWindow() As Long
Dim Counter
Dim AOLIcon4 As Long
Dim AOLStatic7 As Long
Dim AOLIcon3 As Long
Dim AOLStatic6 As Long
Dim AOLIcon2 As Long
Dim AOLStatic5 As Long
Dim AOLCheckbox As Long
Dim RICHCNTL As Long
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
Dim i
For i = 1 To 10
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic7& = FindWindowEx(AOLChild&, AOLStatic6&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
Do While (Counter <> 100) And (AOLStatic& = 0 Or AOLEdit& = 0 Or AOLStatic2& = 0 Or AOLEdit2& = 0 Or AOLStatic3& = 0 Or AOLEdit3& = 0 Or AOLFontCombo& = 0 Or AOLStatic4& = 0 Or AOLCombobox& = 0 Or AOLIcon& = 0 Or RICHCNTL& = 0 Or AOLCheckbox& = 0 Or AOLStatic5& = 0 Or AOLIcon2& = 0 Or AOLStatic6& = 0 Or AOLIcon3& = 0 Or AOLStatic7& = 0 Or AOLIcon4& = 0): DoEvents
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
For i = 1 To 10
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
RICHCNTL& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCheckbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Checkbox", vbNullString)
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
AOLStatic6& = FindWindowEx(AOLChild&, AOLStatic5&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLStatic7& = FindWindowEx(AOLChild&, AOLStatic6&, "_AOL_Static", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
AOLIcon4& = FindWindowEx(AOLChild&, AOLIcon4&, "_AOL_Icon", vbNullString)
If AOLStatic& And AOLEdit& And AOLStatic2& And AOLEdit2& And AOLStatic3& And AOLEdit3& And AOLFontCombo& And AOLStatic4& And AOLCombobox& And AOLIcon& And RICHCNTL& And AOLCheckbox& And AOLStatic5& And AOLIcon2& And AOLStatic6& And AOLIcon3& And AOLStatic7& And AOLIcon4& Then Exit Do
Counter = Val(Counter) + 1
Loop
If Val(Counter) < 100 Then
AOL40MailFlashForwardWindow& = AOLChild&
Exit Function
End If
End Function

Public Function AOL40MailFlashClickForward()
Dim AOLIcon As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
Dim i
For i = 1 To 6
AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i
End Function
Public Sub AOL25SendChat(chat As String)
Dim AOLEdit As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
Dim chatroom As String
chatroom$ = AOL25FindRoom
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", chatroom$)
AOLEdit& = FindWindowEx(AOLChild&, 0&, "_AOL_Edit", vbNullString)
Call SendMessageByString(AOLEdit&, WM_SETTEXT, 0&, chat$)
Call SendMessageLong(AOLEdit&, WM_CHAR, ENTER_KEY, 0&)
End Sub
