Attribute VB_Name = "X5"
'#######################################################################
'## X5.bas Created For AOL Version 4.0
'## This module was created with Visual Basic 6 Professional Edition
'## View the ReadMe before using this file
'## This module was created by: pre
'## Any questions which cannot be answered in the ReadMe file should be directed to:
'## pre@dosfx.com
'## Release Monday, April 17, 2000
'## Version 1.0.0
'#######################################################################


Option Explicit

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lparam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Boolean) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function mciExecute Lib "winmm.dll" (ByVal lpstrCommand As String) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wflags As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lparam As String) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const WM_CHAR = &H102
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MOVE = &HF012
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SETCURSOR = &H20
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112
Public Const WM_VSCROLL = &H115

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146

Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1


Public Const PROCESS_ALL_ACCESS = &HFFF
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PROCESS_VM_READ = &H10

Public Const MF_BYPOSITION = &H400&
Public Const MF_REMOVE = &H1000&

Public Const HTCAPTION = 2

Public Enum MailBox
        NEW_MAIL = 0
        OLD_MAIL = 1
        SENT_MAIL = 2
End Enum

Public Enum WINDCOLORS
        COLOR_SCROLLBAR = 0
        COLOR_BACKGROUND = 1
        COLOR_ACTIVECAPTION = 2
        COLOR_INACTIVECAPTION = 3
        COLOR_MENU = 4
        COLOR_WINDOW = 5
        COLOR_WINDOWFRAME = 6
        COLOR_MENUTEXT = 7
        COLOR_WINDOWTEXT = 8
        COLOR_CAPTIONTEXT = 9
        COLOR_ACTIVEBORDER = 10
        COLOR_INACTIVEBORDER = 11
        COLOR_APPWORKSPACE = 12
        COLOR_HIGHLIGHT = 13
        COLOR_HIGHLIGHTTEXT = 14
        COLOR_BTNFACE = 15
        COLOR_BTNSHADOW = 16
        COLOR_GRAYTEXT = 17
        COLOR_BTNTEXT = 18
        COLOR_INACTIVECAPTIONTEXT = 19
        COLOR_BTNHIGHLIGHT = 20
End Enum

Public Enum MENUITEM
        FILE_MNU = 0
        EDIT_MNU = 1
        WINDOW_MNU = 2
        SIGNOFF_MNU = 3
        HELP_MNU = 4
End Enum

Public Enum SUBMENUITEM
        NEW_MNU = 0
        OPEN_MNU = 1
        OPENPICGAL_MNU = 2
        SAVE_MNU = 4
        SAVEAS_MNU = 5
        SAVECAB_MNU = 6
        PRINTSET_MNU = 8
        PRINT_MNU = 9
        STOPINTEXT_MNU = 11
        EXIT_MNU = 12
        UNDO_MNU = 0
        CUT_MNU = 2
        COPY_MNU = 3
        PASTE_MNU = 4
        SELECTALL_MNU = 5
        FINDTOP_MNU = 7
        SPELLCHK_MNU = 8
        DICTIONARY_MNU = 9
        THESAUR_MNU = 10
        CAPPICTURE_MNU = 12
        CASCADE_MNU = 0
        TILE_MNU = 1
        ARRICONS_MNU = 2
        CLOSEALL_MNU = 3
        ADDTOPTOFAV_MNU = 5
        REMWINSIZEONLY_MNU = 6
        REMWINSIZENPOS_MNU = 7
        FORWINSIZENPOS_MNU = 8
        CHANNELS_MNU = 10
        WELCOME_MNU = 11
        SWITCHSN_MNU = 0
        SIGNOFF_MNU = 1
        MEMSERVICES_MNU = 0
        OFFLINEHELP_MNU = 1
        PARENTALCNTLS_MNU = 3
        HELPKEYWORDS_MNU = 5
        ACCTSNBILLING_MNU = 6
        ACCESSNUMS_MNU = 7
        WHATSNEW_MNU = 9
        ABOUTAOL_MNU = 10
End Enum

Public Enum TOOLICON
        READ_MAIL = 0
        WRITE_MAIL = 1
        MAIL_CENTER = 2
        PRINT_I = 3
        MY_FILES = 4
        MY_AOL = 5
        FAVORITES_I = 6
        INTERNET_I = 7
        CHANNELS_I = 8
        PEOPLE_I = 9
        QUOTES_I = 10
        PERKS_I = 11
        WEATHER_I = 12
End Enum

Public Enum DRIVETYPE
        A_DRIVE = 0
        B_DRIVE = 1
        C_DRIVE = 2
        D_DRIVE = 3
        E_DRIVE = 4
        F_DRIVE = 5
        G_DRIVE = 6
        H_DRIVE = 7
        I_DRIVE = 8
        J_DRIVE = 9
        K_DRIVE = 10
        L_DRIVE = 11
        M_DRIVE = 12
        N_DRIVE = 13
        O_DRIVE = 14
        P_DRIVE = 15
        Q_DRIVE = 16
        R_DRIVE = 17
        S_DRIVE = 18
        T_DRIVE = 19
        U_DRIVE = 20
        V_DRIVE = 21
        W_DRIVE = 22
        X_DRIVE = 23
        Y_DRIVE = 24
        Z_DRIVE = 25
End Enum

Public Enum DNTCONV
        GENERALD_CV = 0
        LONGD_CV = 1
        MEDIUMD_CV = 2
        SHORTD_CV = 3
        LONGT_CV = 4
        MEDIUMT_CV = 5
        SHORTT_CV = 6
End Enum

Public Enum ROOM
        PUBLIC_RM = 0
        PRIVATE_RM = 1
End Enum

'Public Variables that return values for TimeSinceStartUp Function
Public strHours_Days As String
Public strMinutes_Total As String
Public strSeconds_Total As String
'~~~

Public lngRichSend As Long
Public strWho As String
Public strLCaseNick As String
Public blnClass As Boolean
Public strWindowHandleAndClass As String
Public lngReturnHandle As Long
Public blnOff As Boolean


Public Function EnumChildProc(ByVal hWnd As Long, ByVal lparam As Long) As Long
Attribute EnumChildProc.VB_Description = "This procedure is used with EnumWindows and loads each childs class and handle to a public variable"
    
    Dim strClass As String
    
    blnClass = False
    
    strClass = GetWindowClassName(hWnd)
    Do
        DoEvents
    Loop Until blnClass = True

    Do
        strWindowHandleAndClass = strWindowHandleAndClass & hWnd
        strWindowHandleAndClass = strWindowHandleAndClass & "," & strClass & "~"
    Loop Until InStr(strWindowHandleAndClass, hWnd) > 0
    DoEvents
    
    EnumChildProc = 1
    
End Function


Public Sub FilterRoomList(Optional blnFindExact As Boolean = False, Optional blnAddToList As Boolean = False, Optional blnAddToCombo As Boolean = False, Optional lstName As ListBox = "", Optional cmbName As ComboBox = "")
Attribute FilterRoomList.VB_Description = "Sorts room names and places them into their proper location"

    'ok, this function is a little tricky, so rather than just give you the code
    'I am going to explain how this works
    'who knows, you may even learn something

    Dim lngChat As Long, lngList As Long, lngCount As Long
    Dim lngGetId As Long, lngHold As Long, lngOpen As Long
    Dim lngPosition As Long, lngData As Long, strTemp As String
    Dim lngBytesWritten As Long, lngMemCopy As Long, strNames As String
    
    lngChat = FindAolChat
    lngList = FindWindowEx(lngChat, 0, "_AOL_Listbox", vbNullString)
    
    If lngList = 0 Then
        Exit Sub
    End If
    
    'we need to set aside room here
    strTemp = String(4, vbNullChar)
    'first we get the number of people in the room by sending getcount message
    lngCount = SendMessage(lngList, LB_GETCOUNT, 0, 0)
    'next we get the identifier of the thread that created the specified window
    lngGetId = GetWindowThreadProcessId(lngList, lngHold)
    'the return value of OpenProcess is an open handle to the specified process(lngHold)
    'we set the inheritance to false because we dont want it to be inherited by
    'a new process, also, we need to specifiy the access to the object with
    'PROCESS_ALL_ACCESS
    lngOpen = OpenProcess(PROCESS_ALL_ACCESS, False, lngHold)
    'we do our loop through
    For lngPosition = 0 To lngCount
        lngData = SendMessage(lngList, LB_GETITEMDATA, ByVal CLng(lngPosition), ByVal 0)
        lngData = lngData + 24
        'ok, we now need to read data from the memory of our process,  remember OpenProcess
        'returned to us the handle we need for this call
        Call ReadProcessMemory(lngOpen, lngData, strTemp, 4, lngBytesWritten)
        'now we need to copy the information from strTemp to lngMemCopy
        Call CopyMemory(lngMemCopy, ByVal strTemp, 4)
        lngMemCopy = lngMemCopy + 6
        'we need to set our buffer size again for the next call
        strTemp = String(16, vbNullChar)
        Call ReadProcessMemory(lngOpen, lngMemCopy, strTemp, Len(strTemp), lngBytesWritten)
        On Error Resume Next
        strTemp = Left(strTemp, InStr(strTemp, vbNullChar) - 1)
        If blnAddToList = False And blnAddToCombo = False Then
            If blnFindExact = True Then
                If LCase(strTemp) = LCase(strWho) Then
                    Call IgnoreSN(strTemp, lngPosition)
                    Exit Sub
                End If
            ElseIf blnFindExact = False Then
                If InStr(LCase(strTemp), LCase(strWho)) > 0 Then
                    Call IgnoreSN(strTemp, lngPosition)
                    Exit Sub
                End If
            End If
        ElseIf blnAddToList = True Then
            lstName.AddItem (strTemp)
        ElseIf blnAddToCombo = True Then
            cmbName.AddItem (strTemp)
        End If
        
        DoEvents
    Next lngPosition
    
    If blnAddToCombo = True Then
        cmbName.Text = cmbName.List(0)
    End If
    
End Sub

Public Sub TextToScreen(strText As String)
Attribute TextToScreen.VB_Description = "Sends text to chat"

    Dim lngChat As Long, lngSize As Long
    
    lngChat = FindAolChat

    If lngChat <> 0 Then
        DoEvents
        Call ApplyText(lngRichSend, strText)
        Do
            lngSize = SendMessage(lngRichSend, WM_GETTEXTLENGTH, 0, 0)
            Call SendMessage(lngRichSend, WM_CHAR, ENTER_KEY, 0)
        Loop Until lngSize = 0
        DoEvents
    End If
    
End Sub

Public Sub WAVPlay(strPath As String)
Attribute WAVPlay.VB_Description = "Plays a wav file"
    
    Dim lngPlay As Long
    
    lngPlay = mciExecute("Play " & strPath)
    
End Sub

Public Function FindAolChat() As Long
Attribute FindAolChat.VB_Description = "Returns the chat room child handle"

    Dim lngMain As Long, lngChild As Long, lngStatic As Long
    Dim lngRich As Long, lngCombo As Long
    
    lngMain = FindFirstWindow
    lngChild = FindWindowEx(lngMain, 0, "AOL Child", vbNullString)
    lngStatic = FindWindowEx(lngChild, 0, "_AOL_Static", vbNullString)
    lngRich = FindWindowEx(lngChild, 0, "RICHCNTL", vbNullString)
    lngCombo = FindWindowEx(lngChild, 0, "_AOL_Combobox", vbNullString)
    
    If lngCombo = 0 Then
        Do
            lngChild = FindWindowEx(lngMain, lngChild, "AOL Child", vbNullString)
            lngStatic = FindWindowEx(lngChild, 0, "_AOL_Static", vbNullString)
            lngRich = FindWindowEx(lngChild, 0, "RICHCNTL", vbNullString)
            lngCombo = FindWindowEx(lngChild, 0, "_AOL_Combobox", vbNullString)
        Loop Until lngChild <> 0 And lngStatic <> 0 And lngRich <> 0 And lngCombo <> 0
    End If
    
    lngRichSend = FindWindowEx(lngChild, lngRich, "RICHCNTL", vbNullString)
    DoEvents
    FindAolChat = lngChild
    
End Function

Public Function FindFirstWindow() As Long
Attribute FindFirstWindow.VB_Description = "Returns the mdi parent handle"

    Dim lngFrame25 As Long
    
    lngFrame25 = FindWindow("AOL Frame25", vbNullString)
    FindFirstWindow = FindWindowEx(lngFrame25, 0, "MDIClient", vbNullString)
    
End Function




Public Function FindSpecificChild(strWindowText As String) As Long
Attribute FindSpecificChild.VB_Description = "Returns the handle of a child specified by caption"

    Dim lngMain As Long, lngReturned As Long
    lngMain = FindWindow("AOL Frame25", vbNullString)
    lngReturned = FindWindowEx(lngMain, 0, "AOL Child", strWindowText)
    If lngReturned = 0 Then
        lngMain = FindFirstWindow
        lngReturned = FindWindowEx(lngMain, 0, "AOL Child", strWindowText)
    End If
    
    FindSpecificChild = lngReturned
    
End Function

Public Sub IgnoreSN(strName As String, Optional lngValue As Long = 1, Optional blnUnignore As Boolean = False, Optional blnFindExact As Boolean = False)

    Dim lngChat As Long, lngList As Long
    Dim lngIgnoreWindow As Long, lngCheck As Long
    Dim lngState As Long
    
    If lngValue = 1 Then
        strWho = strName
        Call FilterRoomList(blnFindExact)
        Exit Sub
    End If
    
    strWho = strName
    strLCaseNick = LCase(strName)
    strLCaseNick = Replace(strLCaseNick, " ", "")
    DoEvents
    
    lngChat = FindAolChat
    lngList = FindWindowEx(lngChat, 0, "_AOL_Listbox", vbNullString)
    DoEvents
    
    Call SendMessage(lngList, LB_SETCURSEL, lngValue, 0)
    Call PostMessage(lngList, WM_LBUTTONDBLCLK, 0, 0)
    DoEvents
    
    Do
        lngIgnoreWindow = FindIgnWindow
    Loop Until lngIgnoreWindow <> 0
    
    lngCheck = FindWindowEx(lngIgnoreWindow, 0, "_AOL_Checkbox", vbNullString)
    Do Until lngCheck > 0
        lngCheck = FindWindowEx(lngIgnoreWindow, 0, "_AOL_Checkbox", vbNullString)
    Loop
    DoEvents
    
    Call PostMessage(lngCheck, WM_LBUTTONDOWN, 0, 0)
    Call PostMessage(lngCheck, WM_LBUTTONUP, 0, 0)
    lngState = SendMessage(lngCheck, BM_GETCHECK, 0, 0)
    
    If blnUnignore = False Then
        If lngState = 0 Then
            Do
                Call PostMessage(lngCheck, WM_LBUTTONDOWN, 0, 0)
                Call PostMessage(lngCheck, WM_LBUTTONUP, 0, 0)
                lngState = SendMessage(lngCheck, BM_GETCHECK, 0, 0)
            Loop Until lngState <> 0
        End If
    ElseIf blnUnignore = True Then
        If lngState = 1 Then
            Do
                Call PostMessage(lngCheck, WM_LBUTTONDOWN, 0, 0)
                Call PostMessage(lngCheck, WM_LBUTTONUP, 0, 0)
                lngState = SendMessage(lngCheck, BM_GETCHECK, 0, 0)
            Loop Until lngState = 0
        End If
    End If
    DoEvents
    
    Call PostMessage(lngIgnoreWindow, WM_CLOSE, 0&, 0&)
    
End Sub

Public Function FindIgnWindow() As Long
Attribute FindIgnWindow.VB_Description = "Returns the info window handle"

    Dim lngMdi As Long, lngInfo As Long
        
    lngMdi = FindFirstWindow
    Do
        lngInfo = SortEnumWindows(lngMdi, , strLCaseNick, True)
    Loop Until lngInfo <> 0
    
    FindIgnWindow = lngInfo
    
End Function

Public Function WinColorToRGB(ByVal lngColor As Long) As Long
Attribute WinColorToRGB.VB_Description = "Returns an rgb value of a long window color"

    Dim intRed As Integer, intGreen As Integer, intBlue As Integer
    
    intRed = lngColor And &HFF
    intGreen = (lngColor \ &H100) And &HFF
    intBlue = (lngColor \ &H10000) And &HFF
    
    WinColorToRGB = intRed & intGreen & intBlue
    
End Function

Public Function WindowCap(lnghWnd As Long) As String
Attribute WindowCap.VB_Description = "Returns the window caption"

    Dim strHolder As String, lngSize As Long
    
    lngSize = GetWindowTextLength(lnghWnd)
    strHolder = String(lngSize, " ")
    DoEvents
    Call GetWindowText(lnghWnd, strHolder, (lngSize + 1))
    
    WindowCap = strHolder
    
End Function

Public Sub OpenMailBox()
Attribute OpenMailBox.VB_Description = "Opens the new mail mailbox"

    Dim lngIcon As Long
    
    lngIcon = FindToolIcon(READ_MAIL)
    
    Call PushButton(lngIcon)
    
End Sub

Public Function CurrentNick() As String
Attribute CurrentNick.VB_Description = "Returns the current screen name"

    Dim lngChild As Long, strCaption As String
    
    Call OpenMenu(WINDOW_MNU, WELCOME_MNU)
    DoEvents
    
    Do
        lngChild = SortEnumWindows(FindFirstWindow, , "Welcome, ", False)
    Loop Until lngChild <> 0
    
    DoEvents
    
    strCaption = WindowCap(lngChild)
    strCaption = Mid(strCaption, 10, Len(strCaption))
    CurrentNick = Clean(strCaption, "!")
    
End Function

Public Function Clean(strText As String, strFilter As String, Optional strReplaceWith As String = "") As String
Attribute Clean.VB_Description = "Replaces a character from a string with a given character"

    Dim intLength As Integer, intLoop As Integer, strLetter As String
    Dim strNewString As String

    intLength = Len(strText)
    For intLoop = 1 To intLength
        strLetter = Mid(strText, intLoop, 1)
        If strLetter <> strFilter Then
            strNewString = strNewString & strLetter
        Else
            strNewString = strNewString & strReplaceWith
        End If
    Next intLoop
    
    Clean = strNewString

End Function

Public Function SearchMailSubject(strSearchString As String, Optional intStart As Integer = 0, Optional blnFindExact As Boolean = False) As Long
Attribute SearchMailSubject.VB_Description = "returns a boolean to a specified search word"

    Dim lngMailWindow As Long
    Dim lngLoop As Long, strSubject As String
    
    strSearchString = LCase(strSearchString)
    
    Do
        lngMailWindow = FindMailListWindow
    Loop Until lngMailWindow <> 0
    DoEvents
    
    For lngLoop = intStart To (GetMailCount - 1)
        strSubject = GetMailSubject(lngLoop)
        DoEvents
        strSubject = LCase(strSubject)
        
        If blnFindExact = True Then
            If strSubject = strSearchString Then
                SearchMailSubject = lngLoop
                Exit Function
            End If
        ElseIf blnFindExact = False Then
            If InStr(strSubject, strSearchString) > 0 Then
                SearchMailSubject = lngLoop
                Exit Function
            End If
        End If
    Next lngLoop
    
    SearchMailSubject = -1
    
End Function

Public Sub SendNewMail(strNicks As String, strSubject As String, strMessage As String, Optional strCC As String = "")
Attribute SendNewMail.VB_Description = "Sends a new email"

    Dim lngIcon As Long, lngMailwin As Long
    Dim lngEdit1 As Long, lngEdit2 As Long, lngEdit3 As Long, lngRich As Long
    Dim intLoop As Integer, lngModal As String
    
    lngIcon = FindToolIcon(WRITE_MAIL)
    DoEvents
    
    Call PushButton(lngIcon)
    
    lngMailwin = FindSpecificChild("Write Mail")
    Do
        lngMailwin = FindSpecificChild("Write Mail")
    Loop Until lngMailwin <> 0
    DoEvents
    
    lngEdit1 = FindWindowEx(lngMailwin, 0, "_AOL_Edit", vbNullString)
    Call ApplyText(lngEdit1, strNicks)
    
    lngEdit2 = FindWindowEx(lngMailwin, lngEdit1, "_AOL_Edit", vbNullString)
    If strCC <> "" Then
        Call ApplyText(lngEdit2, strCC)
    End If
    
    lngEdit3 = FindWindowEx(lngMailwin, lngEdit2, "_AOL_Edit", vbNullString)
    Call ApplyText(lngEdit3, strSubject)
    
    lngRich = FindWindowEx(lngMailwin, 0, "RICHCNTL", vbNullString)
    Call ApplyText(lngRich, strMessage)
    
    lngIcon = FindWindowEx(lngMailwin, 0, "_AOL_Icon", vbNullString)
    For intLoop = 1 To 13
        lngIcon = FindWindowEx(lngMailwin, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    DoEvents
    
    Call PushButton(lngIcon)

    lngModal = FindWindow("_AOL_Modal", vbNullString)
    lngIcon = FindWindowEx(lngModal, 0, "_AOL_Icon", vbNullString)
    
    If lngModal = 0 Then
        Do
            lngModal = FindWindow("_AOL_Modal", vbNullString)
            lngIcon = FindWindowEx(lngModal, 0, "_AOL_Icon", vbNullString)
        Loop Until lngModal <> 0 And lngIcon <> 0
    End If
    DoEvents
    
    Call PushButton(lngIcon)
    
End Sub

Public Sub SendInstantMessage(strNick As String, strMessage As String)
Attribute SendInstantMessage.VB_Description = "Sends an IM"

    Dim lngIcon As Long, lngIM As Long, lngEdit As Long, lngRich As Long
    Dim lngVisible As Long, intLoop As Integer, lngOther As Long
    
    lngIcon = FindToolIcon(PEOPLE_I)
    Call OpenPopUp(lngIcon, 6)
    DoEvents
    
    Do
        lngIM = SortEnumWindows(FindFirstWindow, , "SendInstant", False)
        lngVisible = IsWindowVisible(lngIM)
    Loop Until lngVisible = 1
    DoEvents
    
    lngEdit = FindWindowEx(lngIM, 0, "_AOL_Edit", vbNullString)
    Call ApplyText(lngEdit, strNick)
    
    lngRich = FindWindowEx(lngIM, 0, "RICHCNTL", vbNullString)
    Call ApplyText(lngRich, strMessage)

    lngIcon = FindWindowEx(lngIM, 0, "_AOL_Icon", vbNullString)
    For intLoop = 0 To 7
        lngIcon = FindWindowEx(lngIM, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    DoEvents
    
    Call PushButton(lngIcon)
    
    Do
        lngIM = FindSpecificChild("Send Instant Message")
        lngOther = FindWindow("#32770", "America Online")
    Loop Until lngIM = 0 Or lngOther <> 0
    DoEvents
    
    If lngOther <> 0 Then
        lngIcon = FindWindowEx(lngOther, 0, "Button", vbNullString)
        Call PostMessage(lngIcon, WM_KEYDOWN, VK_SPACE, 0)
        Call PostMessage(lngIcon, WM_KEYUP, VK_SPACE, 0)
        DoEvents
        Call PostMessage(lngIM, WM_CLOSE, 0&, 0&)
    End If
    
End Sub

Public Function EnumWindows(lngPhWnd As Long) As Long
Attribute EnumWindows.VB_Description = "This function is used with EnumChildProc"

    Dim lngCall As Long
    
    lngCall = EnumChildWindows(lngPhWnd, AddressOf EnumChildProc, 0)
    
End Function

Public Function GetImSender() As String
Attribute GetImSender.VB_Description = "Returns an IM sender"

    Dim lngImHandle As Long, strCaption As String, intColon As Integer
    
    strCaption = ""
    
    Do
        lngImHandle = GetImWindow
    Loop Until lngImHandle <> 0
    DoEvents
    
    Do
        strCaption = WindowCap(lngImHandle)
    Loop Until strCaption <> ""
    DoEvents
    
    intColon = InStr(strCaption, ":")
    GetImSender = Mid(strCaption, (intColon + 1), Len(strCaption))
    
End Function

Public Function GetImMessage() As String
Attribute GetImMessage.VB_Description = "Returns an IM  message"

    Dim lngImHandle As Long, lngRich As Long, strMessage As String
    Dim intChar As Integer, intKeep As Integer, strTempData As String
    
    Do
        lngImHandle = GetImWindow
    Loop Until lngImHandle <> 0
    DoEvents
    
    lngRich = FindWindowEx(lngImHandle, 0, "RICHCNTL", vbNullString)
    
    strMessage = RetrieveText(lngRich)
    intChar = InStr(strMessage, Chr(9))
    Do
        intChar = InStr(intChar + 1, strMessage, ":")
        If intChar <> 0 Then
            intKeep = intChar
        End If
    Loop Until intChar = 0
    DoEvents
    
    If intKeep <> 0 Then
        strTempData = Clean(Mid(strMessage, (intKeep + 1), Len(strMessage)), Chr(9), "")
        GetImMessage = Mid(strTempData, 2, Len(strTempData) - 2)
    End If
    
End Function

Public Function GetImWindow() As Long
Attribute GetImWindow.VB_Description = "Returns the handle to the IM window"
    
    Do
        GetImWindow = SortEnumWindows(FindFirstWindow, , "InstantMessage", False)
    Loop Until GetImWindow <> 0
    DoEvents
    
End Function

Public Sub PushButton(lngButtonHandle As Long)
Attribute PushButton.VB_Description = "Sends a message to push an aol icon"

    Call SendMessage(lngButtonHandle, WM_LBUTTONDOWN, 0, 0)
    Call SendMessage(lngButtonHandle, WM_LBUTTONUP, 0, 0)
    
End Sub

Public Sub ApplyText(lnghWnd As Long, strText As String)
Attribute ApplyText.VB_Description = "Sends text to a window"

    Call SendMessageByString(lnghWnd, WM_SETTEXT, 0, strText)
    
End Sub

Public Function SearchMailMessage(strSearchString As String) As Boolean
Attribute SearchMailMessage.VB_Description = "Returns the index of a specified mail subject"
    
    Dim lngMailwin As Long, lngRich As Long
    Dim strMessage As String
    
    Do
        lngMailwin = FindOpenMailWindow
    Loop Until lngMailwin <> 0
    DoEvents
    
    lngRich = FindWindowEx(lngMailwin, 0, "RICHCNTL", vbNullString)
    
    strMessage = RetrieveText(lngRich)
    DoEvents
    
    If InStr(strMessage, strSearchString) > 0 Then
        SearchMailMessage = True
    Else
        SearchMailMessage = False
    End If
    
End Function

Public Sub AVIPlay(strPath As String)
Attribute AVIPlay.VB_Description = "Plays an AVI File"

    Dim lngPlay As Long
    
    lngPlay = mciExecute("Play " & strPath)
    
End Sub

Public Sub TimeSinceStartUp()
Attribute TimeSinceStartUp.VB_Description = "Returns time since computer was started"

    Dim lngMilli As Long
    Dim sngDiv1 As Single, sngDiv2 As Single, sngDiv3 As Single
    Dim intPoint As Integer, lngMinsTotal As Long
    Dim strHours As String, lngHoursTotal As Long, intDays As Integer
    Dim intRemain As Integer, lngHours As Long, lngMins As Long, lngSecs As Long
    Dim strMins As String, strSeconds As String
    
    lngMilli = GetTickCount
    
    sngDiv1 = (lngMilli / 1000)
    sngDiv2 = (sngDiv1 / 60#)
    sngDiv3 = (sngDiv2 / 60#)
    
    intPoint = InStr(sngDiv3, ".")
    If intPoint <> 0 Then
        strHours = Mid(sngDiv3, 1, (intPoint - 1))
    Else
        strHours = sngDiv3
    End If
    
    lngHoursTotal = Val(strHours)
    
    If strHours = 1 Then
        strHours = strHours & " Hour"
    ElseIf strHours > 1 And strHours < 24 Or lngHours = 0 Then
        strHours = strHours & " Hours"
    ElseIf strHours >= 24 Then
        intDays = (lngHoursTotal / 24)
        intRemain = (intDays * 24)
        lngHours = (lngHoursTotal - intRemain)
        If intDays <> 0 Then
            If intDays < 2 Then
                strHours = intDays & " Day, "
            ElseIf intDays > 1 Then
                strHours = intDays & " Days, "
            End If
        End If
        If lngHours <> 0 Then
            If lngHours = 1 Then
                strHours = strHours & lngHours & " Hour"
            ElseIf lngHours > 1 Then
                strHours = strHours & lngHours & " Hours"
            End If
        Else
            strHours = Clean(strHours, ",")
        End If
    End If

    intPoint = InStr(sngDiv2, ".")
    If intPoint <> 0 Then
        strMins = Mid(sngDiv2, 1, (intPoint - 1))
    Else
        strMins = sngDiv2
    End If
    If strMins <> 0 Then
        lngMins = Val(strMins)
        lngMinsTotal = lngMins
        If lngHoursTotal <> 0 Then
            lngHoursTotal = (lngHoursTotal * 60)
            lngMins = (lngMins - lngHoursTotal)
            If lngMins <> 0 Then
                If lngMins < 2 Then
                    strMins = lngMins & " Minute"
                ElseIf lngMins > 1 Then
                    strMins = lngMins & " Minutes"
                End If
            End If
        ElseIf lngHoursTotal = 0 Then
            If lngMins < 2 Then
                    strMins = lngMins & " Minute"
            ElseIf lngMins > 1 Then
                    strMins = lngMins & " Minutes"
            End If
        End If
    End If
    
    intPoint = InStr(sngDiv1, ".")
    If intPoint <> 0 Then
        strSeconds = Mid(sngDiv1, 1, (intPoint - 1))
    Else
        strSeconds = sngDiv1
    End If
    If strSeconds <> 0 Then
        lngSecs = Val(strSeconds)
        If lngSecs <> 0 Then
            lngMinsTotal = (lngMinsTotal * 60)
            lngSecs = (lngSecs - lngMinsTotal)
            If lngSecs <> 0 Then
                If lngSecs < 2 Then
                    strSeconds = lngSecs & " Second"
                ElseIf lngSecs > 1 Then
                    strSeconds = lngSecs & " Seconds"
                End If
            End If
        ElseIf lngSecs = 0 Then
            If lngSecs < 2 Then
                strSeconds = lngSecs & " Second"
            ElseIf lngSecs > 1 Then
                strSeconds = lngSecs & " Seconds"
            End If
        End If
    End If
    
    strHours_Days = strHours
    strMinutes_Total = strMins
    strSeconds_Total = strSeconds
    
End Sub

Public Function GetWindowClassName(lnghWnd As Long) As String

    Dim lngGet As Long, strSize As String
    
    strSize = String(100, " ")
    lngGet = GetClassName(lnghWnd, strSize, 100)
    GetWindowClassName = strSize

    blnClass = True
    
End Function

Public Function SortEnumWindows(ByVal lngPhWnd As Long, Optional strWinClassName As String = "", Optional strWinCaption As String = "", Optional blnFindExact As Boolean = False) As Long
Attribute SortEnumWindows.VB_Description = "Returns the handle to a window specified by class or caption"

    Dim intSpace As Integer, strExtract As String, intComma As String
    Dim strCaption As String, strHandle As String, strClass As String
    Dim strClassN As String, lngHandle As Long
    Dim intLoop As Integer, strLetter As String
    
    If strWinClassName <> "" And strWinCaption <> "" Then
        SortEnumWindows = 0
        Exit Function
    End If

    Call EnumWindows(lngPhWnd)

    If strWinClassName <> "" Then
        strWinClassName = LCase(strWinClassName)
        strWinClassName = Clean(strWinClassName, Chr(32), "")
        DoEvents
    End If
    
    If strWinCaption <> "" Then
        strWinCaption = LCase(strWinCaption)
        strWinCaption = Replace(strWinCaption, " ", "")
        DoEvents
    End If
    
    intSpace = InStr(strWindowHandleAndClass, "~")
    If intSpace <> 0 Then
        Do
            strExtract = Mid(strWindowHandleAndClass, 1, intSpace)
            strWindowHandleAndClass = Replace(strWindowHandleAndClass, strExtract, "", 1, intSpace)
            intComma = InStr(strExtract, ",")
            DoEvents
            If intComma <> 0 Then
                strHandle = Mid(strExtract, 1, (intComma - 1))
                lngHandle = Val(strHandle)
                DoEvents
                strClass = Mid(strExtract, (intComma + 1), Len(strExtract))
                strClassN = GetWindowClassName(lngHandle)
                DoEvents
                If strClass <> strClassN Then
                    strClass = Trim(strClassN)
                End If
                DoEvents
                If strClass <> "" Then
                    strClass = LCase(strClass)
                End If
                strClass = Clean(strClass, Chr(32), "")
                strClass = Clean(strClass, Chr(0), "")
                DoEvents
                If strWinClassName <> "" Then
                    If blnFindExact = False Then
                        If InStr(strClass, strWinClassName) > 0 Then
                            SortEnumWindows = lngHandle
                            Exit Function
                        End If
                    ElseIf blnFindExact = True Then
                        If strClass = strWinClassName Then
                            SortEnumWindows = lngHandle
                            Exit Function
                        End If
                    End If
                ElseIf strWinCaption <> "" Then
                    strCaption = WindowCap(lngHandle)
                    strCaption = LCase(strCaption)
                    strCaption = Replace(strCaption, " ", "")
                    DoEvents
                    If blnFindExact = False Then
                        If InStr(strCaption, strWinCaption) > 0 Then
                            SortEnumWindows = lngHandle
                            Exit Function
                        End If
                    ElseIf blnFindExact = True Then
                        If strCaption = strWinCaption Then
                            SortEnumWindows = lngHandle
                            Exit Function
                        End If
                    End If
                End If
            End If
        Loop Until InStr(strWindowHandleAndClass, " ") = 0
    End If

End Function



Public Function RetrieveText(ByVal lnghWnd As Long) As String
Attribute RetrieveText.VB_Description = "Returns the text from a specified window"

    Dim strHolder As String, lngSize As Long
    
    lngSize = SendMessage(lnghWnd, WM_GETTEXTLENGTH, 0, 0)
    strHolder = String(lngSize, " ")
    
    Call SendMessageByString(lnghWnd, WM_GETTEXT, lngSize + 1, strHolder)
    RetrieveText = strHolder

End Function

Public Sub ForwardMail(lngIndex As Long, strNicks As String, Optional strSubject As String = "", Optional strMessage As String = "")
Attribute ForwardMail.VB_Description = "Forwards an email"

    Dim lngMailWindow As Long, lngMailwin As Long, lngVisible As Long
    Dim intLoop As Integer, lngIcon As Long, lngFwd As Long, lngEdit As Long
    Dim lngRich As Long, lngModal As Long
    
    Call ReadMail(lngIndex)
    DoEvents
    
    Do
        lngMailwin = FindOpenMailWindow
        lngVisible = IsWindowVisible(lngMailwin)
    Loop Until lngVisible = 1 And lngMailwin <> 0
    DoEvents
 
    lngIcon = FindWindowEx(lngMailwin, 0, "_AOL_Icon", vbNullString)
    For intLoop = 1 To 7
        lngIcon = FindWindowEx(lngMailwin, lngIcon, "_AOL_Icon", vbNullString)
        DoEvents
    Next intLoop
    DoEvents
    
    Call PushButton(lngIcon)
    
    Do
        lngFwd = SortEnumWindows(FindFirstWindow, , "Fwd:", False)
    Loop Until lngFwd <> 0
    DoEvents

    Call PostMessage(lngMailwin, WM_CLOSE, 0&, 0&)
    
    lngEdit = FindWindowEx(lngFwd, 0, "_AOL_Edit", vbNullString)
    Call ApplyText(lngEdit, strNicks)
    
    If strSubject <> "" Then
        DoEvents
        lngEdit = FindWindowEx(lngFwd, 0, "_AOL_Edit", vbNullString)
        For intLoop = 1 To 2
            lngEdit = FindWindowEx(lngFwd, lngEdit, "_AOL_Edit", vbNullString)
        Next intLoop
        DoEvents
        Call ApplyText(lngEdit, strSubject)
    End If
    DoEvents
    
    If strMessage <> "" Then
        Do
            lngRich = FindWindowEx(lngFwd, 0, "RICHCNTL", vbNullString)
        Loop Until lngRich <> 0
        Call ApplyText(lngRich, strMessage)
        DoEvents
    End If
    
    lngIcon = FindWindowEx(lngFwd, 0, "_AOL_Icon", vbNullString)
    For intLoop = 1 To 11
        lngIcon = FindWindowEx(lngFwd, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    DoEvents
    Call PushButton(lngIcon)
    
    Do
        lngModal = FindWindow("_AOL_Modal", vbNullString)
        lngIcon = FindWindowEx(lngModal, 0, "_AOL_Icon", vbNullString)
    Loop Until lngIcon <> 0
    DoEvents
    Call PushButton(lngIcon)
    
End Sub

Public Sub MailReply(lngIndex As Long, strMessage As String, Optional strSubject As String = "", Optional blnMailOpen As Boolean = False)

    Dim lngMailList As Long, lngRich As Long
    Dim lngVisible As Long, lngMailRep As Long, lngEdit As Long, intLoop As Integer
    Dim lngIcon As Long, lngModal As Long
    
    If blnMailOpen = False Then
        Call ReadMail(lngIndex)
        DoEvents
        
        Do
            lngMailList = FindOpenMailWindow
            lngVisible = IsWindowVisible(lngMailList)
        Loop Until lngVisible <> 0
        DoEvents
    End If
    
    Do
        lngRich = FindWindowEx(lngMailList, 0, "RICHCNTL", vbNullString)
    Loop Until lngRich <> 0
    DoEvents
    
    Call PostMessage(lngRich, WM_KEYDOWN, VK_RETURN, 0)
    Call PostMessage(lngRich, WM_KEYUP, VK_RETURN, 0)
    
    Do
        lngMailRep = SortEnumWindows(FindFirstWindow, , "Re:", False)
    Loop Until lngMailRep <> 0
    
    Call PostMessage(lngMailList, WM_CLOSE, 0&, 0&)
    DoEvents
    
    If strSubject <> "" Then
        DoEvents
        lngEdit = FindWindowEx(lngMailRep, 0, "_AOL_Edit", vbNullString)
        For intLoop = 1 To 2
            lngEdit = FindWindowEx(lngMailRep, lngEdit, "_AOL_Edit", vbNullString)
        Next intLoop
        DoEvents
        Call ApplyText(lngEdit, strSubject)
    End If
    DoEvents
    
    Do
        lngRich = FindWindowEx(lngMailRep, 0, "RICHCNTL", vbNullString)
    Loop Until lngRich <> 0
    Call ApplyText(lngRich, strMessage)
    
    lngIcon = FindWindowEx(lngMailRep, 0, "_AOL_Icon", vbNullString)
    For intLoop = 1 To 13
        lngIcon = FindWindowEx(lngMailRep, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    DoEvents
    Call PushButton(lngIcon)
    
    Do
        lngModal = FindWindow("_AOL_Modal", vbNullString)
        lngIcon = FindWindowEx(lngModal, 0, "_AOL_Icon", vbNullString)
    Loop Until lngIcon <> 0
    Call PushButton(lngIcon)

End Sub

Public Function FindToolbar() As Long
Attribute FindToolbar.VB_Description = "Returns the handle of the aol toolbar"

    Dim lngMain As Long, lngToolbar As Long, lngTool_bar As Long
    
    lngMain = FindWindow("AOL Frame25", vbNullString)
    lngToolbar = FindWindowEx(lngMain, 0, "AOL Toolbar", vbNullString)
    FindToolbar = FindWindowEx(lngToolbar, 0, "_AOL_Toolbar", vbNullString)

End Function

Public Function FindMailListWindow() As Long
Attribute FindMailListWindow.VB_Description = "Returns the handle to the mail tree window"

    Dim lngMain As Long, lngTabCnt As Long, lngTabPg As Long
    
    Call OpenMailBox
    Do
        lngMain = SortEnumWindows(FindFirstWindow, , "OnlineMailbox", False)
    Loop Until lngMain <> 0
    DoEvents
    
    lngTabCnt = FindWindowEx(lngMain, 0, "_AOL_TabControl", vbNullString)
    lngTabPg = FindWindowEx(lngTabCnt, 0, "_AOL_TabPage", vbNullString)
    
    FindMailListWindow = FindWindowEx(lngTabPg, 0, "_AOL_Tree", vbNullString)
    
End Function


Public Sub OpenMail(intBox As MailBox)
Attribute OpenMail.VB_Description = "Opens a specified mailbox"
    
    Dim lngTool As Long, lngIcon As Long, intLoop As Integer
    
    If intBox = 0 Then
        Call OpenMailBox
        Exit Sub
    End If
    
    lngIcon = FindToolIcon(MAIL_CENTER)
    
    Select Case intBox
        Case Is = 1
            Call OpenPopUp(lngIcon, 4)
            Exit Sub
        Case Is = 2
            Call OpenPopUp(lngIcon, 5)
            Exit Sub
    End Select
    
End Sub

Public Sub NamesToList(lstName As ListBox)

    Call FilterRoomList(, True, , lstName)

End Sub

Public Sub NamesToCombo(cmbName As ComboBox)

    Call FilterRoomList(, , True, , cmbName)
    
End Sub

Public Function FilterMailSubject(strSubject As String) As String
Attribute FilterMailSubject.VB_Description = "Returns the mail subject only, sn and date are stripped"

    Dim lngChar As Long, lngChar1 As Long
    
    lngChar = InStr(strSubject, Chr(9))
    lngChar1 = InStr((lngChar + 1), strSubject, Chr(9))
    
    FilterMailSubject = Mid(strSubject, (lngChar1 + 1), Len(strSubject))
    
End Function

Public Sub DeleteMail(lngIndex As Long)
Attribute DeleteMail.VB_Description = "Deletes a specified mail item"

   Dim lngMailList As Long, lngMailwin As Long, lngIcon As Long, intLoop As Integer
    
    Do
        lngMailList = FindMailListWindow
    Loop Until lngMailList <> 0
    DoEvents
    
    Call SendMessage(lngMailList, LB_SETCURSEL, lngIndex, 0)

    Do
        lngMailwin = SortEnumWindows(FindFirstWindow, , "OnlineMailbox", False)
    Loop Until lngMailwin <> 0
    DoEvents
    
    lngIcon = FindWindowEx(lngMailwin, 0, "_AOL_Icon", vbNullString)
    For intLoop = 0 To 2
        lngIcon = FindWindowEx(lngMailwin, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    DoEvents
    
    Call PushButton(lngIcon)
    
End Sub



Public Sub TimePause(lngSeconds As Long)
Attribute TimePause.VB_Description = "Pause for x amount of seconds"

    Dim lngNow As Long
    
    lngSeconds = (lngSeconds * 1000)
    
    lngSeconds = (GetTickCount + lngSeconds)
    Do Until lngNow >= lngSeconds
        lngNow = GetTickCount
    Loop

End Sub

Public Function GetWindowColor(lnghWnd As Long, lngColor As WINDCOLORS) As Long

    Call SetActiveWindow(lnghWnd)
    GetWindowColor = GetSysColor(lngColor)

End Function

Public Sub OpenPopUp(ByVal lngBhWnd As Long, intKeyDown As Integer, Optional intKeyRight As Integer = 0, Optional intKeyDownEx As Integer = 0)
Attribute OpenPopUp.VB_Description = "Selects a menu item from the aol toolbar popup menus"
    
    Dim lngMenuWindow As Long, lngVisible As Long, intLoop As Integer
    
    If intKeyDown = 0 Then
        Exit Sub
    End If
    
    Call PostMessage(lngBhWnd, WM_LBUTTONDOWN, 0, 0)
    Call PostMessage(lngBhWnd, WM_LBUTTONUP, 0, 0)

    Do
        lngMenuWindow = FindWindow("#32768", vbNullString)
        lngVisible = IsWindowVisible(lngMenuWindow)
    Loop Until lngVisible = 1
    DoEvents
    
    For intLoop = 1 To intKeyDown
        Call PostMessage(lngMenuWindow, WM_KEYDOWN, VK_DOWN, 0)
        Call PostMessage(lngMenuWindow, WM_KEYUP, VK_DOWN, 0)
    Next intLoop
    DoEvents
    
    If intKeyRight <> 0 Then
        For intLoop = 1 To intKeyRight
            Call PostMessage(lngMenuWindow, WM_KEYDOWN, VK_RIGHT, 0)
            Call PostMessage(lngMenuWindow, WM_KEYUP, VK_RIGHT, 0)
        Next intLoop
    End If
    DoEvents
    
    If intKeyDownEx <> 0 Then
        For intLoop = 1 To intKeyDownEx
            Call PostMessage(lngMenuWindow, WM_KEYDOWN, VK_DOWN, 0)
            Call PostMessage(lngMenuWindow, WM_KEYUP, VK_DOWN, 0)
        Next intLoop
    End If
    DoEvents
    
    
    Call PostMessage(lngMenuWindow, WM_KEYDOWN, VK_RETURN, 0)
    Call PostMessage(lngMenuWindow, WM_KEYUP, VK_RETURN, 0)

End Sub

Public Sub OpenMenu(lngmenu As MENUITEM, lngsubmenu As SUBMENUITEM)
Attribute OpenMenu.VB_Description = "Selects a menu item from the aol menu bar"

    Dim lngMain As Long, lngMenuHandle As Long, lngMainMenu As Long
    Dim lngMenuItem As Long
    
    lngMain = FindWindow("AOL Frame25", vbNullString)
    
    lngMenuHandle = GetMenu(lngMain)
    lngMainMenu = GetSubMenu(lngMenuHandle, lngmenu)
    lngMenuItem = GetMenuItemID(lngMainMenu, lngsubmenu)
    
    Call SendMessage(lngMain, WM_COMMAND, lngMenuItem, 0)
    
End Sub



Public Function GetMailCount() As Long
Attribute GetMailCount.VB_Description = "Returns the number of mails in box"

    Dim lngMailWindow As Long, lngVisible As Long
         
    Do
        lngMailWindow = FindMailListWindow
        lngVisible = IsWindowVisible(lngMailWindow)
    Loop Until lngVisible <> 0
    
    GetMailCount = SendMessage(lngMailWindow, LB_GETCOUNT, 0, 0)

End Function

Public Function FindToolIcon(intIcon As TOOLICON) As Long
Attribute FindToolIcon.VB_Description = "Returns the handle to one of the aol toolbar icons"

    Dim lngTool As Long, lngIcon As Long, intLoop As Integer
    
    lngTool = FindToolbar
    
    If intIcon = READ_MAIL Then
        FindToolIcon = FindWindowEx(lngTool, 0, "_AOL_Icon", vbNullString)
        Exit Function
    End If
    
    lngIcon = FindWindowEx(lngTool, 0, "_AOL_Icon", vbNullString)
    
    For intLoop = 1 To intIcon
        lngIcon = FindWindowEx(lngTool, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    
    FindToolIcon = lngIcon
    
End Function



Public Sub KeepMailNew(lngIndex As Long)

    Dim lngMailList As Long, lngMailwin As Long, lngIcon As Long, intLoop As Integer
    
    lngMailList = FindMailListWindow
    Call SendMessage(lngMailList, LB_SETCURSEL, lngIndex, 0)

    Do
        lngMailwin = SortEnumWindows(FindFirstWindow, , "OnlineMailbox", False)
    Loop Until lngMailwin <> 0
    DoEvents
    
    lngIcon = FindWindowEx(lngMailwin, 0, "_AOL_Icon", vbNullString)
    For intLoop = 0 To 1
        lngIcon = FindWindowEx(lngMailwin, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    DoEvents
    
    Call PushButton(lngIcon)

End Sub

Public Function GetMailSubject(lngIndex As Long) As String
Attribute GetMailSubject.VB_Description = "Returns the mail subject "

    Dim lngMailWindow As Long, lngLength As Long, strSubject As String

    lngMailWindow = FindMailListWindow

    lngLength = SendMessage(lngMailWindow, LB_GETTEXTLEN, lngIndex, 0)
    strSubject = String(lngLength, " ")
    Call SendMessageByString(lngMailWindow, LB_GETTEXT, lngIndex, strSubject)
    
    GetMailSubject = FilterMailSubject(strSubject)

End Function

Public Function GetMailSender(lngIndex As Long) As String
Attribute GetMailSender.VB_Description = "Returns the mail sender"

    Dim lngMailWindow As Long, lngLength As Long, strSender As String
    Dim lngChar As Long
    
    lngMailWindow = FindMailListWindow
    
    lngLength = SendMessage(lngMailWindow, LB_GETTEXTLEN, lngIndex, 0)
    strSender = String(lngLength, " ")
    Call SendMessageByString(lngMailWindow, LB_GETTEXT, lngIndex, strSender)
     
    lngChar = InStr(strSender, Chr(9))

    strSender = Mid(strSender, (lngChar + 1), Len(strSender))
    lngChar = InStr(strSender, Chr(9))
    
    GetMailSender = Mid(strSender, 1, (lngChar - 1))

End Function

Public Sub SaveList(lstName As ListBox, strPath As String)
Attribute SaveList.VB_Description = "Saves a listbox to a .txt file"

    Dim intLoop As Integer

    Open strPath For Output As #1
        For intLoop = 0 To lstName.ListCount
            Print #1, lstName.List(intLoop)
            DoEvents
        Next intLoop
        DoEvents
    Close #1
    
End Sub

Public Sub SaveCombo(cmbName As ComboBox, strPath As String)
Attribute SaveCombo.VB_Description = "Saves a combobox to a .txt file"

    Dim intLoop As Integer
    
    Open strPath For Output As #1
        For intLoop = 0 To cmbName.ListCount
            Print #1, cmbName.List(intLoop)
            DoEvents
        Next intLoop
        DoEvents
    Close #1
    
End Sub

Public Sub FileToList(lstName As ListBox, strPath As String)
Attribute FileToList.VB_Description = "Adds a file to a listbox"

    Dim intLoop As Integer, strTemp As String

    Open strPath For Input As #1
    While Not EOF(1)
        Input #1, strTemp
        lstName.AddItem (strTemp)
    Wend
    Close #1

End Sub

Public Sub FileToCombo(cmbName As ComboBox, strPath As String)
Attribute FileToCombo.VB_Description = "Adds a file to a combobox"

    Dim intLoop As Integer, strTemp As String

    Open strPath For Input As #1
    While Not EOF(1)
        Input #1, strTemp
        cmbName.AddItem (strTemp)
    Wend
    Close #1
    
    cmbName.Text = cmbName.List(0)

End Sub

Public Sub GoKeyword(strkeyword As String)

    Dim lngTool As Long, lngCombo As Long, lngEdit As Long

    lngTool = FindToolbar
    lngCombo = FindWindowEx(lngTool, 0, "_AOL_Combobox", vbNullString)
    lngEdit = FindWindowEx(lngCombo, 0, "Edit", vbNullString)
    DoEvents
    
    Call SendMessageByString(lngEdit, WM_SETTEXT, 0, strkeyword)
    Call SendMessage(lngEdit, WM_CHAR, VK_SPACE, 0)
    Call SendMessage(lngEdit, WM_CHAR, VK_RETURN, 0)
    
End Sub



Public Sub KeepTopMost(frmName As Form, blnTop As Boolean)

    If blnTop = True Then
        Call SetWindowPos(frmName.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    ElseIf blnTop = False Then
        Call SetWindowPos(frmName.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
    
End Sub

Public Sub Drag(frmName As Form)
Attribute Drag.VB_Description = "Drags a window without a caption"

    Call ReleaseCapture
    Call SendMessage(frmName.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub




Public Sub SaveTextField(txtName As TextBox, strPath As String)
Attribute SaveTextField.VB_Description = "Saves a textbox to a .txt file"

    Dim intLoop As Integer
    
    Open strPath For Output As #1
        Print #1, txtName.Text
        DoEvents
    Close #1

End Sub

Public Sub FileToText(txtName As TextBox, strPath As String)
Attribute FileToText.VB_Description = "Adds a file to a textbox"

    Dim intLoop As Integer, strTemp As String, intCount As Integer

    Open strPath For Input As #1
    While Not EOF(1)
        Input #1, strTemp
        intCount = (intCount + 1)
        txtName.SelStart = Len(txtName.Text)
        txtName.SelText = strTemp & vbCrLf
    Wend
    Close #1
    
    For intLoop = 0 To intCount
        Call SendMessage(txtName.hWnd, WM_VSCROLL, 0, 0)
    Next
    
End Sub

Public Sub ReadMail(lngIndex As Long)
Attribute ReadMail.VB_Description = "Opens a mail item"

    Dim lngMailList As Long
    
    lngMailList = FindMailListWindow
    DoEvents
    
    Call SendMessage(lngMailList, LB_SETCURSEL, lngIndex, 0)
    DoEvents
    
    Call PostMessage(lngMailList, WM_KEYDOWN, VK_RETURN, 0)
    Call PostMessage(lngMailList, WM_KEYUP, VK_RETURN, 0)
    DoEvents
    
End Sub

Public Function FindOpenMailWindow() As Long
Attribute FindOpenMailWindow.VB_Description = "Returns the hadle of an open mail window"

    Dim lngMain As Long, lngChild As Long, lngRich As Long, lngView As Long
    
    lngMain = FindFirstWindow
    lngChild = FindWindowEx(lngMain, 0, "AOL Child", vbNullString)
    lngView = FindWindowEx(lngChild, 0, "_AOL_View", vbNullString)
    
    If lngView = 0 Then
        Do
            lngChild = FindWindowEx(lngMain, lngChild, "AOL Child", vbNullString)
            lngView = FindWindowEx(lngChild, 0, "_AOL_View", vbNullString)
        Loop Until lngView <> 0
    End If

    FindOpenMailWindow = lngChild
    
End Function



Public Sub SendImReply(strMessage As String)
Attribute SendImReply.VB_Description = "Sends a reply to an IM"

    Dim lngImWin As Long, lngRich As Long
    Dim lngIcon As Long, intLoop As Integer

    If strMessage = "" Then
        Exit Sub
    End If
    
    Do
        lngImWin = SortEnumWindows(FindFirstWindow, , "Instant Message", False)
    Loop Until lngImWin <> 0
    
    lngRich = FindWindowEx(lngImWin, 0, "RICHCNTL", vbNullString)
    lngRich = FindWindowEx(lngImWin, lngRich, "RICHCNTL", vbNullString)
    
    Call ApplyText(lngRich, strMessage)
    
    DoEvents
    
    lngIcon = FindWindowEx(lngImWin, 0, "_AOL_Icon", vbNullString)
    
    For intLoop = 1 To 8
        lngIcon = FindWindowEx(lngImWin, lngIcon, "_AOL_Icon", vbNullString)
    Next intLoop
    
    Call PushButton(lngIcon)
    
End Sub

Public Function FindInList(lnghWnd As Long, ByVal strSearch As String, Optional lngStartPos As Long = -1, Optional blnExactMatch As Boolean) As Long
Attribute FindInList.VB_Description = "Returns the index of an item in a listbox"
    
    Dim lngSend As Long
    
    lngSend = IIf(blnExactMatch, LB_FINDSTRINGEXACT, LB_FINDSTRING)
    FindInList = SendMessageString(lnghWnd, lngSend, lngStartPos, strSearch)

End Function

Public Function FindInCombo(lnghWnd As Long, ByVal strSearch As String, Optional lngStartPos As Long = -1, Optional blnExactMatch As Boolean) As Long
Attribute FindInCombo.VB_Description = "Returns the index of the item in a combobox"
    
    Dim lngSend As Long
    
    lngSend = IIf(blnExactMatch, CB_FINDSTRINGEXACT, CB_FINDSTRING)
    FindInCombo = SendMessageString(lnghWnd, lngSend, lngStartPos, strSearch)

End Function

Public Sub WindowView(lnghWnd As Long, blnShow As Boolean)
Attribute WindowView.VB_Description = "Show or hide a window"

    If blnShow = True Then
        Call ShowWindow(lnghWnd, SW_SHOW)
    ElseIf blnShow = False Then
        Call ShowWindow(lnghWnd, SW_HIDE)
    End If
    
End Sub

Public Function NumberInRoom() As Long

    Dim lngChat As Long, lngList As Long
    
    lngChat = FindAolChat
    lngList = FindWindowEx(lngChat, 0, "_AOL_ListBox", vbNullString)
    DoEvents
    
    NumberInRoom = SendMessage(lngList, LB_GETCOUNT, 0, 0)

End Function

Public Sub GoToRoom(intRoomPrefix As ROOM, strRoomName As String)
    
    Dim strPrefix As String
    
    If intRoomPrefix = 0 Then
        strPrefix = "aol://2719:21-2-"
    ElseIf intRoomPrefix = 1 Then
        strPrefix = "aol://2719:2-2-"
    End If
    
    Call GoKeyword(strPrefix & strRoomName)

End Sub

Public Sub TimePauseEx(lngMilliSeconds As Long)
Attribute TimePauseEx.VB_Description = "Pause for x amount of milliseconds"

    Dim lngNow As Long
    
    lngMilliSeconds = (GetTickCount + lngMilliSeconds)
    Do Until lngNow >= lngMilliSeconds
        lngNow = GetTickCount
    Loop

End Sub

Public Function CleanWord(strText As String, strFilter As String, Optional strReplaceWith As String = "") As String
Attribute CleanWord.VB_Description = "Replaces a word in a string with a given word"

    Dim intPos As Integer, strTemp As String
    
    Do
        intPos = InStr(strText, strFilter)
        If intPos <> 0 Then
            strTemp = Mid(strText, 1, (intPos - 1))
            strTemp = strTemp & strReplaceWith & Mid(strText, (intPos + (Len(strFilter))), Len(strText))
            DoEvents
            strText = strTemp
        End If
    Loop Until intPos = 0
    
    CleanWord = strText
    
End Function

Public Sub WindowToEllipse(lnghWnd As Long, ByVal lngX1 As Long, ByVal lngY1 As Long, ByVal lngX2 As Long, ByVal lngY2 As Long)
Attribute WindowToEllipse.VB_Description = "Changes the window to an elliptical shape"

    Call SetWindowRgn(lnghWnd, CreateEllipticRgn(lngX1, lngY1, lngX2, lngY2), True)
    DoEvents
    
End Sub

Public Sub WindowToRect(lnghWnd As Long, ByVal lngX1 As Long, ByVal lngY1 As Long, ByVal lngX2 As Long, ByVal lngY2 As Long)
Attribute WindowToRect.VB_Description = "Changes the window to a rectangular shape"

    Call SetWindowRgn(lnghWnd, CreateRectRgn(lngX1, lngY1, lngX2, lngY2), True)
    DoEvents

End Sub

Public Sub ToggleX(frmName As Form, blnOnOff As Boolean)
Attribute ToggleX.VB_Description = "Enables/Disables the X button on a window"

    Dim lngSysMenu As Long, lngPosition As Long
    
    If blnOnOff = True Then
        lngSysMenu = GetSystemMenu(frmName.hWnd, False)
        lngPosition = GetMenuItemCount(lngSysMenu)
        Call RemoveMenu(lngSysMenu, lngPosition - 1, MF_REMOVE Or MF_BYPOSITION)
        Call RemoveMenu(lngSysMenu, lngPosition - 2, MF_REMOVE Or MF_BYPOSITION)
        
        Call DrawMenuBar(frmName.hWnd)
        DoEvents
        
    ElseIf blnOnOff = False Then
        Call GetSystemMenu(frmName.hWnd, True)
        Call DrawMenuBar(frmName.hWnd)
        DoEvents
    End If

End Sub

Public Sub TextToClipboard(strText As String)
Attribute TextToClipboard.VB_Description = "Sets text to clipboard"

    Clipboard.Clear
    Clipboard.SetText strText, vbCFText

End Sub

Public Function TextFromClipboard() As String
Attribute TextFromClipboard.VB_Description = "Returns text from clipboard"

    TextFromClipboard = Clipboard.GetText(vbCFText)
    
End Function



Public Function GetDrive(intDrive As DRIVETYPE) As String
Attribute GetDrive.VB_Description = "Returns the drive type of a specified drive"

    Dim strDrive As String, lngType As String
    
    strDrive = Chr(intDrive + 65) & ":\"
    
    lngType = GetDriveType(strDrive)

    Select Case lngType
        Case Is = 2
            GetDrive = "Floppy"
        Case Is = 3
            GetDrive = "Hard Drive"
        Case Is = 4
            GetDrive = "Remote"
        Case Is = 5
            GetDrive = "CD-Rom"
        Case Is = 6
            GetDrive = "RAM Disk"
    End Select

End Function

Public Function DNTConversions(intType As DNTCONV) As String
Attribute DNTConversions.VB_Description = "Converts date and/or time to a specified format"

    Dim strDat As String, strTim As String, dntNow As Date, strType As String
    
    strDat = Date
    strTim = Time
    
    dntNow = DateValue(strDat) + TimeValue(strTim)
    
    Select Case intType
        Case Is = 0
            strType = "General Date"
        Case Is = 1
            strType = "Long Date"
        Case Is = 2
            strType = "Medium Date"
        Case Is = 3
            strType = "Short Date"
        Case Is = 4
            strType = "Long Time"
        Case Is = 5
            strType = "Medium Time"
        Case Is = 6
            strType = "Short Time"
    End Select
    
    DNTConversions = Format(dntNow, strType)

End Function


