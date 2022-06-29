Attribute VB_Name = "DiVe32"
'* * * * * * * * * * * * * * * * * * * *'
'*              DiVe32.Bas             *'
'*               By WoLF               *'
'*          Release 2.10.24.97         *'
'* * * * * * * * * * * * * * * * * * * *'
'This BaS is Soley for VB5 or 32bit progs

'First off, a message to all the "I THINK
'I am the LEETest on AOL" people. Im dont
'try to say that i was never a Lamer or
'that i always new everything about AOL
'Programming, without the Help of others
'i would probably be playing some online
'game right now.
'So if your not doing something and some
'new person needs some help...Remember
'what it felt like when u needed help and
'no one would help.
'Here at DiVe we lend a hand to each other
'and anyone who needs help. And even though
'Some of the Best Proggers are in DiVe, we
'all still have questions sometimes, so
'remember, if we dont help each other...
'WHO WILL...

'First off id like to thank the Maker of
'Master32.bas, and the person who revised
'it to Masta32.bas...This bas was built on
'the Declares of them...
'GetListAll/GetListSpecific provided by
'AfxOle: AfxOle@aol.com

'I have also provide a set of Functions
'for saving and retrieving data with
'CMDialog.ocx, but the a for a single
'form (in this case MMaker) but can be
'Modified to be universal

'Public Variables...Can hold values though
'the entire prog--These are program
'specific, so delete them and add your
'own
'-----------------

Public majornum
Public minornum
Public timesload As String
Public introyn As String
Public scinyn As String
Public scexyn As String
Public soundyn As String
Public ik As String
Public ima As String
Public removedpeeps
Public soundsyn
Public Chat
Public INTR
Public AOL
Public MDI
Public Chat2
Public AFKTime
Public ICO
Public LB
Public EB
Public UL
Public nobut As Integer
Public access As String
Public fState As FormState
Public gFindString As String
Public gFindCase As Integer
Public gFindDirection As Integer
Public gCurPos As Integer
Public gFirstTime As Integer
Public tabclicked As String

'Public Constants, Variables that dont
'change throughout the program, use to
'Make certain things happen
'----------------

Public Const ThisApp = "MDINote"
Public Const ThisKey = "Recent Files"
Public Const WM_CHAR = &H102
Public Const WM_SETTEXT = &HC
Public Const WM_USER = &H400
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_CLEAR = &H303
Public Const WM_DESTROY = &H2
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MOVE = &H3
Public Const WM_SIZE = &H5
Public Const WM_CREATE = &H1
Public Const WM_MDICREATE = &H220

Public Const BM_GETCHECK = &HF0
Public Const BM_GETSTATE = &HF2
Public Const BM_SETCHECK = &HF1
Public Const BM_SETSTATE = &HF3

Public Const LB_GETITEMDATA = &H199
Public Const LB_GETCOUNT = &H18B
Public Const LB_ADDSTRING = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRING = &H18F
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_GETCURSEL = &H188
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SELECTSTRING = &H18C
Public Const LB_SETCOUNT = &H1A7
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185
Public Const LB_INSERTSTRING = &H181

Public Const VK_HOME = &H24
Public Const VK_RIGHT = &H27
Public Const VK_CONTROL = &H11
Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RETURN = &HD
Public Const VK_SPACE = &H20
Public Const VK_TAB = &H9

Public Const hWnd_TOP = 0
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Const GW_CHILD = 5
Public Const GW_hWndFIRST = 0
Public Const GW_hWndLAST = 1
Public Const GW_hWndNEXT = 2
Public Const GW_hWndPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_HIDE = 0
Public Const SW_RESTORE = 9
Public Const SW_SHOW = 5
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const MF_APPEND = &H100&
Public Const MF_DELETE = &H200&
Public Const MF_CHANGE = &H80&
Public Const MF_ENABLED = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_REMOVE = &H1000&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_UNCHECKED = &H0&
Public Const MF_CHECKED = &H8&
Public Const MF_GRAYED = &H1&
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public Const GWW_HINSTANCE = (-6)
Public Const GWW_ID = (-12)
Public Const GWL_STYLE = (-16)

Public Const PROCESS_VM_READ = &H10
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

'Declared Types, these are custom made
'variables to suite certian needs
'--------------
Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wId As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Type rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Type POINTAPI
   X As Long
   y As Long
End Type

Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

'Function and Subs, items used to access
'data from other programs, the Basis of
'API
'-----------------
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef dest As Any, ByRef Source As Any, ByVal nBytes As Long)
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As rect, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As rect) As Long
Declare Function SetRect Lib "user32" (lpRect As rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Declare Function dwGetStringFromLPSTR Lib "dwspy32.dll" (ByVal lpcopy As Long) As String
Declare Sub dwCopyDataBynum Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source&, ByVal dest&, ByVal nCount&)
Declare Sub dwCopyDataByString Lib "dwspy32.dll" Alias "dwCopyData" (ByVal Source As String, ByVal dest As Long, ByVal nCount&)
Declare Function dwGetAddressForObject& Lib "dwspy32.dll" (object As Any)
Declare Function dwXCopyDataBynumFrom& Lib "dwspy32.dll" Alias "dwXCopyDataFrom" (ByVal mybuf As Long, ByVal foreignbuf As Long, ByVal size As Integer, ByVal foreignPID As Long)
Declare Function dwGetWndInstance& Lib "dwspy32.dll" (ByVal hWnd&)
Declare Function RegisterWindowMessage& Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String)
Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long)
Declare Function EnumWindows& Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long)
Declare Function sendmessagebynum& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Declare Function GetClassName& Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function sendmessagebystring Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function CreatePopupMenu Lib "user32" () As Long
Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByVal lpcMenuItemInfo As MENUITEMINFO) As Long
Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Declare Function DestroyMenu Lib "user32" (ByVal hMenu%) As Integer

Sub FileNew()
'This will open a New File

Dim intResponse As Integer
Dim Response As Integer
Const MB_YESNO = 4
Const MB_YESNOCANCEL = 3
Const MB_ICONSTOP = 16
Const IDYES = 6, IDNO = 7, IDCANCEL = 2


' If the file has changed, save it
' put a FState.dirty = true in the
' Text_change section
    If fState.Dirty = True Then
        Response = MsgBox("You have not saved the current macro, Do u want to save it", MB_YESNOCANCEL + MB_ICONSTOP, "Please dont go....")
        If Response = IDNO Then
        GoTo con1
        ElseIf Response = IDCANCEL Then Exit Sub
        End If
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    ' Clear the textbox and update the caption.
con1:
    Mmaker.Text1.Text = "" 'change to your form
    fState.Dirty = False
    Mmaker.Caption = "Macro Maker - Untitled" 'Change to your form
End Sub
Function FileSave() As Integer
'Save File
    Dim strFilename As String

    If Mmaker.Caption = "Macro Maker - Untitled" Then 'Change to your form
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        strFilename = GetFileName(strFilename)
    Else
        ' The form's Caption contains the name of the open file.
        strFilename = Right(Mmaker.Caption, Len(Mmaker.Caption) - 14) 'Change to your form
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strFilename <> "" Then
        SaveFileAs strFilename
        FileSave = True
    Else
        FileSave = False
    End If
End Function



Sub GetRecentFiles()
    ' This procedure demonstrates the use of the GetAllSettings function,
    ' which returns an array of values from the Windows registry. In this
    ' case, the registry contains the files most recently opened.  Use the
    ' SaveSetting statement to write the names of the most recent files.
    ' That statement is used in the WriteRecentFiles procedure.
    Dim i As Integer
    Dim varFiles As Variant ' Varible to store the returned array.
    
    ' Get recent files from the registry using the GetAllSettings statement.
    ' ThisApp and ThisKey are constants defined in this module.
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
    For i = 0 To UBound(varFiles, 1)
    DoEvents

        Mmaker.mnuRecentFile(0).Visible = True
        Mmaker.mnuRecentFile(i + 1).Caption = varFiles(i, 1)
        Mmaker.mnuRecentFile(i + 1).Visible = True
    Next i
End Sub

Function GetFromINI(AppName$, KeyName$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function
Sub WriteRecentFiles(OpenFileName)
    ' This procedure uses the SaveSettings statement to write the names of
    ' recently opened files to the System registry. The SaveSetting
    ' statement requires three parameters. Two of the parameters are
    ' stored as constants and are defined in this module.  The GetAllSettings
    ' function is used in the GetRecentFiles procedure to retrieve the
    ' file names stored in this procedure.
    
    Dim i As Integer
    Dim strFile As String
    Dim strKey As String

    ' Copy RecentFile1 to RecentFile2, and so on.
    For i = 3 To 1 Step -1
    DoEvents
    strKey = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, strKey)
        If strFile <> "" Then
            strKey = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, strKey, strFile
        End If
    Next i
  
    ' Write the open file to first recent file.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub


Sub FileOpenProc()
'Open A file

    Dim intRetVal
    Dim intResponse As Integer
    Dim strOpenFileName As String
    
    ' If the file has changed, save it
    ' put a FState.dirty = true in the
    ' Text_change section
If fState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    On Error Resume Next
    
    Mmaker.CMDialog1.FileName = "" 'Change to your form
    Mmaker.CMDialog1.ShowOpen 'Change to your form
    If Err <> 32755 Then    ' User chose Cancel.
        Debug.Print Err
        strOpenFileName = Mmaker.CMDialog1.FileName 'Change to your form
        ' If the file is larger than 65K, it can't
        ' be opened, so cancel the operation.
        'If FileLen(strOpenFileName) > 65000 Then
            'MsgBox "The file is too large to open."
            'Exit Sub
        'End If
        
        OpenFile (strOpenFileName)
        UpdateFileMenu (strOpenFileName)
    End If
End Sub

Function GetFileName(FileName As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
    On Error Resume Next
    Mmaker.CMDialog1.FileName = FileName
    Mmaker.CMDialog1.ShowSave
    If Err <> 32755 Then    ' User chose Cancel.
        GetFileName = Mmaker.CMDialog1.FileName
    Else
        GetFileName = ""
    End If
End Function

Function OnRecentFilesList(FileName) As Integer
    Dim i         ' Counter variable.

    For i = 1 To 4
    DoEvents
    If Mmaker.mnuRecentFile(i).Caption = FileName Then
            OnRecentFilesList = True
            Exit Function
        End If
    Next i
    OnRecentFilesList = False
End Function

Sub OpenFile(FileName)
'This will open a file
    Dim fIndex As Integer
    
    On Error Resume Next
    ' Open the selected file.
    Open FileName For Input As #1
    If Err Then
        MsgBox "Can't open file: " + FileName
        Exit Sub
    End If
    ' Change the mouse pointer to an hourglass.
    Screen.MousePointer = 11
    
    ' Change the form's caption and display the new text.
    Mmaker.Caption = "Macro Maker - " & UCase(FileName)
    Mmaker.Text1.Text = Input(LOF(1), 1)
    fState.Dirty = False
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
End Sub

Sub SaveFileAs(FileName)
    On Error Resume Next
    Dim strContents As String

    ' Open the file.
    Open FileName For Output As #1
    ' Place the contents of the notepad into a variable.
    strContents = Mmaker.Text1.Text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to a saved file.
    Print #1, strContents
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    ' Set the form's caption.
    If Err Then
        MsgBox Error, 48, App.TITLE
    Else
        Mmaker.Caption = "Macro Maker - " & FileName
        ' Reset the dirty flag.
        fState.Dirty = False
    End If
End Sub

Sub UpdateFileMenu(FileName)
        Dim intRetVal As Integer
        ' Check if the open filename is already in the File menu control array.
        intRetVal = OnRecentFilesList(FileName)
        If Not intRetVal Then
            ' Write open filename to the registry.
            WriteRecentFiles (FileName)
        End If
        ' Update the list of the most recently opened files in the File menu control array.
        GetRecentFiles
End Sub

Public Function FindandAddSN(txt As String, SN As String) As String
'this will find The phrase "{SN}" in a
'text and replace it with the SN u pass in
'Great For Mass IM and IM answer

If txt Like "*" + "{SN}" + "*" Then
For a = 1 To Len(txt)
DoEvents
strchar = Mid(txt, a, 1)
If strchar = "{" Then
b = a + 1
strchar = Mid(txt, b, 1)
If strchar = "S" Then
c = b + 1
strchar = Mid(txt, c, 1)
If strchar = "N" Then
d = c + 1
strchar = Mid(txt, d, 1)
If strchar = "}" Then
firstpart = Mid(txt, 1, a - 1)
lastpart = Mid(txt, a + 4)
fintxt = firstpart + SN + lastpart
FindandAddSN = fintxt
Exit Function
End If
End If
End If
End If
Next
End If
FindandAddSN = txt
End Function
Public Function GetPercent(hwndwin As Long) As String
'Hehe this was used with AOL_Modal to
'get the percent from a upload or download

welcomelength = GetWindowTextLength(hwndwin)
welcometitle$ = Space(welcomelength)
wintext = GetWindowText(hwndwin, welcometitle$, (welcomelength + 1))
finper = Mid(welcometitle, 17)
GetPercent = finper
End Function
Public Function GetListSpecific(ByVal hChatList As Long, ByVal nListIdx As Integer) As String
'this will get names 1 at a time from a
'from a _Aol_listbox

    '/Setup error handling
    On Error GoTo Err_GetListSpecific
    
    Dim hAOLProcess As Long   ' A handle of AOL's process.
    
    Dim lAddrOfItemData As Long ' Memory location of a list items item data.
    Dim sBuffer As String       ' Buffer to place data read by ReadProcessMemory. NOTE: This will only work if a string
                                ' buffer is used...I don't know why though, but would like to so if you have an idea I
                                ' be be happy to hear it.
    Dim lAddrOfName As Long     ' Memory location of the screen name.
    Dim lBytesRead As Long      ' Number of bytes read by ReadProcessMemory.
    
    
    ' Make sure 0 wasn't passed to this function.
    If hChatList Then
        ' Get a valid process handle for AOL that enables you to read(PROCESS_VM_READ) memory in it's process space.
        hAOLProcess = GetAOLProcessHandle(hChatList)
            
        ' Make sure a handle was retrieved
        If hAOLProcess Then
        
            ' Setup the buffer
            sBuffer = String$(4, vbNullChar)
            
            ' Get the item data for the list item
            lAddrOfItemData = SendMessage(hChatList, LB_GETITEMDATA, ByVal CLng(nListIdx%), ByVal 0&)
                            
            ' lAddrOfItemData is actually a pointer to a 7 element array of 4byte values...reading begins from
            ' the address of the 7th element.  Since lAddrOfItemData already equals the address of the first element
            ' you add 4 for each element, so the 7th element's memory address is...
            lAddrOfItemData = lAddrOfItemData + (4 * 6)
            
            ' lAddrOfItemData is now the address of the 7th element of the array, this element contains a 4byte pointer
            ' to a string(the screen name)
            Call ReadProcessMemory(hAOLProcess, lAddrOfItemData, sBuffer, 4, lBytesRead)
                        
            ' The 4 bytes in sBuffer are actually a pointer, this pointer needs to be incremented by 6 so sBuffer needs
            ' to be convertd to long value
            RtlMoveMemory lAddrOfName, ByVal sBuffer, 4
            
            ' Increment the address
            lAddrOfName = lAddrOfName + 6
            
            ' Setup buffer
            sBuffer = String$(16, vbNullChar)
            
            ' lAddrOfName now holds a pointer to a string(screen name), so retrieve the string
            Call ReadProcessMemory(hAOLProcess, lAddrOfName, sBuffer, Len(sBuffer), lBytesRead)
            
            ' That's it, so add the screen name to the array, but be sure to trim off any extra characters.
            GetListSpecific = Left$(sBuffer$, InStr(sBuffer$, vbNullChar) - 1)
            
            ' Close the handle to AOL's process
            Call CloseHandle(hAOLProcess)
        End If
    End If
    
    Exit Function

'/Error handler
Err_GetListSpecific:
    ' Make sure the handle to AOL's process is closed
    Call CloseHandle(hAOLProcess)
    
    Exit Function
    

End Function


Private Function GetAOLProcessHandle(ByVal hWnd As Long) As Long
    
    '/Setup error handling
    On Error Resume Next
    
    Dim m_AOLThreadID As Long   ' A value that uniquely identifies the thread throughout the system.
    Dim m_AOLProcessID As Long  ' A value that uniquely identifies the process throughout the system.
    
    ' Get the process ID for AOL's main thread. Since AOL is not a multithreaded application each window use the same
    ' thread.
    m_AOLThreadID = GetWindowThreadProcessId(hWnd, m_AOLProcessID)
    
    ' Get a valid process handle for AOL that enables you to read(PROCESS_VM_READ) memory in it's process space.
    GetAOLProcessHandle = OpenProcess(PROCESS_VM_READ Or STANDARD_RIGHTS_REQUIRED, False, m_AOLProcessID)
                
End Function
Function GetCaption(hWnd)
hWndLength% = GetWindowTextLength(hWnd)
hWndTitle$ = String$(hWndLength%, 0)
a% = GetWindowText(hWnd, hWndTitle$, (hWndLength% + 1))

GetCaption = hWndTitle$
End Function

Public Function FindChildByNum&(hWnd&, num&)
'This Finds The Child Window By Its Order In The Window
'Like Using GetWindow and GW_HWNDNEXT but faster

Static child, indx, nextwnd

child = GetWindow(hWnd&, GW_CHILD)

indx = 1

nextwnd = GetWindow(child, GW_hWndFIRST)

Do While indx < num&

    nextwnd = GetWindow(nextwnd, GW_hWndNEXT)
    
Loop

FindChildByNum& = nextwnd

Let child = vbNull
Let indx = vbNull
Let nextwnd = vbNull

End Function

Function findchildbytitle(ParenthWnd, hWndTitle) As Integer

ChildhWnd = FindWindowEx(ParenthWnd, 0, vbNullString, hWndTitle)

findchildbytitle = ChildhWnd
End Function

Function FindChildByTitlePartial(parentw, childhand)
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) & "*" Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) & "*" Then GoTo bone
Wend
FindChildByTitlePartial = 0
bone:
RooM% = firs%
FindChildByTitlePartial = RooM%
End Function
Function FindChildByClass(ParenthWnd, hWndClassName) As Integer
ChildhWnd = FindWindowEx(ParenthWnd, 0, hWndClassName, vbNullString)
FindChildByClass = ChildhWnd
End Function
Public Sub AddRoom(the_AOL_list As Long, listbox_to_add_to As ListBox)
'This will add the room to a list
'use the name of the listbox only
'EXAMPLE
'thelist& = FindChildByClass(FindChatRoom, "_AOL_Listbox")
'Call AddRoom(thelist, List1)

chatcount = SendMessage(the_AOL_list, LB_GETCOUNT, 0, 0)
For i = 0 To chatcount - 1
DoEvents
sname = GetListSpecific(the_AOL_list, i)
If sname <> GetUser Then
For X = 0 To listbox_to_add_to.ListCount - 1
DoEvents
If listbox_to_add_to.List(X) = sname Then GoTo gohere
Next
If sname <> "" Then listbox_to_add_to.AddItem (sname)
gohere:
End If
Next

End Sub
Public Function GetRoomAlike()
GetMDI
newroom = FindChildByClass(MDI, "AOL Child")
yesbut = findchildbytitle(newroom, "Yes")
nobut = findchildbytitle(newroom, "No")
writing = FindChildByClass(newroom, "_AOL_Static")
If yesbut <> 0 And nobut <> 0 And writing <> 0 Then
GetRoomAlike = newroom
Exit Function
End If
GetRoomAlike = 0
End Function


Public Function GetListAll(ByVal hChatList As Long, ByRef sUsers() As String) As Long
'This will get a _AOL_listbox or _AOL_Tree
'and add it to a string

    '/Setup error handling
    On Error GoTo Err_GetListAll
    
    Dim hAOLProcess As Long     ' Handle of AOL's process
    
    Dim lAddrOfItemData As Long ' Memory location of a list items item data
    Dim sBuffer As String       ' Buffer to place data read by ReadProcessMemory. NOTE: This will only work if a string
                                ' buffer is used...I don't know why though, but I would like to so if you have an idea I'd
                                ' be be happy to hear it
    Dim lAddrOfName As Long     ' Memory location of the screen name
    Dim lBytesRead As Long      ' Number of bytes read by ReadProcessMemory
    
    Dim lCnt As Long            ' Counter used to loop through the list
    Dim lListCount As Long      ' Number of items in the list
    
    ' Make sure 0 wasn't passed to this function.
    If hChatList Then
        ' Get a valid process handle for AOL that enables you to read(PROCESS_VM_READ) memory in it's process space.
        hAOLProcess = GetAOLProcessHandle(hChatList)
        
        ' Make sure a handle was retrieved
        If hAOLProcess Then
            
            ' Get the number of items in the list.
            lListCount = SendMessage(hChatList, LB_GETCOUNT, ByVal 0&, ByVal 0&) - 1
            
            ' Re dimension the array to hold all the items in the list.
            ReDim sUsers(lListCount)
            
            ' For each item in the list...
            For lCnt& = 0 To lListCount
            DoEvents
    ' Setup the buffer
                sBuffer = String$(4, vbNullChar)
                
                ' Get the item data for the list item
                lAddrOfItemData = SendMessage(hChatList, LB_GETITEMDATA, ByVal lCnt&, ByVal 0&)
                                
                ' lAddrOfItemData is actually a pointer to a 7 element array of 4byte values...reading begins from
                ' the address of the 7th element.  Since lAddrOfItemData already equals the address of the first element
                ' you add 4 for each element, so the 7th element's memory address is...
                lAddrOfItemData = lAddrOfItemData + (4 * 6)
                
                ' lAddrOfItemData is now the address of the 7th element of the array, this element contains a 4byte pointer
                ' to a string(the screen name)
                Call ReadProcessMemory(hAOLProcess, lAddrOfItemData, sBuffer, 4, lBytesRead)
                            
                ' The 4 bytes in sBuffer are actually a pointer, this pointer needs to be incremented by 6 so sBuffer needs
                ' to be convertd to long value
                RtlMoveMemory lAddrOfName, ByVal sBuffer, 4
                
                ' Increment the address
                lAddrOfName = lAddrOfName + 6
                
                ' Setup buffer
                sBuffer = String$(16, vbNullChar)
                
                ' lAddrOfName now holds a pointer to a string(screen name), so retrieve the string
                Call ReadProcessMemory(hAOLProcess, lAddrOfName, sBuffer, Len(sBuffer), lBytesRead)
                
                ' That's it, so add the screen name to the array, but be sure to trim off any extra characters.
                sUsers(lCnt) = Left$(sBuffer$, InStr(sBuffer$, vbNullChar) - 1)
            Next
            
            ' Return the number of items add to the array
            GetListAll = UBound(sUsers)
            
            ' Close the handle to AOL's process
            Call CloseHandle(hAOLProcess)
        End If
    End If
    
    Exit Function

'/Error handler
Err_GetListAll:
    ' Make sure the handle to AOL's process is closed
    Call CloseHandle(hAOLProcess)
    
    Exit Function
    
End Function


Function AddListToString(thelist As ListBox)
'This will take a list box and add the
'entrys to a string with a comma to
'separate them. Usefull for MMers

For DoList = 0 To thelist.ListCount - 1
DoEvents
AddListToString = AddListToString & thelist.List(DoList) & ", "
Next DoList
AddListToString = Mid(AddListToString, 1, Len(AddListToString) - 2)
End Function


Public Sub waitforok()
Do
DoEvents
okw = FindWindow("#32770", "America Online")
okb = findchildbytitle(okw, "OK")
DoEvents
Loop Until okb <> 0
Do
okw = FindWindow("#32770", "America Online")
    okb = findchildbytitle(okw, "OK")
    okd = sendmessagebynum(okb, WM_LBUTTONDOWN, 0, 0&)
    oku = sendmessagebynum(okb, WM_LBUTTONUP, 0, 0&)
DoEvents
Loop Until okw = 0

End Sub
Sub AddStringToList(theitems, thelist As ListBox)
'This will take a string with multiple
'variables separated by commas and add
'them to a list

If Not Mid(theitems, Len(theitems), 1) = "," Then
theitems = theitems & ","
End If

For DoList = 1 To Len(theitems)
DoEvents
thechars$ = thechars$ & Mid(theitems, DoList, 1)

If Mid(theitems, DoList, 1) = "," Then
thelist.AddItem Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
If Mid(theitems, DoList + 1, 1) = " " Then
DoList = DoList + 1
End If
End If
Next DoList

End Sub


Function ClickList(hWnd)
'This will click on the selected item
'in a _AOL_Listbox or _AOL_Tree

Call sendmessagebynum(hWnd, &H203, 0, 0&)
End Function


Function countmail()
'This will return the number of mails in
'a open mailbox
GetMDI
themail = FindChildByClass(MDI, "AOL Child")
thetree = FindChildByClass(themail, "_AOL_Tree")
countmail = SendMessage(themail, LB_GETCOUNT, 0, 0)
End Function



Sub OpenMailSpecific(which)
'If you pass in the number 1, it will open
'your newmail, number2 your old mail, any
'other number your mail uve read

If which = 1 Then
Call RunMenuByString("Read &New Mail")
End If

If which = 2 Then
Call RunMenuByString("Check Mail You've &Read")
End If

If Not which = 1 Or Not which = 2 Then
Call RunMenuByString("Check Mail You've &Sent")
End If

End Sub


Sub RespondIM(message)
'this will respone to a open IM
GetMDI
im% = findchildbytitle(MDI, ">Instant Message From:")
If im% Then GoTo Z
im% = findchildbytitle(MDI, "  Instant Message From:")
If im% Then GoTo Z
Exit Sub
Z:
E = FindChildByClass(im%, "RICHCNTL")
For i = 1 To 9
DoEvents
E = GetWindow(E, 2)
Next
E2 = GetWindow(E, 2) 'Send Text
E = GetWindow(E2, 2) 'Send Button
Call SetText(E2, message)
ClickIcon (E)
End Sub

Sub RunMenuByString(stringer As String)
'This will run the string from AOLs menu
'that u enter
GetAOL
Call RMBS(AOL, stringer)
End Sub




Function Encrypt_Decrypt(Text, types)
'to encrypt, example:
'encrypted$ = Encrypt_Decrypt("messagetoencrypt", 0)
'to decrypt, example:
'decrypted$ = Encrypt_Decrypt("decryptedmessage", 1)
'* First Paramete is the Message
'* Second Parameter is 0 for encrypt
'  or 1 for decrypt

For God = 1 To Len(Text)
DoEvents
If types = 0 Then
Current$ = Asc(Mid(Text, God, 1)) - 1
Else
Current$ = Asc(Mid(Text, God, 1)) + 1
End If
Process$ = Process$ & Chr(Current$)
Next God

Encrypt_Decrypt = Process$
End Function






Function DescrambleText(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Descrambles the text
For scrambling = 1 To Len(thetext)
DoEvents
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo city
lastchar$ = Mid(chars$, 2, 1)
'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 3, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
DoEvents
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffed

'adds the scrambled text to the full scrambled element
city:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniff

sniffed:
scrambled$ = scrambled$ & lastchar$ & backchar$ & firstchar$ & " "

'clears character and reversed buffers
sniff:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
DescrambleText = scrambled$

End Function



Function GetLineCount(Text)
'This will get the number of lines in
'a Textbox or string
theview$ = Text


For FindChar = 1 To Len(theview$)
DoEvents
thechar$ = Mid(theview$, FindChar, 1)

If thechar$ = Chr(13) Then
numline = numline + 1
End If

Next FindChar

If Mid(Text, Len(Text), 1) = Chr(13) Then
GetLineCount = numline
Else
GetLineCount = numline + 1
End If
End Function

Function IntegerToString(tochange As Integer) As String
'This will convert a integer to string
'for sending to chat room

IntegerToString = Str$(tochange)
End Function







Function RandomNumber(finished)
'This will get a random number
Randomize
RandomNumber = Int((Val(finished) * Rnd) + 1)
End Function

Function Scrambletext(thetext)
'sees if there's a space in the text to be scrambled,
'if found space, continues, if not, adds it
findlastspace = Mid(thetext, Len(thetext), 1)

If Not findlastspace = " " Then
thetext = thetext & " "
Else
thetext = thetext
End If

'Scrambles the text
For scrambling = 1 To Len(thetext)
DoEvents
thechar$ = Mid(thetext, scrambling, 1)
Char$ = Char$ & thechar$

If thechar$ = " " Then
'takes out " " space from the text left of the space
chars$ = Mid(Char$, 1, Len(Char$) - 1)
'gets first character
firstchar$ = Mid(chars$, 1, 1)
'gets last character (if not, makes first character only)
On Error GoTo cityz
lastchar$ = Mid(chars$, Len(chars$), 1)

'finds what is inbetween the last and first character
midchar$ = Mid(chars$, 2, Len(chars$) - 2)
'reverses the text found in between the last and first
'character
For SpeedBack = Len(midchar$) To 1 Step -1
DoEvents
backchar$ = backchar$ & Mid$(midchar$, SpeedBack, 1)
Next SpeedBack
GoTo sniffe

'adds the scrambled text to the full scrambled element
cityz:
scrambled$ = scrambled$ & firstchar$ & " "
GoTo sniffs

sniffe:
scrambled$ = scrambled$ & lastchar$ & firstchar$ & backchar$ & " "

'clears character and reversed buffers
sniffs:
Char$ = ""
backchar$ = ""
End If

Next scrambling
'Makes function return value the scrambled text
Scrambletext = scrambled$

Exit Function
End Function


Function ReplaceText(Text, charfind, charchange)
If InStr(Text, charfind) = 0 Then
ReplaceText = Text
Exit Function
End If

For Replace = 1 To Len(Text)
DoEvents
thechar$ = Mid(Text, Replace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = charfind Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1) + charchange
End If
Next Replace

ReplaceText = thechars$

End Function


Sub SetBackPre()
GetMDI
Call RunMenuByString("Preferences")
Do: DoEvents
prefer% = findchildbytitle(MDI, "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_hWndNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

timeout (0.2)
ClickIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Closewindows% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And Closewindows% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(Closewindows%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 1, 0)

Click_Button (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Function StayOnline()
hWndz% = FindWindow("_AOL_Palette", "America Online")
ChildhWnd% = findchildbytitle(hWndz%, "OK")
Click_Button (ChildhWnd%)
End Function

Function StringToInteger(tochange As String) As Integer


StringToInteger = tochange
On Error GoTo err1234
Exit Function
err1234:
StringToInteger = ""
End Function
Function TrimCharacter(thetext, chars)
TrimCharacter = ReplaceText(thetext, chars, "")

End Function

Function TrimReturns(thetext)
takechr13 = ReplaceText(thetext, Chr$(13), "")
takechr10 = ReplaceText(takechr13, Chr$(10), "")
TrimReturns = takechr10
End Function

Function TrimSpaces(Text)
If InStr(Text, " ") = 0 Then
TrimSpaces = Text
Exit Function
End If

For trimspace = 1 To Len(Text)
DoEvents
thechar$ = Mid(Text, trimspace, 1)
thechars$ = thechars$ & thechar$

If thechar$ = " " Then
thechars$ = Mid(thechars$, 1, Len(thechars$) - 1)
End If
Next trimspace

TrimSpaces = thechars$
End Function





Function UntilWindowClass(parentw, childhand)
GoBack:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone

While firs%
firss% = GetWindow(parentw, 5)
If UCase(Mid(GetClass(firss%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(Mid(GetClass(firs%), 1, Len(childhand))) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBack
FindClassLike = 0

bone:
RooM% = firs%
UntilWindowClass = RooM%
End Function

Function FindFwdWin(dosloop)
'This will find the Fwd mail win
GetAOL
firs% = GetWindow(FindChildByClass(AOL, "MDIClient"), 5)
forw% = findchildbytitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOL, "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOL, "MDIClient"), 5)
forw% = findchildbytitle(firss%, "Forward")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = findchildbytitle(firs%, "Forward")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindFwdWin = firs%

Exit Function
begis:
FindFwdWin = firss%
End Function


Function FindSendWin(dosloop)
GetAOL
'This will find the send window after
'Pressing Fwd on a Mail
firs% = GetWindow(FindChildByClass(AOL, "MDIClient"), 5)
forw% = findchildbytitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
firs% = GetWindow(FindChildByClass(AOL, "MDIClient"), GW_CHILD)

Do: DoEvents
firss% = GetWindow(FindChildByClass(AOL, "MDIClient"), 5)
forw% = findchildbytitle(firss%, "Send Now")
If forw% <> 0 Then GoTo begis
firs% = GetWindow(firs%, 2)
forw% = findchildbytitle(firs%, "Send Now")
If forw% <> 0 Then GoTo bone
If dosloop = 1 Then Exit Do
Loop
Exit Function
bone:
FindSendWin = firs%

Exit Function
begis:
FindSendWin = firss%
End Function

Function UntilWindowTitle(parentw, childhand)
GoBac:
DoEvents
firs% = GetWindow(parentw, 5)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(parentw, GW_CHILD)

While firs%
firss% = GetWindow(parentw, 5)
If UCase(GetCaption(firss%)) Like UCase(childhand) Then GoTo bone
firs% = GetWindow(firs%, 2)
If UCase(GetCaption(firs%)) Like UCase(childhand) Then GoTo bone
Wend
GoTo GoBac
FindWindowLike = 0

bone:
RooM% = firs%
UntilWindowTitle = RooM%

End Function


Function KTEncrypt(ByVal password, ByVal strng, force%)
'Example:
'temp = KTEncrypt ("Paszwerd", text1.text, 0)
'text1.text = temp


  'Set error capture routine
  On Local Error GoTo ErrorHandler

  
  'Is there Password??
  If Len(password) = 0 Then Error 31100
  
  'Is password too long
  If Len(password) > 255 Then Error 31100

  'Is there a strng$ to work with?
  If Len(strng) = 0 Then Error 31100

  
  'Check if file is encrypted and not forcing
  If force% = 0 Then
    
    'Check for encryption ID tag
    chk$ = Left$(strng, 4) + Right$(strng, 4)
    
    If chk$ = Chr$(1) + "KT" + Chr$(1) + Chr$(1) + "KT" + Chr$(1) Then
      
      'Remove ID tag
      strng = Mid$(strng, 5, Len(strng) - 8)
      
      'String was encrypted so filter out CHR$(1) flags
      look = 1
      Do
        look = InStr(look, strng, Chr$(1))
        If look = 0 Then
          Exit Do
        Else
          Addin$ = Chr$(Asc(Mid$(strng, look + 1)) - 1)
          strng = Left$(strng, look - 1) + Addin$ + Mid$(strng, look + 2)
        End If
        look = look + 1
      Loop
      
      'Since it is encrypted we want to decrypt it
      EncryptFlag% = False
    
    Else
      'Tag not found so flag to encrypt string
      EncryptFlag% = True
    End If
  Else
    'force% flag set, ecrypt string regardless of tag
    EncryptFlag% = True
  End If
    


  'Set up variables
  PassUp = 1
  PassMax = Len(password)
  
  
  'Tack on leading characters to prevent repetative recognition
  password = Chr$(Asc(Left$(password, 1)) Xor PassMax) + password
  password = Chr$(Asc(Mid$(password, 1, 1)) Xor Asc(Mid$(password, 2, 1))) + password
  password = password + Chr$(Asc(Right$(password, 1)) Xor PassMax)
  password = password + Chr$(Asc(Right$(password, 2)) Xor Asc(Right$(password, 1)))
  
  
  'If Encrypting add password check tag now so it is encrypted with string
  If EncryptFlag% = True Then
    strng = Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") + strng
  End If
  
  'Loop until scanned though the whole string
  For Looper = 1 To Len(strng)
DoEvents
    'Alter character code
    tochange = Asc(Mid$(strng, Looper, 1)) Xor Asc(Mid$(password, PassUp, 1))

    'Insert altered character code
    Mid$(strng, Looper, 1) = Chr$(tochange)
    
    'Scroll through password string one character at a time
    PassUp = PassUp + 1
    If PassUp > PassMax + 4 Then PassUp = 1
      
  Next Looper

  'If encrypting we need to filter out all bad character codes (0, 10, 13, 26)
  If EncryptFlag% = True Then
    'First get rid of all CHR$(1) since that is what we use for our flag
    look = 1
    Do
    DoEvents

      look = InStr(look, strng, Chr$(1))
      If look > 0 Then
        strng = Left$(strng, look - 1) + Chr$(1) + Chr$(2) + Mid$(strng, look + 1)
        look = look + 1
      End If
    Loop While look > 0

    'Check for CHR$(0)
    Do
    DoEvents

      look = InStr(strng, Chr$(0))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(1) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(10)
    Do
DoEvents
      look = InStr(strng, Chr$(10))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(11) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(13)
    Do
DoEvents
      look = InStr(strng, Chr$(13))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(14) + Mid$(strng, look + 1)
    Loop While look > 0

    'Check for CHR$(26)
    Do
DoEvents
      look = InStr(strng, Chr$(26))
      If look > 0 Then strng = Left$(strng, look - 1) + Chr$(1) + Chr$(27) + Mid$(strng, look + 1)
    Loop While look > 0

    'Tack on encryted tag
    strng = Chr$(1) + "KT" + Chr$(1) + strng + Chr$(1) + "KT" + Chr$(1)

  Else
    
    'We decrypted so ensure password used was the correct one
    If Left$(strng, 9) <> Left$(password, 3) + Format$(Asc(Right$(password, 1)), "000") + Format$(Len(password), "000") Then
      'Password bad cause error
      Error 31100
    Else
      'Password good, remove password check tag
      strng = Mid$(strng, 10)
    End If

  End If


  'Set function equal to modified string
  KTEncrypt = strng
  

  'Were out of here
  Exit Function


ErrorHandler:
  
  'We had an error!  Were out of here
  Exit Function

End Function

Public Sub CenterForm(frmForm As Form)
   With frmForm
      .Left = (Screen.Width - .Width) / 2
      .Top = (Screen.Height - .Height) / 2
   End With
End Sub



Public Function GetChildCount(ByVal hWnd As Long) As Long
'This gets the number of open childs
Dim hChild As Long

Dim i As Integer
   
If hWnd = 0 Then
GoTo Return_False
End If

hChild = GetWindow(hWnd, GW_CHILD)
   

While hChild
hChild = GetWindow(hChild, GW_hWndNEXT)
i = i + 1
Wend

GetChildCount = i
   
Exit Function
Return_False:
GetChildCount = 0
Exit Function
End Function

Public Sub Click_Button(but)
'This will click on a "_AOL_Button"

Call SendMessage(but, WM_KEYDOWN, VK_SPACE, 0)
Call SendMessage(but, WM_KEYUP, VK_SPACE, 0)
End Sub
Function ChatRoom()
'©MKB, do NOT MM this.
'This was written BY MKB FOR MKB and is not intended
'to be distributed to those not associated with DiVe

GetMDI
aoc% = GetWindow(MDI, GW_CHILD)
aoc% = GetWindow(aoc%, GW_hWndFIRST)
aon% = aoc%
Do
    ce% = FindChildByClass(aon%, "_AOL_Edit")
    cv% = FindChildByClass(aon%, "_AOL_View")
    cl% = FindChildByClass(aon%, "_AOL_Listbox")
    If ce% <> 0 And cv% <> 0 And cl% <> 0 Then GoTo 1
    aon% = GetWindow(aon%, GW_hWndNEXT)
Loop Until aon% = aoc%
ChatRoom = 0
Exit Function
1:
ChatRoom = aon%
End Function
Function GetUser()
'This Will get the USER
On Error Resume Next
AOL = FindWindow("AOL Frame25", "America  Online")
MDI = FindChildByClass(AOL, "MDIClient")
welcome = FindChildByTitlePartial(MDI, "Welcome, ")
welcomelength = GetWindowTextLength(welcome)
welcometitle$ = String$(200, 0)
a = GetWindowText(welcome, welcometitle$, (welcomelength + 1))
User = Mid$(welcometitle$, 10, (InStr(welcometitle$, "!") - 10))
GetUser = User
End Function

Sub IMsOff()
'This Turns IM's off, feel free to remove
'the sendchat
Call SendInstantMessage("$IM_OFF", "Turn off!")
SendChat ("•·· ·´¯`·( Total Eclipse, IM's are now OFF!")
End Sub

Sub IMsOn()
'This Turns IM's on, feel free to remove
'the sendchat
Call SendInstantMessage("$IM_ON", "Turn on!")
SendChat ("•·· ·´¯`·( Total Eclipse, IM's are now ON!")
End Sub


Sub SendChat(txt)
'This will send text to the chat room
RooM% = FindChatRoom()
Call SetText(FindChildByClass(RooM%, "_AOL_Edit"), txt)
DoEvents
Call SendCharNum(FindChildByClass(RooM%, "_AOL_Edit"), 13)
End Sub


Sub CloseWindow(winew)
'This will close a window

closes = SendMessage(winew, WM_CLOSE, 0, 0)
End Sub
Function GetChatText()
'This will get the entire chat window
'see GetLastChatLine for getting just the
'last line
childs% = FindChatRoom()
child = FindChildByClass(childs%, "_AOL_View")
gettrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = sendmessagebystring(child, 13, gettrim + 1, trimspace$)
theview$ = trimspace$
GetChatText = theview$
End Function

Function GetText(child)
'This will get the Text from any window

gettrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = sendmessagebystring(child, 13, gettrim + 1, trimspace$)
GetText = trimspace$
End Function

Sub ClickIcon(icon%)
'This will click on a "_AOL_Icon"

Click% = SendMessage(icon%, WM_LBUTTONDOWN, 0, 0&)
Click% = SendMessage(icon%, WM_LBUTTONUP, 0, 0&)
End Sub

Function SendInstantMessage(PERSON As String, message As String)
GetMDI
RunMenuByString ("Send an Instant Message")
Do: DoEvents
im = findchildbytitle(MDI, "Send Instant Message")
imrich = FindChildByClass(im, "RICHCNTL")
imtext = FindChildByClass(im, "_AOL_Static")
imicon = FindChildByClass(im, "_AOL_Icon")
If im <> 0 And imrich <> 0 And imtext <> 0 And imicon <> 0 Then Exit Do
Loop
imedit = GetWindow(imtext, 2)
For i = 1 To 8
DoEvents
imicon = GetWindow(imicon, 2)
Next
Call SetText(imedit, PERSON)
Call SetText(imrich, message)
imicon = FindChildByClass(im, "_AOL_Icon")
For i = 1 To 9
DoEvents
imicon = GetWindow(imicon, 2)
Next
ClickIcon (imicon)
Do: DoEvents
im = findchildbytitle(MDI, "Send Instant Message")
aolcl = FindWindow("#32770", "America Online")
If aolcl <> 0 Then closer = SendMessage(aolcl, WM_CLOSE, 0, 0): closer2 = SendMessage(im, WM_CLOSE, 0, 0): Exit Do
If im = 0 Then Exit Do
Loop
End Function
Function CheckOnline()
'This will see if AOL is signed ON

AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
welcome = findchildbytitle(MDI, "Welcome, ")
If welcome = 0 Then
MsgBox "Please sign on before using this feature.", 64, "Online"
CheckOnline = 0
Exit Function
End If
CheckOnline = 1
End Function


Sub SendKeyword(Text)
'This will goto to a keyword

Call RunMenuByString("Keyword...")
Do: DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
keyw = findchildbytitle(MDI, "Keyword")
kedit = FindChildByClass(keyw, "_AOL_Edit")
If kedit Then Exit Do
Loop

editsend = sendmessagebystring(kedit, WM_SETTEXT, 0, Text)
pausing = DoEvents()
Sending = SendMessage(kedit, 258, 13, 0)
pausing = DoEvents()
End Sub

Function GetLastChatLine()
'This will get the last line in the chat
'room

getpar = FindChatRoom()
child = FindChildByClass(getpar, "_AOL_View")
gettrim = sendmessagebynum(child, 14, 0&, 0&)
trimspace$ = Space$(gettrim)
getString = sendmessagebystring(child, 13, gettrim + 1, trimspace$)

theview$ = trimspace$


For FindChar = 1 To Len(theview$)
thechar$ = Mid(theview$, FindChar, 1)
thechars$ = thechars$ & thechar$

If thechar$ = Chr(13) Then
thechatext$ = Mid(thechars$, 1, Len(thechars$) - 1)
thechars$ = ""
End If

Next FindChar

lastlen = Val(FindChar) - Len(thechars$)
If thechars = "" Then GoTo bad
lastline = Mid(theview$, lastlen + 1, Len(thechars$) - 1)
If lastline <> "" Then
GetLastChatLine = lastline
Else
bad:
GetLastChatLine = " "
End If
End Function
Sub SendMail2(PERSON, SUBJECT, message)
Call RunMenuByString("Compose Mail")

Do: DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
MailWin = findchildbytitle(MDI, "Compose Mail")
icone = FindChildByClass(MailWin, "_AOL_Icon")
peepz = FindChildByClass(MailWin, "_AOL_Edit")
subjt = findchildbytitle(MailWin, "Subject:")
subjec = GetWindow(subjt, 2)
mess = FindChildByClass(MailWin, "RICHCNTL")
If icone <> 0 And peepz <> 0 And subjec <> 0 And mess <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz, WM_SETTEXT, 0, PERSON)
a = sendmessagebystring(subjec, WM_SETTEXT, 0, SUBJECT)
a = sendmessagebystring(mess, WM_SETTEXT, 0, message)

ClickIcon (icone)
ClickIcon (icone)

End Sub
Sub SendMail(PERSON, SUBJECT, message)
'This will send a EMail

Call RunMenuByString("Compose Mail")

Do: DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
MailWin = findchildbytitle(MDI, "Compose Mail")
icone = FindChildByClass(MailWin, "_AOL_Icon")
peepz = FindChildByClass(MailWin, "_AOL_Edit")
subjt = findchildbytitle(MailWin, "Subject:")
subjec = GetWindow(subjt, 2)
mess = FindChildByClass(MailWin, "RICHCNTL")
If icone <> 0 And peepz <> 0 And subjec <> 0 And mess <> 0 Then Exit Do
Loop

a = sendmessagebystring(peepz, WM_SETTEXT, 0, PERSON)
a = sendmessagebystring(subjec, WM_SETTEXT, 0, SUBJECT)
a = sendmessagebystring(mess, WM_SETTEXT, 0, message)

ClickIcon (icone)


Do: DoEvents
ClickIcon (icone)
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
MailWin = findchildbytitle(MDI, "Compose Mail")
erro = findchildbytitle(MDI, "Error")
aolw = FindWindow("#32770", "America Online")
If MailWin = 0 Then Exit Do
If aolw <> 0 Then
a = SendMessage(aolw, WM_CLOSE, 0, 0)
a = SendMessage(MailWin, WM_CLOSE, 0, 0)
Exit Do
End If
If erro <> 0 Then
a = SendMessage(erro, WM_CLOSE, 0, 0)
a = SendMessage(MailWin, WM_CLOSE, 0, 0)
Exit Do
End If
Loop
End Sub



Function GetRoomCount()
'This will get the number of users in a
'room

thechild% = FindChatRoom()
lister% = FindChildByClass(thechild%, "_AOL_Listbox")
getcount = SendMessage(lister%, LB_GETCOUNT, 0, 0)
GetRoomCount = getcount
End Function

Sub SetText(win, txt)
'This will send text to a window

thetext% = sendmessagebystring(win, WM_SETTEXT, 0, txt)
End Sub

Sub SignOff()
'This will sign u off

RunMenuByString ("Sign Off")
End Sub

Function GetAOLVersion()
'This will return your Version

AOL = FindWindow("AOL Frame25", vbNullString)
hMenu = GetMenu(AOL)

submenu = GetSubMenu(hMenu, 0)
subitem = GetMenuItemID(submenu, 8)
MenuString$ = String$(100, " ")

FindString = GetMenuString(submenu, subitem, MenuString$, 100, 1)

If UCase(MenuString$) Like UCase("P&ersonal Filing Cabinet") & "*" Then
GetAOLVersion = 3
Else
GetAOLVersion = 2.5
End If
End Function







Function GetClass(child)
'This will return the Class name of a
'child
Buffer$ = String$(250, 0)
getclas% = GetClassName(child, Buffer$, 250)

GetClass = Buffer$
End Function


Sub NotOnTop(the As Form)
'This will take a form and make it so that
'it does not stay on top of other forms
'U HAVE TO MAKE THE EXE to SEE IT WERK

SetWinOnTop = SetWindowPos(the.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub timeout(interval)
'This will pause for However many seconds
'your decide
Current = Timer
Do While Timer - Current < Val(interval)
DoEvents
Loop
End Sub

Sub SendCharNum(win, chars)
E = sendmessagebynum(win, WM_CHAR, chars, 0)

End Sub

Function SetChildFocus(child)
setchild% = SetFocusAPI(child)
End Function

Sub SetPreference()
GetMDI
Call RunMenuByString("Preferences")

Do: DoEvents
prefer% = findchildbytitle(MDI, "Preferences")
maillab% = findchildbytitle(prefer%, "Mail")
mailbut% = GetWindow(maillab%, GW_hWndNEXT)
If maillab% <> 0 And mailbut% <> 0 Then Exit Do
Loop

timeout (0.2)
ClickIcon (mailbut%)

Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Closewindows% = findchildbytitle(aolmod%, "Close mail after it has been sent")
aolconfirm% = findchildbytitle(aolmod%, "Confirm mail after it has been sent")
aolOK% = findchildbytitle(aolmod%, "OK")
If aolOK% <> 0 And Closewindows% <> 0 And aolconfirm% <> 0 Then Exit Do
Loop
sendcon% = SendMessage(Closewindows%, BM_SETCHECK, 1, 0)
sendcon% = SendMessage(aolconfirm%, BM_SETCHECK, 0, 0)

Click_Button (aolOK%)
Do: DoEvents
aolmod% = FindWindow("_AOL_Modal", "Mail Preferences")
Loop Until aolmod% = 0

closepre% = SendMessage(prefer%, WM_CLOSE, 0, 0)

End Sub

Sub stayontop(frm As Form)
Dim success%
success% = SetWindowPos(frm.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Sub runmenu(menu1 As Integer, menu2 As Integer)
Dim AOLWorks As Long
Static Working As Integer

AOLMenus% = GetMenu(FindWindow("AOL Frame25", vbNullString))
aolsubmenu% = GetSubMenu(AOLMenus%, menu1)
aolitemid = GetMenuItemID(aolsubmenu%, menu2)
AOLWorks = CLng(0) * &H10000 Or Working
ClickAOLMenu = sendmessagebynum(FindWindow("AOL Frame25", vbNullString), 273, aolitemid, 0&)

End Sub

Sub RMBS(ApplicationOfMenu, STringToSearchFor)
SearchString$ = STringToSearchFor
hMenu = GetMenu(ApplicationOfMenu)
Cnt = GetMenuItemCount(hMenu)
For i = 0 To Cnt - 1
DoEvents
PopUphMenu = GetSubMenu(hMenu, i)
Cnt2 = GetMenuItemCount(PopUphMenu)
For O = 0 To Cnt2 - 1
DoEvents
    hMenuID = GetMenuItemID(PopUphMenu, O)
    MenuString$ = String$(100, " ")
    X = GetMenuString(PopUphMenu, hMenuID, MenuString$, 100, 1)
    If InStr(UCase(MenuString$), UCase(SearchString$)) Then
        SendtoID = hMenuID
        GoTo Initiate
    End If
Next O
Next i
Initiate:
X = sendmessagebynum(ApplicationOfMenu, &H111, SendtoID, 0)
End Sub


Sub WaitWindow()
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
topmdi = GetWindow(MDI, 5)

Do: DoEvents
AOL = FindWindow("AOL Frame25", vbNullString)
MDI = FindChildByClass(AOL, "MDIClient")
topmdi2 = GetWindow(MDI, 5)
If Not topmdi2 = topmdi Then Exit Do
Loop

End Sub


Function FreeProcess()
Do: DoEvents
Process = Process + 1
If Process = 50 Then Exit Do
Loop
'frees process of freezes in your program
'and other stuff that makes your program
'slow down.  Works great.

End Function

Public Function GetIMSender() As String
'This will get the person sending the IM
GetMDI
im% = FindChildByTitlePartial(MDI, ">Instant Message From:")
If im% Then
imlength = GetWindowTextLength(im)
imtitle$ = Space(imlength)
Call GetWindowText(im, imtitle$, (imlength + 1))
finper = Mid(imtitle, 23)
GetIMSender = finper
End If
GetMDI
im% = findchildbytitle(MDI, "  Instant Message From:")
If im% Then
End If
End Function
Public Sub GetAOL()
'This Will return AOLs handle and AOL
'is a public var it can be used every
'where.
'Example
'GetAOL
'MDI = FindChildByClass(AOL, "MDIClient")
AOL = FindWindow("AOL Frame25", vbNullString)
End Sub
Public Sub GetMDI()
'This Will return the MDI handle and MDI
'is a public var it can be used every
'where.
GetAOL
MDI = FindChildByClass(AOL, "MDIClient")
End Sub

Sub PlaySound(E1)
'This will play a wav file
SoundName$ = E1
wFlags% = SND_ASYNC Or SND_NODEFAULT
X% = sndPlaySound(SoundName$, wFlags%)
End Sub
Public Sub Clicker(Handle)
a = SendMessage(Handle, WM_LBUTTONDOWN, 0, 0)
b = SendMessage(Handle, WM_LBUTTONUP, 0, 0)
End Sub
Public Function FindChatRoom()
GetAOL
GetMDI
windws = 0
Chat = FindChildByClass(MDI, "AOL CHILD")
LB = FindChildByClass(Chat, "_AOL_LISTBOX")
EB = FindChildByClass(Chat, "_AOL_EDIT")
If EB <> 0 And LB <> 0 Then
FindChatRoom = Chat
Exit Function
End If
Do
Chat = GetWindow(Chat, GW_hWndNEXT)
LB = FindChildByClass(Chat, "_AOL_LISTBOX")
EB = FindChildByClass(Chat, "_AOL_EDIT")
If EB <> 0 And LB <> 0 Then
FindChatRoom = Chat
Exit Function
End If
windws = windws + 1
Loop While windws < 100
Chat = 0

End Function
Public Function GetRoomFull()
'This finds the window that says this
'room is full

GetRoomFull = FindWindow("#32770", "America Online")
End Function
