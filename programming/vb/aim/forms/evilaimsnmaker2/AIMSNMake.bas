Attribute VB_Name = "AIMSNMake"
'most of the public declare functions i took
'from my source60.bas (there are ALOT, that arent
'even used in this coding but for the chatsends, and
'aol options, i didnt have enough time to pick out
'the minimal amount needed so i added them all...
'sorry if this messes you up at all. -source

Option Explicit
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByNum& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    'Constant
    Const WM_VSCROLL = &H115
    Const SB_BOTTOM = 7
    'constant
Public Const LB_ADDSTRING& = &H180
Public Const LB_DELETESTRING = &H182
Public Const LB_FINDSTRINGEXACT& = &H1A2
Public Const LB_GETCOUNT& = &H18B
Public Const LB_GETCURSEL& = &H188
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN& = &H18A
Public Const LB_INSERTSTRING = &H181
Public Const LB_RESETCONTENT& = &H184
Public Const LB_SETHORIZONTALEXTENT = &H194
Public Const LB_SETSEL = &H185
Public Const WM_MOVE = &HF012
Public Const WM_SYSCOMMAND = &H112
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE




Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long

Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function ExitWindowsEx& Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long)
Public Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "Kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "Kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal length As Long)
Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Long) As Long
'Public Const

Public Const LB_SETCURSEL = &H186
Public Const SW_SHOW = 5
Public Const SW_SHOWNORMAL = 1
Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3
Public Const BM_SETCHECK = &HF1
Public Const BM_GETCHECK = &HF0
Public Const WM_CLOSE = &H10
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101
Public Const WM_LBUTTONUP = &H202
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26
Public Const HWND_NOTOPMOST = -2

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000
Public Const ENTER_KEY = 13
Global Const GW_CHILD = 5

Public Type POINTAPI
      X As Long
      Y As Long
End Type

Public Const Op_Flags = PROCESS_READ Or RIGHTS_REQUIRED

Public Const SW_RESTORE = 9














Public Const CB_ADDSTRING& = &H143
Public Const CB_DELETESTRING& = &H144
Public Const CB_FINDSTRINGEXACT& = &H158
Public Const CB_GETCOUNT& = &H146
Public Const CB_GETITEMDATA = &H150
Public Const CB_GETLBTEXT& = &H148
Public Const CB_RESETCONTENT& = &H14B
Public Const CB_SETCURSEL& = &H14E

Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10

Public Const Sys_Add = &H0
Public Const Sys_Delete = &H2
Public Const Sys_Message = &H1
Public Const Sys_Icon = &H2
Public Const Sys_Tip = &H4

Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT
Public Const Snd_Flag2 = SND_ASYNC Or SND_LOOP







Public Const WM_CHAR = &H102

Public Const WM_CLEAR = &H303
Public Const WM_MOUSEMOVE = &H200
Public Const WM_COMMAND = &H111




Public Const MF_BYPOSITION = &H400&

Public Const EM_GETLINECOUNT& = &HBA





Public Enum MAILTYPE
        mailFLASH
        mailNEW
        mailOLD
        mailSENT
End Enum

Public systray As NOTIFYICONDATA

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uId As Long
        uFlags As Long
        ucallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type

Public Enum OnScreen
    scon
    scoff
End Enum



Public Sub OnTOP(FormName As Form)
    Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
Public Sub Drag(TheForm As Form)
    Call ReleaseCapture 'public declare
    
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Function Pause(Time As Long)
    'pause for a certain amount of time
    'Call pause(1)
    Dim Current As Long
    
    Current = Timer
    
    Do Until Timer - Current >= Time
        
        DoEvents
     
     Loop
End Function
Public Sub Load2Lists(ListSN As ListBox, ListPW As ListBox, Target As String)
    'self explanatory
    On Error Resume Next
    
    Dim lstInput As String, strSN As String, strPW As String
    
    If FileExists(Target$) = True Then
        Open Target$ For Input As #1
            While Not EOF(1) = True
                'DoEvents
                Input #1, lstInput$
                If InStr(1, lstInput$, "]-[") <> 0& And InStr(1, lstInput$, "=") <> 0& Then
                    lstInput$ = Mid(lstInput$, InStr(1, lstInput$, "]-[") + 3, Len(lstInput$) - 6)
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "=") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "="))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, ":") <> 0& Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, ":") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, ":"))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, "-") Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "-") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "-"))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, "=") Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "=") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "="))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                ElseIf InStr(1, lstInput$, "·") Then
                    strSN$ = Left(lstInput$, InStr(1, lstInput$, "·") - 1)
                    strPW$ = Right(lstInput$, Len(lstInput$) - InStr(1, lstInput$, "·"))
                    If Trim(strSN$) <> "" And Trim(strPW$) <> "" Then
                        ListSN.AddItem Trim(strSN$)
                        ListPW.AddItem Trim(strPW$)
                    End If
                End If
            Wend
        Close #1
    End If
End Sub
Public Sub Save2Lists(ListSN As ListBox, ListPW As ListBox, Target As String)
    'self explanatory
    Dim sLong As Long
    
    On Error Resume Next
    
    Open Target$ For Output As #1
        
        For sLong& = 0 To ListSN.ListCount - 1
            
            Print #1, "" + ListSN.List(sLong&) + ":" + ListPW.List(sLong&) + ""
        
        Next sLong&
    
    Close #1
End Sub
Public Function FileExists(TheFileName As String) As Boolean
'Sees if the string(file) you specified exists
If Len(TheFileName$) = 0 Then

FileExists = False

Exit Function

End If

If Len(Dir$(TheFileName$)) Then

FileExists = True

Else

FileExists = False

End If

End Function
Public Sub AddHScroll(List As ListBox)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    'Find Longest Text in Listbox


    For i = 0 To List.ListCount - 1


        If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    'Get Twips
    lngGreatestWidth = List.Parent.TextWidth(List.List(intGreatestLen) + Space(1))
    'Space(1) is used to prevent the last Ch
    '     aracter from being cut off
    'Convert to Pixels
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    'Use api to add scrollbar
    SendMessage List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    
End Sub
Public Sub ListRemoveSelected(ListBox As ListBox)
        Dim ListCount As Long
ListCount& = ListBox.ListCount
Do While ListCount& > 0&
ListCount& = ListCount& - 1
If ListBox.Selected(ListCount&) = True Then
ListBox.RemoveItem (ListCount&)
End If
Loop
End Sub


Sub SaveFormState(ByVal SourceForm As Form)
 Dim a As Long ' general purpose
 Dim b As Long
 Dim c As Long
 Dim FileName As String ' where to save to
 Dim FHandle As Long ' FileHandle
 ' error handling code
 On Error GoTo fError
 ' we create a filename based on the formname
 FileName = App.Path + "\" + SourceForm.name + ".set"
 ' Get a filehandle
 FHandle = FreeFile()
 ' open the file
 #If DebugMode = 1 Then
  Debug.Print "--------------------------------------------------------->"
  Debug.Print "Saving Form State:" + SourceForm.name
  Debug.Print "FileName=" + FileName
 #End If
 Open FileName For Output As FHandle
 ' loop through all controls
 ' first we save the type then the name
 For a = 0 To SourceForm.Controls.Count - 1
  #If DebugMode = 1 Then
   Debug.Print "Saving control:" + SourceForm.Controls(a).name
  #End If
  ' if its textbox we save the .text property
  If TypeOf SourceForm.Controls(a) Is TextBox Then
   Print #FHandle, "TextBox"
   Print #FHandle, SourceForm.Controls(a).name
   Print #FHandle, "StartText"
   Print #FHandle, SourceForm.Controls(a).Text
   Print #FHandle, "EndText"
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a checkbox we save the .value property
  If TypeOf SourceForm.Controls(a) Is CheckBox Then
   Print #FHandle, "CheckBox"
   Print #FHandle, SourceForm.Controls(a).name
   Print #FHandle, Str(SourceForm.Controls(a).value)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a option button we save its value
  If TypeOf SourceForm.Controls(a) Is OptionButton Then
   Print #FHandle, "OptionButton"
   Print #FHandle, SourceForm.Controls(a).name
   Print #FHandle, Str(SourceForm.Controls(a).value)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a listbox we save the .text and list contents
  If TypeOf SourceForm.Controls(a) Is ListBox Then
   Print #FHandle, "ListBox"
   Print #FHandle, SourceForm.Controls(a).name
   Print #FHandle, SourceForm.Controls(a).Text
   Print #FHandle, "StartList"
   For b = 0 To SourceForm.Controls(a).ListCount - 1
    Print #FHandle, SourceForm.Controls(a).List(b)
   Next b
   Print #FHandle, "EndList"
   ' save listindex
   Print #FHandle, CStr(SourceForm.Controls(a).ListIndex)
    ' print a separator
   Print #FHandle, "|<->|"
  End If
  ' if its a combobox, save .text and list items
  If TypeOf SourceForm.Controls(a) Is ComboBox Then
   Print #FHandle, "ComboBox"
   Print #FHandle, SourceForm.Controls(a).name
   Print #FHandle, SourceForm.Controls(a).Text
   Print #FHandle, "StartList"
   For b = 0 To SourceForm.Controls(a).ListCount - 1
    Print #FHandle, SourceForm.Controls(a).List(b)
   Next b
   Print #FHandle, "EndList"
    ' print a separator
   Print #FHandle, "|<->|"
  End If
 Next a
' close file
 #If DebugMode = 1 Then
  Debug.Print "Closing File."
  Debug.Print "<----------------------------------------------------------"
 #End If
 Close #FHandle
 ' stop error handler
 On Error GoTo 0
 Exit Sub
fError: ' Simple error handler
 c = MsgBox("Error in SaveFormState. " + Err.Description + ", Number=" + CStr(Err.Number), vbAbortRetryIgnore)
 If c = vbIgnore Then Resume Next
 If c = vbRetry Then Resume
 ' else abort
End Sub
'@===========================================================================
' LoadFormState:
'  Loads the state of controls from file
'
'  Currently Supports: TextBox, CheckBox, OptionButton, Listbox, ComboBox
'=============================================================================
Sub LoadFormState(ByVal SourceForm As Form)
 Dim a As Long ' general purpose
 Dim b As Long
 Dim c As Long
 
 Dim txt As String ' general purpose
 Dim fData As String ' used to hold File Data
' these are variables used for controls data
 Dim cType As String ' Type of control
 Dim Cname As String ' Name of control
 Dim cNum As Integer ' number of control
' vars for the file
 Dim FileName As String ' where to save to
 Dim FHandle As Long ' FileHandle
 ' error handling code
 'On Error GoTo fError
 ' we create a filename based on the formname
 FileName = App.Path + "\" + SourceForm.name + ".set"
 ' abort if file does not exist
 If Dir(FileName) = "" Then
  #If DebugMode = 1 Then
   Debug.Print "File Not found:" + FileName
  #End If
  Exit Sub
 End If
 ' Get a filehandle
 FHandle = FreeFile()
 ' open the file
 #If DebugMode = 1 Then
  Debug.Print "------------------------------------------------------>"
  Debug.Print "Loading FormState:" + SourceForm.name
  Debug.Print "FileName:" + FileName
 #End If
 Open FileName For Input As FHandle
' go through file
 While Not EOF(FHandle)
  Line Input #FHandle, cType
  Line Input #FHandle, Cname
  ' Get control number
  cNum = -1
  For a = 0 To SourceForm.Controls.Count - 1
   If SourceForm.Controls(a).name = Cname Then cNum = a
  Next a
  ' add some debug info if in debugmode
  #If DebugMode = 1 Then
   Debug.Print "Control Type=" + cType
   Debug.Print "Control Name=" + Cname
   Debug.Print "Control Number=" + CStr(cNum)
  #End If
  ' if we find control
  If Not cNum = -1 Then
   ' Depending on type of control, what data we get
   Select Case cType
   Case "TextBox"
    Line Input #FHandle, fData
    fData = "": txt = ""
    While Not fData = "EndText"
     If Not txt = "" Then txt = txt + vbCrLf
     txt = txt + fData
     Line Input #FHandle, fData
    Wend
    ' update control
    SourceForm.Controls(cNum).Text = txt
   Case "CheckBox"
    ' we get the value
    Line Input #FHandle, fData
    ' update control
    SourceForm.Controls(cNum).value = fData
   Case "OptionButton"
    ' we get the value
    Line Input #FHandle, fData
    ' update control
    SourceForm.Controls(cNum).value = fData
   Case "ListBox"
    ' clear listbox
    SourceForm.Controls(cNum).Clear
    ' get .text property
    Line Input #FHandle, fData
    SourceForm.Controls(cNum).Text = fData
    ' read past /startlist
    Line Input #FHandle, fData
    fData = "": txt = ""
    ' Get List
    While Not fData = "EndList"
     If Not fData = "" Then SourceForm.Controls(cNum).AddItem fData
     Line Input #FHandle, fData
    Wend
    ' get listindex
     Line Input #FHandle, fData
     SourceForm.Controls(cNum).ListIndex = Val(fData)
   Case "ComboBox"
    ' Clear combobox
    SourceForm.Controls(cNum).Clear
    ' Get Text
    Line Input #FHandle, fData
    SourceForm.Controls(cNum).Text = fData
    ' readpast /startlist
    Line Input #FHandle, fData
    fData = "": txt = ""
    ' get list
    While Not fData = "EndList"
     If Not fData = "" Then SourceForm.Controls(cNum).AddItem fData
     Line Input #FHandle, fData
    Wend
   End Select ' what type of control
  End If ' if we found control
  ' read till seperator
  fData = ""
  While Not fData = "|<->|"
   Line Input #FHandle, fData
  Wend
 Wend ' not end of File (EOF)
' close file
 #If DebugMode = 1 Then
  Debug.Print "Closing file.."
  Debug.Print "<------------------------------------------------------"
 #End If
 Close #FHandle
 Exit Sub
fError: ' Simple error handler
 c = MsgBox("Error in LoadFormState. " + Err.Description + ", Number=" + CStr(Err.Number), vbAbortRetryIgnore)
 If c = vbIgnore Then Resume Next
 If c = vbRetry Then Resume
 ' else abort
End Sub
Public Function ChatSend(Text As String)
'Send chat to chatroom
'Call chatsend("source ownz")

'if chat notify isnt checked then exit the function
If menu.Check2.value = 0 Then Exit Function

Dim richcntl As Long
richcntl& = FindWindowEx(FindChatRoom(), 0&, "RICHCNTL", vbNullString)
richcntl& = FindWindowEx(FindChatRoom(), richcntl&, "RICHCNTL", vbNullString)
Call SendMessageByString(richcntl&, WM_SETTEXT, 0&, Text$)
Call SendMessageByNum(richcntl&, WM_CHAR, 13, 0&)
End Function
Public Function FindChatRoom() As Long
'prozac
'Finds the aol chatroom
'Example:
'If Findchatroom <> 0& then
'msgbox GetWindowCaption(Findchatroom) + "window found!"
'else
'msgbox "chatroom not found!"
'end if
Dim counter As Long
Dim AOLStatic5 As Long
Dim AOLIcon3 As Long
Dim AOLStatic4 As Long
Dim aollistbox As Long
Dim AOLStatic3 As Long
Dim AOLImage As Long
Dim AOLIcon2 As Long
Dim RICHCNTL2 As Long
Dim AOLStatic2 As Long
Dim i As Long
Dim AOLIcon As Long
Dim AOLCombobox As Long
Dim richcntl As Long
Dim AOLStatic As Long
Dim AOLChild As Long
Dim MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString)
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString)
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString)
AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
For i& = 1& To 3&
    AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
For i& = 1& To 2&
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
Next i&
AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
aollistbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
For i& = 1& To 7&
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
Next i&
AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
Do While (counter& <> 100&) And (AOLStatic& = 0& Or richcntl& = 0& Or AOLCombobox& = 0& Or AOLIcon& = 0& Or AOLStatic2& = 0& Or RICHCNTL2& = 0& Or AOLIcon2& = 0& Or AOLImage& = 0& Or AOLStatic3& = 0& Or aollistbox& = 0& Or AOLStatic4& = 0& Or AOLIcon3& = 0& Or AOLStatic5& = 0&): DoEvents
    AOLChild& = FindWindowEx(MDIClient&, AOLChild&, "AOL Child", vbNullString)
    AOLStatic& = FindWindowEx(AOLChild&, 0&, "_AOL_Static", vbNullString)
    richcntl& = FindWindowEx(AOLChild&, 0&, "RICHCNTL", vbNullString)
    AOLCombobox& = FindWindowEx(AOLChild&, 0&, "_AOL_Combobox", vbNullString)
    AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString)
    For i& = 1& To 3&
        AOLIcon& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic2& = FindWindowEx(AOLChild&, AOLStatic&, "_AOL_Static", vbNullString)
    RICHCNTL2& = FindWindowEx(AOLChild&, richcntl&, "RICHCNTL", vbNullString)
    AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon&, "_AOL_Icon", vbNullString)
    For i& = 1& To 2&
        AOLIcon2& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    Next i&
    AOLImage& = FindWindowEx(AOLChild&, 0&, "_AOL_Image", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic2&, "_AOL_Static", vbNullString)
    AOLStatic3& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    aollistbox& = FindWindowEx(AOLChild&, 0&, "_AOL_Listbox", vbNullString)
    AOLStatic4& = FindWindowEx(AOLChild&, AOLStatic3&, "_AOL_Static", vbNullString)
    AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon2&, "_AOL_Icon", vbNullString)
    For i& = 1& To 7&
        AOLIcon3& = FindWindowEx(AOLChild&, AOLIcon3&, "_AOL_Icon", vbNullString)
    Next i&
    AOLStatic5& = FindWindowEx(AOLChild&, AOLStatic4&, "_AOL_Static", vbNullString)
    If AOLStatic& And richcntl& And AOLCombobox& And AOLIcon& And AOLStatic2& And RICHCNTL2& And AOLIcon2& And AOLImage& And AOLStatic3& And aollistbox& And AOLStatic4& And AOLIcon3& And AOLStatic5& Then Exit Do
    counter& = Val(counter&) + 1&
Loop
If Val(counter&) < 100& Then
    FindChatRoom& = AOLChild&
    Exit Function
End If
End Function

Public Function adv()
Call ChatSend("<font color=black></b></u></i>< a href=www.vbfx.net><font color=#000000></u>`•–•» <sup><b>è</b></sup>víl <b>á</b>im <b>s</b>n <b>m</b>àkèr ²</font></a> - <b>s</b>ôúrçè</font>")
End Function
Public Function chatsend2(txt As String)
ChatSend ("<font color=black></b></u></i>< a href=www.vbfx.net><font color=#000000></u>`•–•»" & txt & "</font></a>")
End Function

Public Function sendaims(txt As String)
ChatSend ("<font color=black></b></u></i>< a href=www.vbfx.net><font color=#000000></u>`•–•»<b>" & txt & "</font></a></b> - was created")
End Function

