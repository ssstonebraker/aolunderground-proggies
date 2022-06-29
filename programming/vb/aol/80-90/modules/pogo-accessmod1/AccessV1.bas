Attribute VB_Name = "AccessV1"
'-----------------------------------------------------------'
'          //\                                              '
'         / \ \                                             '
'        /   \ \                                            '
'       / //\ \ \                                           '
'      /  ___  \ \                                          '
'     /__//  \__\/ccess Version 1.2                         '
' Author:   Jake a.k.a Pogo / Team Access 2k3               '
'   Date:   May and Early June of 2003                      '
'  Usage:   aol 7 / 8                                       '
'                                                           '
'   Subs:   110+                                            '
'           all subs created by jake unless otherwise noted!'
'   Note:   its a little messy with the dimensioning and    '
'           ordering and stuff but it all works!            '
'  Props:   Lauren*,Clown,Morph,Kevin aka Seven,Snypa,Sinn, '
'           Brandon aka Zippo,Def,Adam,Chrome,The MLD Crew, '
'           Real,Nyte,Epik, and specially to the REAL       '
'                           DeDux [RIP]                     '
'                       Partying in Heaven.                 '
'-----------------------------------------------------------'
Option Explicit
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Integer, ByVal wCmd As Integer) As Integer
Public Declare Function getclassname& Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long)
Public Declare Function GetWindowClassName Lib "user32" Alias "GetWindowClassname" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const ENTER_KEY = 13
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const GW_CHILD = 5
Public Const LB_GETCOUNT = &H18B
Public Const LB_GETITEMDATA = &H199
Public Const LB_GETTEXT = &H189
Public Const LB_GETTEXTLEN = &H18A
Public Const LB_SETCURSEL = &H186
Public Const LB_SETSEL = &H185

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_MENU = &H12
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

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
Public Const WM_SETTEXT = &HC
Public Const WM_SYSCOMMAND = &H112

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_DWORD = 4

Const REG_OPTION_NON_VOLATILE = 0
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_PERFORMANCE_DATA = &H80000004

Public Const CB_GETCOUNT = &H146
Public Const CB_GETLBTEXT = &H148
Public Const CB_SETCURSEL = &H14E

Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type
Public Type POINTAPI
        X As Long
        Y As Long
End Type
Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Type info
    sTitle As String * 30
    sArtist As String * 30
    sAlbum As String * 30
    sComment As String * 30
    sYear As String * 4
End Type
Type HeaderInfo
    Layer As String
    Frequency As String
    Bitrate As String
    moDe As String
    MpegVersion As String
    Emphasis As String
    FPlayTime As String
    mFileSize As String
End Type
Public Mp3Info As info
Public MP3HeaderInfo As HeaderInfo


Global ha As Integer
Global hey As String
Public Sub Aim_SignOn(sn As String, pw As String)
Dim aimcsignonwnd As Long, combobox As Long, editx As Long
Dim oscariconbtn As Long
aimcsignonwnd = FindWindow("aim_csignonwnd", vbNullString)
combobox = FindWindowEx(aimcsignonwnd, 0&, "combobox", vbNullString)
editx = FindWindowEx(combobox, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, sn$)
editx = FindWindowEx(aimcsignonwnd, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, pw$)
oscariconbtn = FindWindowEx(aimcsignonwnd, 0&, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimcsignonwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
oscariconbtn = FindWindowEx(aimcsignonwnd, oscariconbtn, "_oscar_iconbtn", vbNullString)
Call ClickIcon(oscariconbtn)

End Sub
Public Sub AOL_SignOnGuest(sn As String, pw As String)
Dim editx As Long
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolcombobox As Long
Dim LCount As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
aolcombobox = FindWindowEx(aolchild, 0&, "_aol_combobox", vbNullString)
Dim LIndex As Long
LCount = SendMessageLong(aolcombobox, CB_GETCOUNT, 0&, 0&)
LIndex = LCount - 1
Call SendMessageLong(aolcombobox, CB_SETCURSEL, LIndex, 0&)

Dim aolicon As Long
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
Dim aoledit As Long
Dim aolmodal As Long
Do
DoEvents
Dim signmeon As Long
signmeon = FindWindow("_aol_modal", vbNullString)
aoledit = FindWindowEx(signmeon, 0&, "_aol_edit", vbNullString)
Loop Until signmeon& <> 0& And aoledit& <> 0&
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, sn$)
aoledit = FindWindowEx(signmeon, aoledit, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, pw$)
aolicon = FindWindowEx(signmeon, 0&, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
Call FindRoomFull
End Sub

Sub LoadResStrings(frm As Form)
  On Error Resume Next
  
  Dim ctl As Control
  Dim obj As Object
  

  If IsNumeric(frm.Tag) Then
    frm.Caption = LoadResString(CInt(frm.Tag))
  End If
  
  For Each ctl In frm.Controls
    Err.Clear
    If TypeName(ctl) = "Menu" Then
      If IsNumeric(ctl.Caption) Then
        If Err = 0 Then
          ctl.Caption = LoadResString(CInt(ctl.Caption))
        End If
      End If
    ElseIf TypeName(ctl) = "TabStrip" Then
      For Each obj In ctl.Tabs
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.Caption = LoadResString(CInt(obj.Tag))
        End If
        'check for a tooltip
        If IsNumeric(obj.ToolTipText) Then
          If Err = 0 Then
            obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
          End If
        End If
      Next
    ElseIf TypeName(ctl) = "Toolbar" Then
      For Each obj In ctl.Buttons
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.ToolTipText = LoadResString(CInt(obj.Tag))
        End If
      Next
    ElseIf TypeName(ctl) = "ListView" Then
      For Each obj In ctl.ColumnHeaders
        Err.Clear
        If IsNumeric(obj.Tag) Then
          obj.text = LoadResString(CInt(obj.Tag))
        End If
      Next
    Else
      If IsNumeric(ctl.Tag) Then
        If Err = 0 Then
          ctl.Caption = LoadResString(CInt(ctl.Tag))
        End If
      End If
 
      If IsNumeric(ctl.ToolTipText) Then
        If Err = 0 Then
          ctl.ToolTipText = LoadResString(CInt(ctl.ToolTipText))
        End If
      End If
    End If
  Next

End Sub

Public Function UpdateKey(KeyRoot As Long, KeyName As String, SubKeyName As String, SubKeyValue As String) As Boolean
    'how to use:
'msgbox UpdateKey(HKEY_CLASSES_ROOT, "keyname", "newvalue")

    Dim rc As Long
    Dim hKey As Long
    Dim hDepth As Long
    Dim lpAttr As SECURITY_ATTRIBUTES
    
    lpAttr.nLength = 50
    lpAttr.lpSecurityDescriptor = 0
    lpAttr.bInheritHandle = True
    rc = RegCreateKeyEx(KeyRoot, KeyName, _
                        0, REG_SZ, _
                        REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpAttr, _
                        hKey, hDepth)
    
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError

    If (SubKeyValue = "") Then SubKeyValue = " "
    
    rc = RegSetValueEx(hKey, SubKeyName, _
                       0, REG_SZ, _
                       SubKeyValue, LenB(StrConv(SubKeyValue, vbFromUnicode)))
                       
    If (rc <> ERROR_SUCCESS) Then GoTo CreateKeyError
    rc = RegCloseKey(hKey)
    
    UpdateKey = True
    Exit Function
CreateKeyError:
    UpdateKey = False
    rc = RegCloseKey(hKey)
End Function

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'how to use it
'msgbox GetKeyValue(HKEY_CLASSES_ROOT, "COMCTL.ListviewCtrl.1\CLSID", "")
    
    Dim i As Long
    Dim rc As Long
    Dim hKey As Long
    Dim hDepth As Long
    Dim sKeyVal As String
    Dim lKeyValType As Long
    Dim tmpVal As String
    Dim KeyValSize As Long
    
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey)
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
    
    tmpVal = String$(1024, 0)
    KeyValSize = 1024
    
   rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    Select Case lKeyValType
    Case REG_SZ, REG_EXPAND_SZ
        sKeyVal = tmpVal
    Case REG_DWORD
        For i = Len(tmpVal) To 1 Step -1
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))
        Next
        sKeyVal = Format$("&h" + sKeyVal)
    End Select
    
    GetKeyValue = sKeyVal
    rc = RegCloseKey(hKey)
    Exit Function
    
GetKeyError:
    GetKeyValue = vbNullString
    rc = RegCloseKey(hKey)
End Function



Public Function String_TrimSpaces(MainString As String) As String
        Dim NewMain As String, Instance As Long
NewMain$ = MainString$
Do While InStr(1, NewMain$, " ")
Instance& = InStr(1, NewMain$, " ")
NewMain$ = Left(NewMain$, (Instance& - 1)) & "" & Right(NewMain$, Len(NewMain$) - Instance&)
Loop
String_TrimSpaces$ = NewMain$
End Function
Public Function FindIM() As Long
'
'Created by clown
'
    Dim aolframe As Long, mdiclient As Long, aolchild As Long
    aolframe = FindWindow("aol frame25", vbNullString)
    mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
    aolchild = FindWindowEx(mdiclient, 0&, "aol child", vbNullString)
    aolchild = FindWindowEx(mdiclient, aolchild, "aol child", vbNullString)
    Dim Winkid1 As Long, Winkid2 As Long, Winkid3 As Long, Winkid4 As Long, Winkid5 As Long, Winkid6 As Long, Winkid7 As Long, Winkid8 As Long, Winkid9 As Long, FindOtherWin As Long
    FindOtherWin = GetWindow(aolchild, GW_HWNDFIRST)
    Do While FindOtherWin <> 0
           DoEvents
           Winkid1 = FindWindowEx(FindOtherWin, 0&, "_aol_view", vbNullString)
           Winkid2 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid3 = FindWindowEx(FindOtherWin, 0&, "richcntlreadonly", vbNullString)
           Winkid4 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid5 = FindWindowEx(FindOtherWin, 0&, "_aol_editimage", vbNullString)
           Winkid6 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid7 = FindWindowEx(FindOtherWin, 0&, "_aol_icon", vbNullString)
           Winkid8 = FindWindowEx(FindOtherWin, 0&, "_aol_static", vbNullString)
           Winkid9 = FindWindowEx(FindOtherWin, 0&, "richcntl", vbNullString)
           If (Winkid1 <> 0) And (Winkid2 <> 0) And (Winkid3 <> 0) And (Winkid4 <> 0) And (Winkid5 <> 0) And (Winkid6 <> 0) And (Winkid7 <> 0) And (Winkid8 <> 0) And (Winkid9 <> 0) Then
                  FindIM = FindOtherWin
                  Exit Function
           End If
           FindOtherWin = GetWindow(FindOtherWin, GW_HWNDNEXT)
    Loop
    FindIM = 0
End Function
Public Function FindNewIM() As Long
    Dim AOL As Long, MDI As Long, child As Long, Caption As String
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Caption$ = GetCaption(child&)
    If InStr(Caption$, "Send Instant Message") = 1 Or InStr(Caption$, "Send Instant Message") = 2 Or InStr(Caption$, "Send Instant Message") = 3 Then
        FindNewIM& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Caption$ = GetCaption(child&)
            If InStr(Caption$, "Send Instant Message") = 1 Or InStr(Caption$, "Send Instant Message") = 2 Or InStr(Caption$, "Send Instant Message") = 3 Then
                FindNewIM& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindNewIM& = child&
End Function
Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
   Dim strBuffer As String
   strBuffer = String(750, Chr(0))
   Key$ = LCase$(Key$)
   GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Function Window_GetCaption(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    Window_GetCaption$ = buffer$
End Function
Public Function GetCaption(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = GetWindowTextLength(WindowHandle&)
    buffer$ = String(TextLength&, 0&)
    Call GetWindowText(WindowHandle&, buffer$, TextLength& + 1)
    GetCaption$ = buffer$
End Function
Public Sub Pause(Duration As Double)
    Dim Current As Double
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub
Public Sub Chat_PrivateRoom7(room As String)
Call Keyword7("aol://2719:2-2-" & room)
End Sub
Public Sub Chat_PrivateRoom5(room As String)
Call Keyword5("aol://2719:2-2-" & room)
End Sub
Public Sub Chat_MemberRoom7(room As String)
    Call Keyword7("aol://2719:61-2-" & room$)
End Sub
Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
    Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub
Public Sub PlayWav(WavFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub
Public Sub Chat_PublicRoom7(room As String)
    Call Keyword7("aol://2719:21-2-" & room$)
End Sub
Public Sub Chat_PublicRoom5(room As String)
    Call Keyword5("aol://2719:21-2-" & room$)
End Sub
Public Function Chat_Count7() As Long
    Dim AOL As Long, MDI As Long, rMail As Long, rList As Long
    Dim Count As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDICLIENT", vbNullString)
    rList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    Count& = SendMessage(rList&, LB_GETCOUNT, 0&, 0&)
    Chat_Count7& = Count&
End Function
Public Sub List_Load(lst As ListBox, CmDlg As CommonDialog)
'needs microsoft common dialog control 6.0
Dim Jake$, ListItem$, X As Integer
On Error GoTo Op
CmDlg.CancelError = True
CmDlg.Filter = "All Files (*.*)|*.*"
CmDlg.FilterIndex = 0
CmDlg.ShowOpen
Jake$ = CmDlg.FileName
lst.Clear

Open Jake$ For Input As #1
Do While Not EOF(1)
    Line Input #1, ListItem$
    If Not (ListItem$ = "") Then
        lst.AddItem ListItem$
    End If
Loop
Close #1
Op: Exit Sub
End Sub
Public Sub List_RemoveSelected(lst As ListBox)
'Sad to say not many people know actual vb
'   the end if statement prevents trying to
'       remove if unselected.
If lst.ListIndex <> -1 Then
    lst.RemoveItem (lst.ListIndex)
End If
End Sub
Public Sub List_Add(text As String, lst As ListBox)
Call lst.AddItem(text, lst.ListCount)
End Sub
Public Sub List_Save(lst As ListBox, CmDlg As CommonDialog)
'needs microsoft common dialog controls 6.0
Dim Jake$, X As Integer
On Error GoTo O
CmDlg.CancelError = True
CmDlg.Filter = "All Files (*.*)|*.*"
CmDlg.FilterIndex = 0
CmDlg.ShowSave
Jake$ = CmDlg.FileName

If Len(Dir$(Jake$)) <> 0 Then
    X = MsgBox("This file already exists: " + Jake$ + ", do you wish replace it?", 33, "Error")
    If X > 1 Then Exit Sub
End If

Open Jake$ For Output As #1
For X = 0 To lst.ListCount - 1
    Print #1, lst.list(X) + Chr(13)
Next X
Close #1
O: Exit Sub
End Sub
Public Sub List_Save2Boxes(Directory As String, ListA As ListBox, ListB As ListBox)
    Dim SaveLists As Long
    On Error Resume Next
    Open Directory$ For Output As #1
    For SaveLists& = 0 To ListA.ListCount - 1
        Print #1, ListA.list(SaveLists&) & "*" & ListB.list(SaveLists)
    Next SaveLists&
    Close #1
End Sub
Public Sub Form_OnTop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub
Public Sub Form_NotOnTop(FormName As Form)
Call SetWindowPos(FormName.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
End Sub
Public Sub Form_ExitDown(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) + 300))
    Loop Until TheForm.Top > 7200
End Sub
Public Sub Form_ExitUp(TheForm As Form)
    Do
        DoEvents
        TheForm.Top = Trim(Str(Int(TheForm.Top) - 300))
    Loop Until TheForm.Top < -TheForm.Width
End Sub
Public Sub Form_ExitRight(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) + 300))
    Loop Until TheForm.Left > Screen.Width
End Sub
Public Sub Form_ExitLeft(TheForm As Form)
    Do
        DoEvents
        TheForm.Left = Trim(Str(Int(TheForm.Left) - 300))
    Loop Until TheForm.Left < -TheForm.Width
End Sub
Public Function MSN_FindIm() As Long
Dim msnIM As Long
Dim child As Long
msnIM = FindWindow("imwindowclass", vbNullString)

  If msnIM& <> 0& Then
        MSN_FindIm& = child&
        Exit Function
    Else
        Do
           msnIM = FindWindow("imwindowclass", vbNullString)
           If msnIM& <> 0& Then
                MSN_FindIm& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    MSN_FindIm& = child&

End Function
Public Function MSN_IMLastMessage() As String
Dim imwindowclass As Long, richedita As Long
Dim MsgString As String
Dim strMsgs
Dim intLoc
Dim strChar
imwindowclass = FindWindow("imwindowclass", vbNullString)
richedita = FindWindowEx(imwindowclass, 0&, "richedit20a", vbNullString)
richedita = FindWindowEx(imwindowclass, richedita, "richedit20a", vbNullString)
MsgString$ = GetText(richedita&)
MSN_IMLastMessage$ = MsgString$

 strMsgs = MSN_IMLastMessage$
        intLoc = InStrRev(strMsgs, Chr(13))
        strMsgs = Mid(strMsgs, intLoc + 1)
        intLoc = InStr(strMsgs, " ")
        Do: DoEvents
            intLoc = intLoc + 1
            strChar = Mid(strMsgs, intLoc, 1)
        Loop Until strChar <> " "
    MSN_IMLastMessage$ = Mid(strMsgs, intLoc)

End Function
Public Function FindRoom() As Long
    Dim AOL As Long, MDI As Long, child As Long
    Dim Rich As Long, AOLList As Long
    Dim aolicon As Long, AOLStatic As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(MDI&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And AOLStatic& <> 0& Then
        FindRoom& = child&
        Exit Function
    Else
        Do
            child& = FindWindowEx(MDI&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            AOLList& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            AOLStatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            If Rich& <> 0& And AOLList& <> 0& And aolicon& <> 0& And AOLStatic& <> 0& Then
                FindRoom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
    FindRoom& = child&
End Function

Public Sub Form_Drag(TheForm As Form)
    Call ReleaseCapture
    Call SendMessage(TheForm.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
Public Sub AOLChangeCaption(NewCap As String)
Dim AOL As Long
AOL& = FindWindow("AOL Frame25", vbNullString)
Call SendMessageByString(AOL&, WM_SETTEXT, 0, NewCap)
End Sub
Public Sub File_SetHidden(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbHidden
    End If
End Sub
Public Sub File_SetReadOnly(TheFile As String)
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        SetAttr TheFile$, vbReadOnly
    End If
End Sub
Public Function File_Exists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        File_Exists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        File_Exists = True
    Else
        File_Exists = False
    End If
End Function

Public Function TagExists(ByVal sPassFilename As String) As Boolean
 'Taken from sinns mod bas
 'ty sinn ;]
    Dim iFreefile As Integer
    Dim lFilePos As Long
    Dim sData As String * 128
    Mp3Info.sTitle = ""
    Mp3Info.sArtist = ""
    Mp3Info.sAlbum = ""
    Mp3Info.sYear = ""
    Mp3Info.sComment = ""
    If Dir(sPassFilename) = "" Then TagExists = False: GoTo closeme
    TagExists = True
    iFreefile = FreeFile
    lFilePos = FileLen(sPassFilename) - 127
    Open sPassFilename For Binary As #iFreefile
    Get #iFreefile, lFilePos, sData
    Close #iFreefile
        If Left(sData, 3) = "TAG" Then
        Mp3Info.sTitle = RTrim(Mid(sData, 4, 30))
        Mp3Info.sArtist = RTrim(Mid(sData, 34, 30))
        Mp3Info.sAlbum = RTrim(Mid(sData, 64, 30))
        Mp3Info.sYear = RTrim(Mid(sData, 94, 4))
        Mp3Info.sComment = RTrim(Mid(sData, 98, 30))
    End If
closeme:
Close #iFreefile
End Function
Public Function Mp3_Year(FileName As String)
Dim strp As String
If TagExists(FileName) = True Then
strp$ = Mp3Info.sYear
GoTo sTrim
sTrim:
If Right(strp, 1) = " " Then
strp$ = Left$(strp$, Len(strp$) - 1)
Mp3_Year = strp$
GoTo sTrim
Else
Mp3_Year = strp$
End If
End If
End Function
Public Function AOL()
AOL = FindWindow("AOL Frame25", vbNullString)
End Function
Public Function MDI()
MDI = FindWindowEx(AOL, 0&, "MDIClient", vbNullString)
End Function
Public Function FindWelcome()
Dim Window&, Caption$
Window& = FindWindowEx(MDI, 0&, "AOL Child", vbNullString)
Do
DoEvents
Caption$ = GetCaption(Window&)
If InStr(Caption$, "Welcome,") <> 0 Then
FindWelcome = Window&
Exit Function
End If
Window& = FindWindowEx(MDI, Window&, "AOL Child", vbNullString)
Loop Until Window& = 0
FindWelcome = 0
End Function
Public Sub Gen3Chr_NumericVowels(HowManySNs As Integer, list As ListBox)
    Dim AlphaNumericString As String, AlphaString As String
    Dim strLetter As String, sn As String, RandomTime As String
    Dim rndX As Integer, rndY As Integer, MakinSNs As Integer, i As Long
    Dim numericstring As String
    Dim numericvowelString As String

Dim spacestring As String
    Randomize
    RandomTime = Int(10 * Rnd)
    If RandomTime = 10 Then RandomTime = 9
    
    AlphaNumericString = "1234567890abcdefghijklmnopqrstuvwxyz"
    AlphaString = "aeioyu"
    numericstring = "123456789"
    numericvowelString = "7a1e2i3o4u5y6890"
Dim z As Long
For z = 1 To HowManySNs
        sn = ""
        
        For i = 0 To RandomTime
            rndX = Int((6 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndX, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((16 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(numericvowelString, rndY, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((16 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(numericvowelString, rndY, 1)
        sn = sn + strLetter
      If Len(sn) < 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
        If Len(sn) > 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
Next z

End Sub
Public Sub Gen3Chr_Pronouncables(HowManySNs As Integer, list As ListBox)
    Dim AlphaNumericString As String, AlphaString As String
    Dim strLetter As String, sn As String, RandomTime As String
    Dim rndX As Integer, rndY As Integer, MakinSNs As Integer, i As Long
    Dim numericstring As String
    Dim numericvowelString As String

Dim spacestring As String
    Randomize
    RandomTime = Int(10 * Rnd)
    If RandomTime = 10 Then RandomTime = 9
    
    AlphaNumericString = "abcdefghijklmnopqrstuvwxyz"
    AlphaString = "aeioyu"
    numericstring = "123456789"
    numericvowelString = "7a1e2i3o4u5y6890"
Dim z As Long
For z = 1 To HowManySNs
        sn = ""
        
        For i = 0 To RandomTime
            rndX = Int((26 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaNumericString, rndX, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((6 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndY, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((26 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaNumericString, rndY, 1)
        sn = sn + strLetter
        If Len(sn) < 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
        If Len(sn) > 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
Next z

End Sub
Public Sub Gen3Chr_MixedAll(HowManySNs As Integer, list As ListBox)
    Dim AlphaNumericString As String, AlphaString As String
    Dim strLetter As String, sn As String, RandomTime As String
    Dim rndX As Integer, rndY As Integer, MakinSNs As Integer, i As Long
    Dim numericstring As String
Dim spacestring As String
    Randomize
    RandomTime = Int(10 * Rnd)
    If RandomTime = 10 Then RandomTime = 9
    
    AlphaNumericString = "1234567890abcdefghijklmnopqrstuvwxyz"
    AlphaString = "abcdefghijklmnopqrstuvwxyz"
    numericstring = "123456789"
Dim z As Long
For z = 1 To HowManySNs
        sn = ""
        
        For i = 0 To RandomTime
            rndX = Int((26 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndX, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((35 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaNumericString, rndY, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((35 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaNumericString, rndY, 1)
        sn = sn + strLetter
        If Len(sn) < 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
        If Len(sn) > 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
Next z

End Sub
Public Sub Gen3Chr_LettersOnly(HowManySNs As Integer, list As ListBox)
    Dim AlphaNumericString As String, AlphaString As String
    Dim strLetter As String, sn As String, RandomTime As String
    Dim rndX As Integer, rndY As Integer, MakinSNs As Integer, i As Long
    Dim numericstring As String
Dim spacestring As String
    Randomize
    RandomTime = Int(10 * Rnd)
    If RandomTime = 10 Then RandomTime = 9
    
    AlphaNumericString = "1234567890abcdefghijklmnopqrstuvwxyz"
    AlphaString = "abcdefghijklmnopqrstuvwxyz"
    numericstring = "123456789"
Dim z As Long
For z = 1 To HowManySNs
        sn = ""
        
        For i = 0 To RandomTime
            rndX = Int((26 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndX, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((26 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndY, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((26 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndY, 1)
        sn = sn + strLetter
         If Len(sn) < 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
        If Len(sn) > 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
Next z

End Sub
Public Sub Gen3Chr_NoVowels(HowManySNs As Integer, list As ListBox)
    Dim AlphaNumericString As String, AlphaString As String
    Dim strLetter As String, sn As String, RandomTime As String
    Dim rndX As Integer, rndY As Integer, MakinSNs As Integer, i As Long
    Dim numericstring As String
Dim spacestring As String
    Randomize
    RandomTime = Int(10 * Rnd)
    If RandomTime = 10 Then RandomTime = 9
    
    AlphaNumericString = "1234567890bcdfghjklmnpqrstvwxzbzqr62"
    AlphaString = "bcdfghjklmnpqrstvwxz"
    numericstring = "123456789"
Dim z As Long
For z = 1 To HowManySNs
        sn = ""
        
        For i = 0 To RandomTime
            rndX = Int((30 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndX, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((30 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaNumericString, rndY, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((30 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaNumericString, rndY, 1)
        sn = sn + strLetter
        If Len(sn) < 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
        If Len(sn) > 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
Next z

End Sub
Public Sub Gen3Chr_VowelsOnly(HowManySNs As Integer, list As ListBox)
    Dim AlphaNumericString As String, AlphaString As String
    Dim strLetter As String, sn As String, RandomTime As String
    Dim rndX As Integer, rndY As Integer, MakinSNs As Integer, i As Long
    Dim numericstring As String
Dim spacestring As String
    Randomize
    RandomTime = Int(10 * Rnd)
    If RandomTime = 10 Then RandomTime = 9
    AlphaString = "aeiouy"

Dim z As Long
For z = 1 To HowManySNs
        sn = ""
        
        For i = 0 To RandomTime
            rndX = Int((6 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndX, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((6 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndY, 1)
        sn = sn + strLetter
        For i = 0 To HowManySNs
            rndY = Int((6 - 1 + 1) * Rnd + 1)
        Next i
        strLetter = Mid(AlphaString, rndY, 1)
        sn = sn + strLetter
         If Len(sn) < 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
        If Len(sn) > 3 Then
        Else
        list.AddItem sn
        MakinSNs = MakinSNs + 1
        End If
Next z

End Sub


Public Function GetUser() As String
    Dim AOL As Long, MDI As Long, welcome As Long
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
Public Sub ChatSend5(Chat As String)
    Dim room As Long, AORich As Long, AORich2 As Long
    room& = FindRoom&
    AORich& = FindWindowEx(room, 0&, "RICHCNTL", vbNullString)
    AORich2& = FindWindowEx(room, AORich, "RICHCNTL", vbNullString)
    Call SendMessageByString(AORich2, WM_SETTEXT, 0&, Chat$)
    Call SendMessageLong(AORich2, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub ChatSend7(Chat As String)
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
Dim room As Long
room& = FindRoom&
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(room&, 0&, "aol child", vbNullString)
RICHCNTL = FindWindowEx(room&, 0&, "richcntl", vbNullString)
Call SendMessageByString(RICHCNTL, WM_SETTEXT, 0&, Chat$)
Call SendMessageLong(RICHCNTL, WM_CHAR, ENTER_KEY, 0&)
End Sub
Public Sub ClickToolbar(IconNumber&, letter$)
Dim aolframe As Long
Dim Menu As Long
Dim clickToolbar1 As Long
Dim ClickToolbar2 As Long
Dim aolicon As Long
Dim Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
ClickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
aolicon = FindWindowEx(ClickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
aolicon = FindWindowEx(ClickToolbar2, aolicon, "_AOL_Icon", vbNullString)
Next Count
Pause (0.1)
Call PostMessage(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Do
DoEvents
Menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(Menu)
Loop Until found <> 0
letter = Asc(letter)
Call PostMessage(Menu, WM_CHAR, letter, 0&)
End Sub
Public Sub ClickToolbar2(IconNumber&, letter$, letter2$)
Dim aolframe As Long
Dim Menu As Long
Dim clickToolbar1 As Long
Dim ClickToolbar2 As Long
Dim aolicon As Long
Dim Count As Long
Dim found As Long
aolframe = FindWindow("aol frame25", vbNullString)
clickToolbar1 = FindWindowEx(aolframe, 0&, "AOL Toolbar", vbNullString)
ClickToolbar2 = FindWindowEx(clickToolbar1, 0&, "_AOL_Toolbar", vbNullString)
aolicon = FindWindowEx(ClickToolbar2, 0&, "_AOL_Icon", vbNullString)
For Count = 1 To IconNumber
aolicon = FindWindowEx(ClickToolbar2, aolicon, "_AOL_Icon", vbNullString)
Next Count
Call PostMessage(aolicon, WM_LBUTTONDOWN, 0&, 0&)
Call PostMessage(aolicon, WM_LBUTTONUP, 0&, 0&)
Do
DoEvents
Menu = FindWindow("#32768", vbNullString)
found = IsWindowVisible(Menu)
Loop Until found <> 0
letter = Asc(letter)
letter2 = Asc(letter2)
Call PostMessage(Menu, WM_CHAR, letter, 0&)
Call PostMessage(Menu, WM_CHAR, letter2, 0&)
End Sub
Public Sub Keyword5(KW As String)
    Dim AOL As Long, tool As Long, Toolbar As Long
    Dim Combo As Long, EditWin As Long
    AOL& = FindWindow("AOL Frame25", vbNullString)
    tool& = FindWindowEx(AOL&, 0&, "AOL Toolbar", vbNullString)
    Toolbar& = FindWindowEx(tool&, 0&, "_AOL_Toolbar", vbNullString)
    Combo& = FindWindowEx(Toolbar&, 0&, "_AOL_Combobox", vbNullString)
    EditWin& = FindWindowEx(Combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(EditWin&, WM_SETTEXT, 0&, KW$)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(EditWin&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub Keyword7(KW As String)
Dim aolframe As Long, aoltoolbar As Long, aolcombobox As Long
Dim editx As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolcombobox = FindWindowEx(aoltoolbar, 0&, "_aol_combobox", vbNullString)
editx = FindWindowEx(aolcombobox, 0&, "edit", vbNullString)
Call SendMessageByString(editx, WM_SETTEXT, 0&, KW$)
Call SendMessageLong(editx&, WM_CHAR, VK_SPACE, 0&)
Call SendMessageLong(editx&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Sub Chat_RoomListings()
Call Keyword7("chat room listings")
End Sub
Public Sub Idler()
'This is for anyone who has no idea about vb
'This makes one of those idlers in the chatroom to keep
'   you online


'What you need on your form:
    '2 command buttons. 1 named start, 1 named stop
    '1 textbox
    '1 timer. set the interval to 60000
    '3 labels. 1 named minute, 1 named hour, and 1 named day

'*You will also need something that chatsends!!!*

'In the command button start put:
    'start.enabled=false
    'stop.enabled=true
    'timer1.Enabled = True

'In the command button stop put:
    'start.enabled=true
    'stop.enabled=false
    'timer1.enabled=false
    'Hour.Caption = "0"
    'Minute.Caption = "0"
    'day.caption="0"

'In the timer put:
    'minute = minute + 1
        'if minute > 60 then
            'Minute = "0"
            'hour = hour + 1
        'End If
        'if hour > 24 then
            'hour = 0
            'day = day + 1
        'End If
    
    'chatsend ("i have been idle for " & day & "days, " & hour & "hours, and " & minute & "minutes!")
    'chatsend (" reason i am idle:" & text1)
'----------------------------
    'THERE you have it. An aol idler
End Sub
Public Function AOLOnline() As Boolean
If GetUser = "" Then
AOLOnline = False
Else
AOLOnline = True
End If
End Function
Public Function Mp3_FileSize(file As String) As String
Dim LSize As String
If file = "" Then
Mp3_FileSize = ""
Exit Function
End If
LSize = FileLen(file)
Mp3_FileSize = LSize
End Function
Public Function Mp3_Comment(FileName As String)
Dim strp As String
If TagExists(FileName) = True Then
strp$ = Mp3Info.sComment
GoTo sTrim
sTrim:
If Right(strp, 1) = " " Then
strp$ = Left$(strp$, Len(strp$) - 1)
Mp3_Comment = strp$
GoTo sTrim
Else
Mp3_Comment = strp$
End If
End If
End Function
Sub File_Rename(file$, NewName$)
Dim NoFreeze
    Name file$ As NewName$
    NoFreeze = DoEvents()
End Sub
Public Function File_GetAttributes(TheFile As String) As Integer
    Dim SafeFile As String
    SafeFile$ = Dir(TheFile$)
    If SafeFile$ <> "" Then
        File_GetAttributes% = GetAttr(TheFile$)
    End If
End Function
Public Function AOLVersion() As String
'Gets the aol versions of aols 3-8
'very nice
Dim AOL As Long, hWndMenu As Long, MenuCount As Long, hWndSubMenu As Long
Dim SubMenuCount As Long, menuid As Long, MenuBuffer As String, TheMenuString As Long
Dim LoopCount As Integer, ExtractText As Integer

AOL& = FindWindow("AOL Frame25", vbNullString)
hWndMenu& = GetMenu(AOL&)
MenuCount& = GetMenuItemCount(hWndMenu&)

If MenuCount& = 8 Then
AOLVersion = "3.0"
Exit Function
End If
If MenuCount& = 7 Then
AOLVersion = "2.5"
Exit Function
End If

For LoopCount% = 0 To MenuCount& - 1
hWndSubMenu& = GetSubMenu(hWndMenu&, LoopCount%)
SubMenuCount& = GetMenuItemCount(hWndSubMenu&)

For ExtractText% = 0 To SubMenuCount& - 1
menuid& = GetMenuItemID(hWndSubMenu&, ExtractText%)
MenuBuffer$ = String(100, " ")
TheMenuString& = GetMenuString(hWndSubMenu&, menuid&, MenuBuffer$, 100, 1)
If InStr(LCase(MenuBuffer$), "aol 8.0") Then
AOLVersion = "8.0"
Exit Function
End If
If InStr(LCase(MenuBuffer$), "aol 7.0") Then
AOLVersion = "7.0"
Exit Function
End If
If InStr(LCase(MenuBuffer$), "aol 6.0") Then
AOLVersion = "6.0"
Exit Function
End If
If InStr(LCase(MenuBuffer$), "aol 5.0") Then
AOLVersion = "5.0"
Exit Function
End If
If InStr(LCase(MenuBuffer$), "aol 4.0") Then
AOLVersion = "4.0"
Exit Function
End If
Next ExtractText%
Next LoopCount%

AOLVersion = "n/a"
End Function
Public Function Mp3_Album(FileName As String)
Dim strp As String
If TagExists(FileName) = True Then
strp$ = Mp3Info.sAlbum
GoTo sTrim
sTrim:
If Right(strp, 1) = " " Then
strp$ = Left$(strp$, Len(strp$) - 1)
Mp3_Album = strp$
GoTo sTrim
Else
Mp3_Album = strp$
End If
End If
End Function
Public Function Mp3_Artist(FileName As String)
Dim strp As String
If TagExists(FileName) = True Then
strp$ = Mp3Info.sArtist
GoTo sTrim
sTrim:
If Right(strp, 1) = " " Then
strp$ = Left$(strp$, Len(strp$) - 1)
Mp3_Artist = strp$
GoTo sTrim
Else
Mp3_Artist = strp$
End If
End If
End Function
Public Function Mp3_Title(FileName As String)
Dim strp As String
If TagExists(FileName) = True Then
strp$ = Mp3Info.sTitle
GoTo sTrim
sTrim:
If Right(strp, 1) = " " Then
strp$ = Left$(strp$, Len(strp$) - 1)
Mp3_Title = strp$
GoTo sTrim
Else
Mp3_Title = strp$
End If
End If
End Function


Public Function FindChildByClass(ByVal hParent As Long, ByVal sClassName As String, Optional ByVal nIndex) As Long
   Dim hChild As Long
   Dim i As Integer

   If IsMissing(nIndex) Then
      nIndex = 1
   ElseIf nIndex < 1 Then
      Exit Function
   End If
   hChild = GetWindow(hParent, GW_CHILD)
   While i < nIndex And hChild
      If GetWindowClassName(hChild) = sClassName Then
         i = i + 1
      End If
      
      If i < nIndex Then
         hChild = GetWindow(hChild, GW_HWNDNEXT)
      End If
   Wend
   FindChildByClass = hChild
   Exit Function
End Function
Public Function FindRoomFull()
Dim X As Long
X = FindWindow("#32770", vbNullString)
Call SendMessageLong(X, WM_CLOSE, 0&, 0&)
End Function
Public Function Chat_LeaveRoom()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
Dim room As Long
room& = FindRoom&
Call SendMessageLong(room&, WM_CLOSE, 0&, 0&)
End Function

Public Function Chat_IgnoreByIndex(Index As Long)
'Ignores by a certain number
';]

Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
Dim room As Long
Dim rList As Long
Dim aolicon As Long
Dim AOLIcon2 As Long
room& = FindRoom&
rList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
Call SendMessage(rList&, LB_SETCURSEL, Index&, 0&)
aolicon& = FindWindowEx(room&, 0&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, aolicon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
Call ClickIcon(AOLIcon2&)
End Function

Public Function Chat_EjectByIndex(Index As Long)
'ejects by a certain number
';]
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
Dim room As Long
Dim rList As Long
Dim aolicon As Long
Dim AOLIcon2 As Long

room& = FindRoom&
rList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
Call SendMessage(rList&, LB_SETCURSEL, Index&, 0&)
aolicon& = FindWindowEx(room&, 0&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, aolicon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
Call ClickIcon(AOLIcon2&)
End Function
  Public Function ListCount(ListBox As Long) As Long
    ListCount& = SendMessageLong(ListBox&, LB_GETCOUNT, 0&, 0&)
End Function
Public Sub Chat_EjectByName(strUser As String, blnPartial As Boolean)

On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, Screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lngCheckBox As Long, lngChatUserInfo As Long
    
    strUser = LCase(strUser)
    rList = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To ListCount(rList) - 1
           Screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 28
                Call ReadProcessMemory(mThread&, itmHold&, Screenname$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4)
            psnHold& = psnHold& + 6
            Screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, Screenname$, Len(Screenname$), rBytes&)
            Screenname$ = LCase(Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1))
            
            If (blnPartial = True And InStr(Screenname, strUser)) Or (blnPartial = False And Screenname = strUser) Then
                Call Chat_EjectByIndex(Index&)
          Call Window_Close(lngChatUserInfo)
                Call CloseHandle(mThread)
                Exit Sub
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub Chat_IgnoreByName(strUser As String, blnPartial As Boolean)
On Error Resume Next
    
    Dim cProcess As Long, itmHold As Long, Screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lngCheckBox As Long, lngChatUserInfo As Long
    
    strUser = LCase(strUser)
    rList = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To ListCount(rList) - 1
           Screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + 28
                Call ReadProcessMemory(mThread&, itmHold&, Screenname$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4)
            psnHold& = psnHold& + 6
            Screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, Screenname$, Len(Screenname$), rBytes&)
            Screenname$ = LCase(Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1))
            
            If (blnPartial = True And InStr(Screenname, strUser)) Or (blnPartial = False And Screenname = strUser) Then
                Call Chat_IgnoreByIndex(Index&)
                Call CloseHandle(mThread)
                Exit Sub
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub
Public Sub Chat_InfoByName(strUser As String, blnPartial As Boolean)
On Error Resume Next
    Dim cProcess As Long, itmHold As Long, Screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long
    Dim lngCheckBox As Long, lngChatUserInfo As Long
    strUser = LCase(strUser)
    rList = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
     sThread& = GetWindowThreadProcessId(rList, cProcess&)
 mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
If mThread& Then
For Index& = 0 To ListCount(rList) - 1
 Screenname$ = String$(4, vbNullChar)
itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
itmHold& = itmHold& + 28
Call ReadProcessMemory(mThread&, itmHold&, Screenname$, 4, rBytes)
Call CopyMemory(psnHold&, ByVal Screenname$, 4)
psnHold& = psnHold& + 6
Screenname$ = String$(16, vbNullChar)
Call ReadProcessMemory(mThread&, psnHold&, Screenname$, Len(Screenname$), rBytes&)
Screenname$ = LCase(Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1))
If (blnPartial = True And InStr(Screenname, strUser)) Or (blnPartial = False And Screenname = strUser) Then
Call Chat_InfoByIndex(Index&)
Call CloseHandle(mThread)
                Exit Sub
            End If
            
Next Index&
Call CloseHandle(mThread)
    End If
End Sub

Public Function Chat_InfoByIndex(Index As Long)
'Info by a certain number
';]
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
Dim room As Long
Dim rList As Long
Dim aolicon As Long
Dim AOLIcon2 As Long

room& = FindRoom&
rList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
Index& = Index&
Call SendMessage(rList&, LB_SETCURSEL, Index&, 0&)
aolicon& = FindWindowEx(room&, 0&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, aolicon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon2&, WM_KEYUP, VK_SPACE, 0&)
End Function
Public Function Chat_CloseRoomAll()
'Info by a certain number
';]
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim RICHCNTL As Long
Dim room As Long
Dim rList As Long
Dim aolicon As Long
Dim AOLIcon2 As Long

room& = FindRoom&

aolicon& = FindWindowEx(room&, 0&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, aolicon&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
AOLIcon2& = FindWindowEx(room&, AOLIcon2&, "_AOL_Icon", vbNullString)
Call SendMessageLong(AOLIcon2&, WM_LBUTTONDOWN, 0&, 0&)
Call SendMessageLong(AOLIcon2&, WM_KEYUP, VK_SPACE, 0&)
Dim aolmodal As Long

Do
DoEvents
aolmodal = FindWindow("_aol_modal", vbNullString)
Loop Until aolmodal& <> 0&
AOLIcon2& = FindWindowEx(aolmodal&, 0&, "_AOL_Icon", vbNullString)
Call ClickIcon(AOLIcon2)
End Function

Public Sub Window_StartMenu()
Dim shelltraywnd As Long, button As Long
Dim i As Long
Dim toolbar1 As Long
Dim MenuBlah As Long
shelltraywnd = FindWindow("shell_traywnd", vbNullString)
button = FindWindowEx(shelltraywnd, 0&, "button", vbNullString)
    Call SendMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(button&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Sub Window_Close(Window As Long)
    Call PostMessage(Window&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub Browser_GoBack()
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub BuddyList_Chat(people As String, msg As String, room As String)
'just put like    john,mary,sue or whoeever as ur people
Dim aoledit As Long

Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Buddy List")
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)

Do
DoEvents
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Buddy Chat")
Loop Until aolchild& <> 0&

aoledit = FindWindowEx(aolchild, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, "")
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, people$)
aoledit = FindWindowEx(aolchild, aoledit&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, msg$)
aoledit = FindWindowEx(aolchild, aoledit&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, room$)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Buddy Chat")
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub BuddyList_Setup()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Buddy List")
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub BuddyList_AddBuddy(person As String)
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim buddylist1 As Long
Dim aolicon As Long
Call BuddyList_Setup

Do
DoEvents
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Buddy List Setup")
Pause (0.9)
Loop Until aolchild& <> 0&

aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)

Do
DoEvents
Dim aolmodal As Long
aolframe = FindWindow("aol frame25", vbNullString)
aolmodal = FindWindow("_aol_modal", "Add New Buddy")
Loop Until aolmodal& <> 0&

Dim aoledit As Long
aoledit = FindWindowEx(aolmodal, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, person$)
aolicon = FindWindowEx(aolmodal, 0&, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)

Call SendMessageLong(aolmodal&, WM_CLOSE, 0&, 0&)
Call FindRoomFull
Call SendMessageLong(aolchild&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub BuddyList_AddGroup(group As String)
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim buddylist1 As Long
Dim aolicon As Long
Call BuddyList_Setup

Do
DoEvents
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Buddy List Setup")
Pause (0.9)
Loop Until aolchild& <> 0&

aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Pause (0.2)
Call ClickIcon(aolicon)

Do
DoEvents
Dim aolmodal As Long
aolframe = FindWindow("aol frame25", vbNullString)
aolmodal = FindWindow("_aol_modal", "Add New Group")
Loop Until aolmodal& <> 0&

Dim aoledit As Long
aoledit = FindWindowEx(aolmodal, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, group$)
aolicon = FindWindowEx(aolmodal, 0&, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)

Call SendMessageLong(aolmodal&, WM_CLOSE, 0&, 0&)
Call FindRoomFull
Call SendMessageLong(aolchild&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub BuddyList_Preferences()
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
Call BuddyList_Setup

Do
DoEvents
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Buddy List Setup")
Pause (0.9)
Loop Until aolchild& <> 0&

aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub Browser_GoFoward()
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub Browser_Stop()
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub Browser_Go()
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub Browser_Search()
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub Browser_Refresh()
Dim aolframe As Long, aoltoolbar As Long, aolicon As Long
aolframe = FindWindow("aol frame25", vbNullString)
aoltoolbar = FindWindowEx(aolframe, 0&, "aol toolbar", vbNullString)
aoltoolbar = FindWindowEx(aoltoolbar, 0&, "_aol_toolbar", vbNullString)
aolicon = FindWindowEx(aoltoolbar, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aoltoolbar, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
End Sub
Public Sub ClickIcon(icon As Long)
    Call SendMessageLong(icon, WM_LBUTTONDOWN, 0&, 0&)
    Call SendMessageLong(icon, WM_KEYUP, VK_SPACE, 0&)
End Sub

Public Sub Chat_Addroom7(TheList As ListBox, AddUser As Boolean)
'
    On Error Resume Next
    Dim cProcess As Long, itmHold As Long, Screenname As String
    Dim psnHold As Long, rBytes As Long, Index As Long, room As Long
    Dim rList As Long, sThread As Long, mThread As Long, itmNum As Long
    room& = FindRoom&
    If room& = 0& Then Exit Sub
    itmNum& = 28
Top:
    rList& = FindWindowEx(FindRoom&, 0&, "_AOL_Listbox", vbNullString)
    sThread& = GetWindowThreadProcessId(rList, cProcess&)
    mThread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cProcess&)
    If mThread& Then
        For Index& = 0 To SendMessage(rList, LB_GETCOUNT, 0, 0) - 1
            Screenname$ = String$(4, vbNullChar)
            itmHold& = SendMessage(rList, LB_GETITEMDATA, ByVal CLng(Index&), ByVal 0&)
            itmHold& = itmHold& + itmNum&
            Call ReadProcessMemory(mThread&, itmHold&, Screenname$, 4, rBytes)
            Call CopyMemory(psnHold&, ByVal Screenname$, 4)
            psnHold& = psnHold& + 6
            Screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mThread&, psnHold&, Screenname$, Len(Screenname$), rBytes&)
            Screenname$ = Left$(Screenname$, InStr(Screenname$, vbNullChar) - 1)
            'This is used because if you use the old way it returns either
            'a blank or just a 'p'
            If Screenname$ <> GetUser Or AddUser = True Then
                TheList.AddItem Screenname$
            End If
        Next Index&
        Call CloseHandle(mThread)
    End If
End Sub
Sub String_Save(txtSave As String, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub
Public Function String_Replace(MyString As String, ToFind As String, ReplaceWith As String) As String
    Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    String_Replace$ = NewString$
End Function

Public Function String_Reverse(MyString As String) As String
    Dim TempString As String, StringLength As Long
    Dim Count As Long, NextChr As String, NewString As String
    TempString$ = MyString$
    StringLength& = Len(TempString$)
    Do While Count& <= StringLength&
        Count& = Count& + 1
        NextChr$ = Mid$(TempString$, Count&, 1)
        NewString$ = NextChr$ & NewString$
    Loop
    String_Reverse$ = NewString$
End Function

Public Sub button(mButton As Long)
    Call SendMessage(mButton&, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessage(mButton&, WM_KEYUP, VK_SPACE, 0&)
End Sub
Public Function String_Switch(MyString As String, String1 As String, String2 As String) As String
    Dim TempString As String, Spot1 As Long, Spot2 As Long
    Dim Spot As Long, ToFind As String, ReplaceWith As String
    Dim NewSpot As Long, LeftString As String, RightString As String
    Dim NewString As String
    If Len(String2) > Len(String1) Then
        TempString$ = String1$
        String1$ = String2$
        String2$ = TempString$
    End If
    Spot1& = InStr(MyString$, String1$)
    Spot2& = InStr(MyString$, String2$)
    If Spot1& = 0& And Spot2& = 0& Then
        String_Switch$ = MyString$
        Exit Function
    End If
    If Spot1& < Spot2& Or Spot2& = 0 Or Len(String1$) = Len(String2$) Then
        If Spot1& > 0 Then
            Spot& = Spot1&
            ToFind$ = String1$
            ReplaceWith$ = String2$
        End If
    End If
    If Spot2& < Spot1& Or Spot1& = 0& Then
        If Spot2& > 0& Then
            Spot& = Spot2&
            ToFind$ = String2$
            ReplaceWith$ = String1$
        End If
    End If
    If Spot1& = 0& And Spot2& = 0& Then
        String_Switch$ = MyString$
        Exit Function
    End If
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString$ = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot + Len(ReplaceWith$) - Len(ToFind$) + 1
        If Spot& <> 0& Then
            Spot1& = InStr(Spot&, MyString$, String1$)
            Spot2& = InStr(Spot&, MyString$, String2$)
        End If
        If Spot1& = 0& And Spot2& = 0& Then
            String_Switch$ = MyString$
            Exit Function
        End If
        If Spot1& < Spot2& Or Spot2& = 0& Or Len(String1$) = Len(String2$) Then
            If Spot1& > 0& Then
                Spot& = Spot1&
                ToFind$ = String1$
                ReplaceWith$ = String2$
            End If
        End If
        If Spot2& < Spot1& Or Spot1& = 0& Then
            If Spot2& > 0& Then
                Spot& = Spot2&
                ToFind$ = String2$
                ReplaceWith$ = String1$
            End If
        End If
        If Spot1& = 0& And Spot2& = 0& Then
            Spot& = 0&
        End If
        If Spot& > 0& Then
            NewSpot& = InStr(Spot&, MyString$, ToFind$)
        Else
            NewSpot& = Spot&
        End If
    Loop Until NewSpot& < 1&
    String_Switch$ = NewString$
End Function

Public Sub List_KillDupes(TheList As ListBox)
Dim Count&, Count2&, Count3&
If TheList.ListCount = 0 Then Exit Sub
For Count& = 0 To TheList.ListCount - 1
DoEvents
For Count2& = Count& + 1 To TheList.ListCount - 1
DoEvents
If TheList.list(Count&) = TheList.list(Count2&) Then
TheList.RemoveItem (Count2&)

End If
Next Count2
Next Count&
End Sub

Public Sub Chat_Lobby7()
Call ClickToolbar("3", "N")
End Sub

Public Sub People_Connection7()
Call ClickToolbar("3", "C")
End Sub
Public Sub People_SendIM()
Do
Call ClickToolbar("3", "i")
Loop Until FindIM <> 0
End Sub
Public Sub IM8(person As String, message As String)
    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim SendButton As Long, OK As Long, button As Long
    Dim i As Long

    AOL& = FindWindow("AOL Frame25", vbNullString)
    MDI& = FindWindowEx(AOL&, 0&, "MDIClient", vbNullString)
    Call Keyword7("aol://9293:" & person$)
    Do
        DoEvents
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
        Rich& = FindWindowEx(IM&, 0&, "RICHCNTL", vbNullString)
    
    SendButton& = 0&
    For i = 1 To 12
        SendButton& = FindWindowEx(IM&, SendButton&, "_aol_icon", vbNullString)
    Next i
    
    Loop Until IM& <> 0& And Rich& <> 0& And SendButton& <> 0&
    Call SendMessageByString(Rich&, WM_SETTEXT, 0&, message$)
    Call ClickIcon(SendButton&)

    Do
        DoEvents
        OK& = FindWindow("#32770", "America Online")
        IM& = FindWindowEx(MDI&, 0&, "AOL Child", "Send Instant Message")
    Loop Until OK& <> 0& Or IM& = 0&
    If OK& <> 0& Then
        button& = FindWindowEx(OK&, 0&, "Button", vbNullString)
        Call PostMessage(button&, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(button&, WM_KEYUP, VK_SPACE, 0&)
        Call PostMessage(IM&, WM_CLOSE, 0&, 0&)
    End If
End Sub

Public Function IM8_LastMessage() As String


    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim SendButton As Long, OK As Long, button As Long
    Dim i As Long
    Dim MsgString As String
    Dim NewSpot As Long
    Dim Spot As Long
    Dim strMsgs As String, intLoc As Integer, strChar As String
    
    Rich& = FindWindowEx(FindIM&, 0&, "RICHCNTLREADONLY", vbNullString)

    MsgString$ = GetText(Rich&)
    NewSpot& = InStr(MsgString$, Chr(9))
    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
   MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
   IM8_LastMessage$ = Left(MsgString$, Len(MsgString$) - 1)
 
 
 strMsgs = IM8_LastMessage$
    'find the first enter
        intLoc = InStrRev(strMsgs, Chr(13))
        strMsgs = Mid(strMsgs, intLoc + 1)
    'find the ":"
        intLoc = InStr(strMsgs, " ")
    'find the first character of the actual message
        Do: DoEvents
            intLoc = intLoc + 1
            strChar = Mid(strMsgs, intLoc, 1)
        Loop Until strChar <> " "
    'return last message
    IM8_LastMessage$ = Mid(strMsgs, intLoc)

 
 
 
End Function




Public Function Chat_LastLine() As String


    Dim AOL As Long, MDI As Long, IM As Long, Rich As Long
    Dim SendButton As Long, OK As Long, button As Long
    Dim i As Long
    Dim MsgString As String
    Dim NewSpot As Long
    Dim Spot As Long
    Dim strMsgs As String, intLoc As Integer, strChar As String
    
    Rich& = FindWindowEx(FindRoom&, 0&, "RICHCNTLREADONLY", vbNullString)
    MsgString$ = GetText(Rich&)
    NewSpot& = InStr(MsgString$, Chr(9))

    Do
        Spot& = NewSpot&
        NewSpot& = InStr(Spot& + 1, MsgString$, Chr(9))
    Loop Until NewSpot& <= 0&
    MsgString$ = Right(MsgString$, Len(MsgString$) - Spot& - 1)
    Chat_LastLine$ = Left(MsgString$, Len(MsgString$) - 1)
    strMsgs = Chat_LastLine$
    intLoc = InStrRev(strMsgs, Chr(13))
    strMsgs = Mid(strMsgs, intLoc + 1)
    intLoc = InStr(strMsgs, "")
Do: DoEvents
    intLoc = intLoc + 1
    strChar = Mid(strMsgs, intLoc, 1)
    Loop Until strChar <> " "
    Chat_LastLine$ = Mid(strMsgs, intLoc)
 
End Function
Public Sub Settings_Favorites(name As String, url As String)
Call ClickToolbar("11", "f")
End Sub
Public Sub Settings_AddFavorite(name As String, url As String)
Call ClickToolbar("11", "f")
Pause (0.2)
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
Dim aoledit As Long
Dim aolmodal As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
Do
DoEvents
aolchild = FindWindowEx(mdiclient, 0&, "aol child", GetUser & "'s Favorite Places")
Loop Until aolchild& <> 0&
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)

Do
DoEvents
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Add New Folder/Favorite Place")
Loop Until aolchild& <> 0&
aoledit = FindWindowEx(aolchild, 0&, "_aol_edit", vbNullString)
aoledit = FindWindowEx(aolchild, aoledit, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, name$)
aoledit = FindWindowEx(aolchild, aoledit, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit, WM_SETTEXT, 0&, url$)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
aolicon = FindWindowEx(aolchild, aolicon, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)

aolmodal = FindWindow("_aol_modal", vbNullString)
If aolmodal& <> 0& Then
Call Window_Close(aolmodal)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", GetUser & "'s Favorite Places")
Call Window_Close(aolchild)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Add New Folder/Favorite Place")
Call Window_Close(aolchild)
Else
Pause (0.2)
Call FindRoomFull
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Add New Folder/Favorite Place")
Call Window_Close(aolchild)
aolchild = FindWindowEx(mdiclient, 0&, "aol child", GetUser & "'s Favorite Places")
Call Window_Close(aolchild)
End If
End Sub

Public Function GetText(WindowHandle As Long) As String
    Dim buffer As String, TextLength As Long
    TextLength& = SendMessage(WindowHandle&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(TextLength&, 0&)
    Call SendMessageByString(WindowHandle&, WM_GETTEXT, TextLength& + 1, buffer$)
    GetText$ = buffer$
End Function
Public Sub IM8_Off()
Call IM8("$im_Off", "yay")
End Sub
Public Sub IM8_On()
Call IM8("$im_On", "yay")
End Sub
Public Sub People_FindChat()
Call ClickToolbar("3", "A")
End Sub
Public Sub People_GetProfile()
Call ClickToolbar("3", "G")
End Sub
Public Sub People_Locate()
Call ClickToolbar("3", "L")
End Sub
Public Function People_LocatePerson(sn As String) As String
' MsgBox People_LocatePerson("barghh")
'
Call ClickToolbar("3", "L")
Dim aolframe As Long, mdiclient As Long, aolchild As Long
Dim aolicon As Long
Dim richcntlstatic As Long
aolframe = FindWindow("aol frame25", vbNullString)
mdiclient = FindWindowEx(aolframe, 0&, "mdiclient", vbNullString)
Do
DoEvents
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Locate Member Online")
Loop Until aolchild& <> 0&
Dim where As String
Dim aoledit As Long
aoledit = FindWindowEx(aolchild, 0&, "_aol_edit", vbNullString)
Call SendMessageByString(aoledit&, WM_SETTEXT, 0&, sn$)
aolicon = FindWindowEx(aolchild, 0&, "_aol_icon", vbNullString)
Call ClickIcon(aolicon)
Call Window_Close(aolchild)
Do
DoEvents
aolchild = FindWindowEx(mdiclient, 0&, "aol child", "Info")
Loop Until aolchild& <> 0&
Pause (1)
richcntlstatic = FindWindowEx(aolchild, 0&, "richcntlstatic", vbNullString)
People_LocatePerson$ = GetText(richcntlstatic&)
If People_LocatePerson$ = "I am online." Then
People_LocatePerson$ = sn$ & " is online."
End If
If People_LocatePerson$ = "I am offline." Then
People_LocatePerson$ = sn$ & " is offline."
End If
Call Window_Close(aolchild)
End Function
Public Sub People_BuddyList()
Call ClickToolbar("3", "B")
End Sub
Public Sub People_MemberDirectory()
Call ClickToolbar("3", "D")
End Sub
Public Sub Mail_Read()
Call ClickToolbar("0", "r")
End Sub
Public Sub Mail_ReadOld()
Call ClickToolbar2("0", "r", "O")
End Sub
Public Sub Mail_ReadSent()
Call ClickToolbar2("0", "r", "S")
End Sub
Public Sub Mail_ReadNew()
Call ClickToolbar2("0", "r", "N")
End Sub
Public Sub Mail_Write()
Call ClickToolbar("0", "w")
End Sub
Public Sub Mail_RecentDeleted()
Call ClickToolbar("0", "D")
End Sub
Public Sub Mail_AddressBook()
Call ClickToolbar("0", "a")
End Sub
Public Sub Mail_Deleted()
Call ClickToolbar("0", "d")
End Sub
Public Sub Mail_Prefs()
Call ClickToolbar("0", "p")
End Sub
Public Sub Settings_MyAol()
Call ClickToolbar("9", "m")
End Sub
Public Sub Settings_EditProfile()
Call ClickToolbar("9", "l")
End Sub
Public Sub Window_Hide(hwnd As Long)
    Call ShowWindow(hwnd&, SW_HIDE)
End Sub
Public Sub Window_Show(hwnd As Long)
    Call ShowWindow(hwnd&, SW_SHOW)
End Sub
