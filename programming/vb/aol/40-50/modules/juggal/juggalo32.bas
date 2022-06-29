Attribute VB_Name = "juggalo32"
'  |¯|   GpX - GpX@hider.com   |¯|
'   ¯     www.hider.com/gpx    | |      |¯¯¯¯\/¯¯¯¯\
'  |¯||¯||¯|/¯¯¯¯\/¯¯¯¯\/¯¯¯\¯|| |/¯¯¯¯\ ¯¯\ ||_/\ |
'  | || || || /\ || /\ || /\  || || /\ | |¯  |  / /
'  | || \/ || \/ || \/ || \/  || || \/ |  ¯/ | / /
'  | |\____/\__  |\__  |\___/_||_|\____/|¯¯ / |  ¯¯|
'|¯| |6/25/99__/ / __/ /                 ¯¯¯   ¯¯¯¯
'\___/      |___/ |___/beta testers: ~izekial~kid-web~wolph~
'                      contributed:  ~dos~wolph~izekial~neo~

Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function ReadProcessMemory Lib "KERNEL32" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef lpNumberOfBytesWritten As Long) As Long
Public Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName As Any, ByVal lpDefault As String, ByVal lpRetunedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)

Public Const SW_MAXIMIZE = 3
Public Const SW_MINIMIZE = 6
Public Const SW_NORMAL = 1
Public Const SW_HIDE = 0
Public Const SW_SHOW = 5

Public Const EM_REPLACESEL = &HC2
Public Const EM_SETSEL = &HB1

Public Const WM_CHAR = &H102
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_ENABLE = &HA
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

Public Const CB_ADDSTRING = &H143
Public Const CB_DELETESTRING = &H144
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_GETCOUNT = &H146
Public Const CB_RESETCONTENT = &H14B
Public Const CB_INSERTSTRING = &H14A
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETCURSEL = &H147
Public Const CB_GETITEMDATA = &H150
Public Const CB_SELECTSTRING = &H14D
Public Const CB_SETCURSEL = &H14E
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_GETLBTEXT = &H148

Public Const VK_DELETE = &H2E
Public Const VK_DOWN = &H28
Public Const VK_RETURN = &HD
Public Const VK_RIGHT = &H27
Public Const VK_SPACE = &H20
Public Const VK_UP = &H26

Public Const BM_GETCHECK = &HF0
Public Const BM_SETCHECK = &HF1

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
Public Const LB_SETCURSEL = &H186

Public Const PROCESS_READ = &H10
Public Const RIGHTS_REQUIRED = &HF0000

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const Flags = SWP_NOMOVE Or SWP_NOSIZE

Public Type POINTAPI
        x As Long
        Y As Long
End Type

Public Enum MAILTEXT
        mtDATE
        mtSENDER
        mtSUBJECT
        mtALL
End Enum
Public Enum MAILTYPE
        mtFLASH
        mtNEW
        mtOLD
        mtSENT
End Enum
Public Enum SEX
        sMALE
        sFEMALE
        sNEITHER
End Enum
Public Enum INFO
        iSYSTEMINFO
        iERRORMSG
        iCONNECTION
        iPPPINFO
End Enum
Public Enum CHECKVALUE
        cvCHECKED
        cvUNCHECKED
End Enum
Public Enum SPEED
        s9600
        s14400
        s19200
        s28800
        s33600
        s38400
        s57600
        s115200
End Enum
Public Enum NETWORK
        nAOLGLOBALNET
        nAOLNET
        nSPRINTNET
End Enum
Public Enum FORMPOS
        fpCENTER
        fpTOPLEFT
        fpTOPRIGHT
        fpBOTTOMLEFT
        fpBOTTOMRIGHT
End Enum
Public Enum FRMEXIT
        feTOP
        feBOTTOM
        feRIGHT
        feLEFT
        feTOPLEFT
        feTOPRIGHT
        feBOTTOMLEFT
        feBOTTOMRIGHT
End Enum

Global roomrunstop As Boolean
Global roombuststop As Boolean
Global idlebotstop As Boolean

Global roombustcount As Long
Public Sub addaccessnumber(name As String, location As Long, connectusing As Long, timestorepeat As String, phonenumber As String, reachoutsideline As Boolean, reachoutsidelinecode As String, callwaiting As Boolean, callwaitingdisablecode As String, modemspeed As SPEED, networktype As NETWORK)
    'thanks wolph for the idea
    Dim aol As Long, mdi As Long, signonwin As Long, setupbutton As Long
    Dim setupwin As Long, connectwin As Long, connecttabs As Long
    Dim connecttab As Long, addnumberbutton1 As Long, addnumberbutton As Long
    Dim namebox As Long, locationcombo As Long, connectcombo As Long
    Dim repeatbox As Long, numberbox1 As Long, numberbox As Long, okbutton As Long
    Dim outsidecheck As Long, callwaitingcheck As Long, outsidebox As Long
    Dim callwaitingbox As Long, speedcombo As Long, networkcombo As Long
    Dim lspeed As Long, lnetwork As Long, setupwin1 As Long, setupbutton1 As Long
    Dim setupbuttona As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    signonwin& = findsignonwin()
    If signonwin& = 0& Then
        Call opensignonwin
        Do: DoEvents
            signonwin& = findsignonwin()
            setupbutton& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
            setupbuttona& = FindWindowEx(signonwin&, setupbutton&, "_AOL_Icon", vbNullString)
        Loop Until signonwin& <> 0& And setupbuttona& <> 0&
    End If
    setupbutton& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
    setupbuttona& = FindWindowEx(signonwin&, setupbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(setupbuttona&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(setupbuttona&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        setupwin& = FindWindow("_AOL_Modal", "AOL Setup")
        setupbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And setupbutton1& <> 0&
    Call PostMessage(setupbutton1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(setupbutton1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        connectwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Connection Setup")
        connecttabs& = FindWindowEx(connectwin&, 0&, "_AOL_TabControl", vbNullString)
        connecttab& = FindWindowEx(connecttabs&, 0&, "_AOL_TabPage", vbNullString)
        addnumberbutton1& = FindWindowEx(connecttab&, 0&, "_AOL_Icon", vbNullString)
        addnumberbutton1& = FindWindowEx(connecttab&, addnumberbutton1&, "_AOL_Icon", vbNullString)
        addnumberbutton& = FindWindowEx(connecttab&, addnumberbutton1&, "_AOL_Icon", vbNullString)
    Loop Until connectwin& <> 0& And connecttabs& <> 0& And connecttab& <> 0& And addnumberbutton& <> 0&
    Call PostMessage(addnumberbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(addnumberbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        setupwin1& = FindWindowEx(mdi&, 0&, "AOL Child", "AOL Setup")
        namebox& = FindWindowEx(setupwin1&, 0&, "_AOL_Edit", vbNullString)
        locationcombo& = FindWindowEx(setupwin1&, 0&, "_AOL_Combobox", vbNullString)
        connectcombo& = FindWindowEx(setupwin1&, locationcombo&, "_AOL_Combobox", vbNullString)
        repeatbox& = FindWindowEx(setupwin1&, 0&, "_AOL_Spin", vbNullString)
        numberbox1& = FindWindowEx(setupwin1&, namebox&, "_AOL_Edit", vbNullString)
        numberbox1& = FindWindowEx(setupwin1&, numberbox1&, "_AOL_Edit", vbNullString)
        numberbox& = FindWindowEx(setupwin1&, numberbox1&, "_AOL_Edit", vbNullString)
        outsidecheck& = FindWindowEx(setupwin1&, 0&, "_AOL_Checkbox", vbNullString)
        callwaitingcheck& = FindWindowEx(setupwin1&, outsidecheck&, "_AOL_Checkbox", vbNullString)
        outsidebox& = FindWindowEx(setupwin1&, numberbox&, "_AOL_Edit", vbNullString)
        callwaitingbox& = FindWindowEx(setupwin1&, outsidebox&, "_AOL_Edit", vbNullString)
        speedcombo& = FindWindowEx(setupwin1&, connectcombo&, "_AOL_Combobox", vbNullString)
        networkcombo& = FindWindowEx(setupwin1&, speedcombo&, "_AOL_Combobox", vbNullString)
        okbutton& = FindWindowEx(setupwin1&, 0&, "_AOL_Icon", vbNullString)
    Loop Until setupwin1& <> 0& And namebox& <> 0& And locationcombo& <> 0& And connectcombo& <> 0& And repeatbox& <> 0& And numberbox& <> 0& And outsidecheck& <> 0& And callwaitingcheck& <> 0& And outsidebox& <> 0& And callwaitingbox& <> 0& And speedcombo& <> 0& And networkcombo& <> 0& And okbutton& <> 0&
    Call SendMessageByString(namebox&, WM_SETTEXT, 0&, name$)
    Call SendMessageByString(repeatbox&, WM_SETTEXT, 0&, timestorepeat$)
    Call SendMessageByString(numberbox&, WM_SETTEXT, 0&, phonenumber$)
    Call SendMessageByString(outsidebox&, WM_SETTEXT, 0&, reachoutsidelinecode$)
    Call SendMessageByString(callwaitingbox&, WM_SETTEXT, 0&, callwaitingdisablecode$)
    Call PostMessage(outsidecheck&, BM_SETCHECK, reachoutsideline, 0&)
    Call PostMessage(callwaitingcheck&, BM_SETCHECK, callwaiting, 0&)
    Call PostMessage(locationcombo&, CB_SETCURSEL, location&, 0&)
    Call PostMessage(connectcombo&, CB_SETCURSEL, connectusing&, 0&)
    Select Case modemspeed
        Case s9600: lspeed& = 0&
        Case s14400: lspeed& = 1&
        Case s19200: lspeed& = 2&
        Case s28800: lspeed& = 3&
        Case s33600: lspeed& = 4&
        Case s38400: lspeed& = 5&
        Case s57600: lspeed& = 6&
        Case s115200: lspeed& = 7&
    End Select
    Select Case networktype
        Case nAOLGLOBALNET: lnetwork& = 0&
        Case nAOLNET: lnetwork& = 1&
        Case nSPRINTNET: lnetwork& = 2&
    End Select
    Call PostMessage(speedcombo&, CB_SETCURSEL, lspeed&, 0&)
    Call PostMessage(networkcombo&, CB_SETCURSEL, lnetwork&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(connectwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function addallbuddystostring(separator As String) As String
    Dim aol As Long, mdi As Long, buddylistwin As Long, whereat As Long
    Dim setupbutton As Long, setupwin As Long, editbutton1 As Long
    Dim editbutton As Long, grouplist As Long, groupcount As Long
    Dim editwin As Long, groupname As String, snlist As Long
    Dim index As Long, getcount1 As Long, getcount2 As Long, getcount3 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddylistwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddylistwin& = 0& Then
        Call keyword("buddy list")
    Else
        setupbutton& = FindWindowEx(buddylistwin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddylistwin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddylistwin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        If setupwin& = 0& Then setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy List")
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until setupwin& <> 0& And editbutton& <> 0& And grouplist& <> 0&
    groupcount& = SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        groupname$ = getlistitemtext(grouplist&, 0&)
        whereat& = InStr(groupname$, Chr(9))
        groupname$ = Left(groupname$, whereat& - 1)
        editwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & getlistitemtext(grouplist&, 0&))
        snlist& = FindWindowEx(editwin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until editwin& <> 0& And snlist& <> 0&
    For index& = 1 To groupcount&
        groupname$ = getlistitemtext(grouplist&, index& - 1)
        whereat& = InStr(groupname$, Chr(9))
        groupname$ = Left(groupname$, whereat& - 1)
        Do: DoEvents
            editwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & groupname$)
            snlist& = FindWindowEx(editwin&, 0&, "_AOL_Listbox", vbNullString)
        Loop Until editwin& <> 0& And snlist& <> 0&
        Do: DoEvents
            getcount1& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
            pause 0.2
            getcount2& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
            pause 0.2
            getcount3& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
        Loop Until getcount1& = getcount2& And getcount2& = getcount3&
        addallbuddystostring$ = addallbuddystostring$ & separator$ & addlisttostring(snlist&, separator$)
        Call SendMessage(editwin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(grouplist&, LB_SETCURSEL, CLng(index&), 0&)
        Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Next index&
    Call SendMessage(editwin&, WM_CLOSE, 0&, 0&)
    Call SendMessage(setupwin&, WM_CLOSE, 0&, 0&)

End Function

Public Sub addbuddy(screenname As String, group As String)
    Dim aol As Long, mdi As Long, buddywin As Long, setupbutton As Long
    Dim setupwin As Long, grouplist As Long, editbutton As Long, groupindex As Long
    Dim groupwin As Long, addbutton As Long, removebutton As Long, savebutton As Long
    Dim buddylist As Long, addbox1 As Long, buddylistindex As Long, addbox As Long
    Dim grouptext As String, whereat As Long, editbutton1 As Long, errorstatic1 As Long
    Dim errorstatic As Long, screennameindex As Long, statictxt As String
    group$ = LCase(group$)
    group$ = removechar(group$, " ")
    screenname$ = LCase(screenname$)
    screenname$ = removechar(screenname$, " ")
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0& Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        If setupwin& = 0& Then setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy List")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0& And editbutton& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&)
        grouptext$ = getlistitemtext(grouplist&, groupindex&)
        whereat& = InStr(grouptext$, Chr(9))
        grouptext$ = Left(grouptext$, whereat& - 1)
        grouptext$ = LCase(grouptext$)
        grouptext$ = removechar(grouptext$, " ")
        If grouptext$ = group$ Then Exit For
    Next groupindex&
    If groupindex& = -1& Then
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    grouptext$ = getlistitemtext(grouplist&, groupindex&)
    grouptext$ = Left(grouptext$, whereat& - 1)
    Call PostMessage(grouplist&, LB_SETCURSEL, CLng(groupindex&), 0&)
    Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        groupwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & grouptext$)
        addbox1& = FindWindowEx(groupwin&, 0&, "_AOL_Edit", vbNullString)
        addbox& = FindWindowEx(groupwin&, addbox1&, "_AOL_Edit", vbNullString)
        buddylist& = FindWindowEx(groupwin&, 0&, "_AOL_Listbox", vbNullString)
        addbutton& = FindWindowEx(groupwin&, 0&, "_AOL_Icon", vbNullString)
        removebutton& = FindWindowEx(groupwin&, addbutton&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(groupwin&, removebutton&, "_AOL_Icon", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, 0&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
    Loop Until groupwin& <> 0& And addbox& <> 0& And buddylist& <> 0& And addbutton& <> 0& And removebutton& <> 0& And savebutton& <> 0& And errorstatic& <> 0&
    Call waitforlisttoload(buddylist&)
    If getlistitemindex(buddylist&, screenname$, True, True) <> -1& Then
        Call PostMessage(groupwin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    Call SendMessageByString(addbox&, WM_SETTEXT, 0&, screenname$)
    Call PostMessage(addbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(addbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        screennameindex& = getlistitemindex(buddylist&, screenname$, True, True)
        statictxt$ = gettext(errorstatic&)
    Loop Until screennameindex& <> 0& Or statictxt$ <> " "
    If statictxt$ <> " " Then
        Call PostMessage(groupwin&, WM_CLOSE, 0&, 0&)
    ElseIf statictxt$ = " " Then
        Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Call waitforok
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addfavorite(name As String, location As String)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, favwin As Long, newbutton1 As Long, newbutton As Long
    Dim addwin As Long, namebox As Long, locationbox As Long
    Dim okbutton As Long, okbutton1 As Long, namebox1 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 1&)
    Do: DoEvents
        favwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Favorite Places")
        newbutton1& = FindWindowEx(favwin&, 0&, "_AOL_Icon", vbNullString)
        newbutton& = FindWindowEx(favwin&, newbutton1&, "_AOL_Icon", vbNullString)
    Loop Until favwin& <> 0& And newbutton& <> 0&
    Call PostMessage(newbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(newbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        addwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Add New Folder/Favorite Place")
        namebox1& = FindWindowEx(addwin&, 0&, "_AOL_Edit", vbNullString)
        namebox& = FindWindowEx(addwin&, namebox1&, "_AOL_Edit", vbNullString)
        locationbox& = FindWindowEx(addwin&, namebox&, "_AOL_Edit", vbNullString)
        okbutton1& = FindWindowEx(addwin&, 0&, "_AOL_Icon", vbNullString)
        okbutton& = FindWindowEx(addwin&, okbutton1&, "_AOL_Icon", vbNullString)
    Loop Until addwin& <> 0& And namebox& <> 0& And locationbox& <> 0& And okbutton& <> 0&
    Call SendMessageByString(namebox&, WM_SETTEXT, 0&, name$)
    Call SendMessageByString(locationbox&, WM_SETTEXT, 0&, location$)
    Do: DoEvents
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        okbutton& = FindWindowEx(addwin&, okbutton1&, "_AOL_Icon", vbNullString)
    Loop Until okbutton& = 0&
    Call SendMessage(favwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addlisttocontrol(thelist As Long, addto As Control)
    Dim listcount As Long, sthread As Long, mthread As Long, thestring As String
    Dim itmhold As Long, psnhold As Long, cprocess As Long, rbytes As Long
    Dim index As Long
    listcount& = SendMessage(thelist&, LB_GETCOUNT, 0&, 0&)
    If listcount& = 0& Then Exit Sub
    sthread& = GetWindowThreadProcessId(thelist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To listcount& - 1
            thestring$ = String$(4, vbNullChar)
            itmhold& = SendMessage(thelist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, thestring$, 4, rbytes&)
            Call CopyMemory(psnhold&, ByVal thestring$, 4)
            psnhold& = psnhold& + 6
            thestring$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, thestring$, Len(thestring$), rbytes&)
            If InStr(thestring$, vbNullChar) <> 0& Then thestring$ = Left$(thestring$, InStr(thestring$, vbNullChar) - 1)
            addto.AddItem thestring$
        Next index&
        Call CloseHandle(mthread&)
    End If
End Sub
Public Function addlisttostring(thelist As Long, separator As String) As String
    Dim listcount As Long, sthread As Long, mthread As Long, index As Long
    Dim itmhold As Long, cprocess As Long, rbytes As Long, psnhold As Long
    Dim thestring As String
    listcount& = SendMessage(thelist&, LB_GETCOUNT, 0&, 0&)
    If listcount& = 0& Then Exit Function
    sthread& = GetWindowThreadProcessId(thelist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To listcount& - 1
            thestring$ = String$(4, vbNullChar)
            itmhold& = SendMessage(thelist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, thestring$, 4, rbytes&)
            Call CopyMemory(psnhold&, ByVal thestring$, 4)
            psnhold& = psnhold& + 6
            thestring$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, thestring$, Len(thestring$), rbytes&)
            thestring$ = Left$(thestring$, InStr(thestring$, vbNullChar) - 1)
            addlisttostring$ = addlisttostring$ & separator$ & thestring$
        Next index&
        Call CloseHandle(mthread&)
    End If
End Function
Public Sub addlocation(location As String, timestorepeat As String)
    Dim aol As Long, mdi As Long, signonwin As Long, setupbutton As Long
    Dim setupbuttona As Long, setupwin As Long, setupbutton1 As Long
    Dim connectwin As Long, connecttabs As Long, connecttab As Long
    Dim addlocationbutton As Long, setupwin1 As Long, namebox As Long
    Dim repeatbox As Long, okbutton1 As Long, okbutton As Long
    Dim setupwin2 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    signonwin& = findsignonwin()
    If signonwin& = 0& Then
        Call opensignonwin
        Do: DoEvents
            signonwin& = findsignonwin()
            setupbutton& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
            setupbuttona& = FindWindowEx(signonwin&, setupbutton&, "_AOL_Icon", vbNullString)
        Loop Until signonwin& <> 0& And setupbuttona& <> 0&
    End If
    setupbutton& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
    setupbuttona& = FindWindowEx(signonwin&, setupbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(setupbuttona&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(setupbuttona&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        setupwin& = FindWindow("_AOL_Modal", "AOL Setup")
        setupbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And setupbutton1& <> 0&
    Call PostMessage(setupbutton1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(setupbutton1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        connectwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Connection Setup")
        connecttabs& = FindWindowEx(connectwin&, 0&, "_AOL_TabControl", vbNullString)
        connecttab& = FindWindowEx(connecttabs&, 0&, "_AOL_TabPage", vbNullString)
        addlocationbutton& = FindWindowEx(connecttab&, 0&, "_AOL_Icon", vbNullString)
    Loop Until connectwin& <> 0& And connecttabs& <> 0& And connecttab& <> 0& And addlocationbutton& <> 0&
    Call PostMessage(addlocationbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(addlocationbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        setupwin1& = FindWindowEx(mdi&, 0&, "AOL Child", "AOL Setup")
        namebox& = FindWindowEx(setupwin1&, 0&, "_AOL_Edit", vbNullString)
        repeatbox& = FindWindowEx(setupwin1&, 0&, "_AOL_Spin", vbNullString)
        okbutton1& = FindWindowEx(setupwin1&, 0&, "_AOL_Icon", vbNullString)
        okbutton1& = FindWindowEx(setupwin1&, okbutton1&, "_AOL_Icon", vbNullString)
        okbutton& = FindWindowEx(setupwin1&, okbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin1& <> 0& And namebox& <> 0& And repeatbox& <> 0& And okbutton& <> 0&
    Call SendMessageByString(namebox&, WM_SETTEXT, 0&, location$)
    Call SendMessageByString(repeatbox&, WM_SETTEXT, 0&, timestorepeat$)
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        setupwin2& = FindWindow("_AOL_Modal", "AOL Setup")
    Loop Until setupwin2& <> 0&
    Call PostMessage(setupwin2&, WM_CLOSE, 0&, 0&)
    Call PostMessage(connectwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addlocationstocontrol(thecontrol As Control, signonwin As Long)
    Dim index As Long
    For index& = 0& To getlocationcount(signonwin&) - 1&
        thecontrol.AddItem getlocationtext(signonwin&, index&)
    Next index&
End Sub
Public Function addlocationstostring(signonwin As Long, separator As String) As String
    Dim index As Long
    For index& = 0& To getlocationcount(signonwin&) - 1&
        addlocationstostring$ = getlocationtext(signonwin&, index&) & separator$ & addlocationstostring$
    Next index&
End Function
Public Sub addmemberdirtocontrol(addto As Control, searchfor As String, onlyon As Boolean)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long, icon As Long
    Dim dirwin As Long, searchbox As Long, checkbox1 As Long, checkbox As Long
    Dim searchbutton1 As Long, searchbutton As Long, resultswin As Long, snlist As Long
    Dim okwin As Long, okbutton As Long, index As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 9&)
    Do: DoEvents
        dirwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
        searchbox& = FindWindowEx(dirwin&, 0&, "_AOL_Edit", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, 0&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, 0&, "_AOL_Icon", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
        searchbutton& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
    Loop Until dirwin& <> 0& And searchbox& <> 0& And checkbox& <> 0& And searchbutton& <> 0&
    Call SendMessageByString(searchbox&, WM_SETTEXT, 0&, searchfor$)
    If onlyon = True Then
        Call PostMessage(checkbox&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(checkbox&, WM_LBUTTONUP, 0&, 0&)
    End If
    Call PostMessage(searchbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(searchbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        resultswin& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory Search Results")
        snlist& = FindWindowEx(resultswin&, 0&, "_AOL_Listbox", vbNullString)
        okwin& = FindWindow("#32770", "America Online")
    Loop Until resultswin& <> 0& And snlist& <> 0& Or okwin& <> 0&
    If okwin& <> 0& Then
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
        Do: DoEvents
            Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
            okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
        Loop Until okbutton& = 0&
        Call SendMessage(dirwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    Call waitforlisttoload(snlist&)
    For index& = 0& To SendMessage(snlist&, LB_GETCOUNT, 0&, 0&) - 1
        addto.AddItem getsnfrommemdir(getlistitemtext(snlist&, index&))
    Next index&
    Call SendMessage(dirwin&, WM_CLOSE, 0&, 0&)
    Call SendMessage(resultswin&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub addmultiplebuddies(buddies As String, separator As String, group As String)
    Dim numofbuddies As Long, numofoldbuddies As Long, whereat As Long
    Dim aol As Long, mdi As Long, buddythere As Long, buddyadded As Long
    Dim setupwin As Long, grouplist As Long, editbutton1 As Long, editbutton As Long
    Dim groupindex As Long, grouptext As String, groupwin As Long, lbcount1 As String
    Dim addbox1 As Long, addbox As Long, buddylist As Long, addbutton As Long
    Dim removebutton As Long, savebutton As Long, errorstatic1 As Long, errorstatic As Long
    Dim statictxt As String, buddiesnum As Long, buddynum As Long, buddy As String
    Dim lbcount As Long, buddyadded1 As String, originalcount As Long, groupnum As String
    group$ = LCase(group$)
    group$ = removechar(group$, " ")
    numofbuddies& = countchar(buddies$, separator$)
    numofoldbuddies& = getbuddycount(False)
    If numofbuddies& > 99 - numofoldbuddies& Then
        For buddiesnum& = -1& To 99 - numofoldbuddies& - 1
            whereat& = InStr(whereat& + 1, buddies$, separator$)
        Next buddiesnum&
        buddies$ = Left(buddies$, whereat& - Len(separator$))
    End If
    If InStr(buddies$, separator$) = 1& Then buddies$ = Right(buddies$, Len(buddies$) - Len(separator$))
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        If setupwin& = 0& Then setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy List")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0& And editbutton& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&)
        grouptext$ = getlistitemtext(grouplist&, groupindex&)
        whereat& = InStr(grouptext$, Chr(9))
        grouptext$ = Left(grouptext$, whereat& - 1)
        grouptext$ = LCase(grouptext$)
        grouptext$ = removechar(grouptext$, " ")
        If grouptext$ = group$ Then Exit For
    Next groupindex&
    If groupindex& = -1& Then
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    grouptext$ = getlistitemtext(grouplist&, groupindex&)
    groupnum$ = Right(grouptext$, Len(grouptext$) - whereat&)
    grouptext$ = Left(grouptext$, whereat& - 1)
    Call PostMessage(grouplist&, LB_SETCURSEL, CLng(groupindex&), 0&)
    Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        groupwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & grouptext$)
        addbox1& = FindWindowEx(groupwin&, 0&, "_AOL_Edit", vbNullString)
        addbox& = FindWindowEx(groupwin&, addbox1&, "_AOL_Edit", vbNullString)
        buddylist& = FindWindowEx(groupwin&, 0&, "_AOL_Listbox", vbNullString)
        addbutton& = FindWindowEx(groupwin&, 0&, "_AOL_Icon", vbNullString)
        removebutton& = FindWindowEx(groupwin&, addbutton&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(groupwin&, removebutton&, "_AOL_Icon", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, 0&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
    Loop Until groupwin& <> 0& And addbox& <> 0& And buddylist& <> 0& And addbutton& <> 0& And removebutton& <> 0& And savebutton& <> 0& And errorstatic& <> 0&
    Do: DoEvents
        lbcount1$ = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&)
    Loop Until lbcount1$ = groupnum$
    originalcount& = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&)
    For buddynum& = 0& To numofbuddies& - 1
        If buddynum& + numofoldbuddies& >= 99 Then Exit For
        whereat& = 0&
        Do: DoEvents
            whereat& = InStr(whereat& + 1, buddies$, separator$)
            If whereat& = 0& Then Exit For
            buddy$ = Left(buddies$, whereat& - Len(separator$))
            buddies$ = Right(buddies$, Len(buddies$) - Len(buddy$) - Len(separator$))
            buddythere& = getlistitemindex(buddylist&, buddy$, True, True)
            buddy$ = LCase(buddy$)
            buddy$ = removechar(buddy$, " ")
            If buddythere& = -1& Then Exit Do
        Loop Until buddythere& = -1&
        Call SendMessageByString(addbox&, WM_SETTEXT, 0&, buddy$)
        Call PostMessage(addbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(addbutton&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            lbcount& = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&)
            buddyadded1$ = getlistitemtext(buddylist&, lbcount& - 1)
            statictxt$ = gettext(errorstatic&)
        Loop Until buddyadded1$ = buddy$ Or statictxt$ = "Screen Name already in list." Or statictxt$ = "Screen Name too short."
        Call SendMessageByString(errorstatic&, WM_SETTEXT, 0&, "")
    Next buddynum&
    If originalcount& = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&) Then
        Call PostMessage(groupwin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
    End If
    If buddies$ = "" Or buddies$ = separator$ Then
        Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    If buddynum& + numofoldbuddies& >= 99 Then
        Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    buddythere& = getlistitemindex(buddylist&, buddies$, True, True)
    If buddythere& <> -1& Then
        Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    Call SendMessageByString(addbox&, WM_SETTEXT, 0&, buddies$)
    Call PostMessage(addbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(addbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        lbcount& = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&)
        buddyadded1$ = getlistitemtext(buddylist&, lbcount& - 1)
        statictxt$ = gettext(errorstatic&)
    Loop Until buddyadded1$ = buddy$ Or statictxt$ = "Screen Name already in list." Or statictxt$ = "Screen Name too short."
    Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function addroomtostring(separator As String) As String
    'parts taken from dos's addroomtolist sub
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long, index As Long
    Dim screenname As String, itmhold As Long, psnhold As Long
    Dim rbytes As Long, cprocess As Long, room As Long
    room& = findroom()
    rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
            Call CopyMemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            addroomtostring$ = addroomtostring$ & separator$ & screenname$
        Next index&
        Call CloseHandle(mthread&)
    End If
End Function
Public Sub addtreetocontrol(thetree As Long, addto As Control)
    Dim listcount As Long, index As Long, indexlength As Long, thestring As String
    listcount& = SendMessage(thetree&, LB_GETCOUNT, 0&, 0&)
    If listcount& = 0& Then Exit Sub
    For index& = 0& To listcount& - 1
        indexlength& = SendMessage(thetree&, LB_GETTEXTLEN, index&, 0&)
        thestring$ = String(indexlength& + 1, 0)
        Call SendMessageByString(thetree&, LB_GETTEXT, index&, thestring$)
        addto.AddItem thestring$
    Next index&
End Sub

Public Function addtreetostring(thetree As Long, separator As String) As String
    Dim listcount As Long, index As Long, indexlength As Long, thestring As String
    listcount& = SendMessage(thetree&, LB_GETCOUNT, 0&, 0&)
    If listcount& = 0& Then Exit Function
    For index& = 0& To listcount& - 1
        indexlength& = SendMessage(thetree&, LB_GETTEXTLEN, index&, 0&)
        thestring$ = String(indexlength& + 1, 0)
        Call SendMessageByString(thetree&, LB_GETTEXT, index&, thestring$)
        addtreetostring$ = addtreetostring$ & separator$ & thestring$
    Next index&
End Function

Public Sub buddyinvite(buddies As String, tosay As String, room As String, gotochat As Boolean)
    Dim aol As Long, mdi As Long, buddywin As Long, inviteicon As Long, invitewin As Long
    Dim peoplebox As Long, tosaybox As Long, roombox As Long, sendicon As Long
    Dim itationwin As Long, goicon As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy chat")
    ElseIf buddywin& <> 0& Then
        inviteicon& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        inviteicon& = FindWindowEx(buddywin&, inviteicon&, "_AOL_Icon", vbNullString)
        inviteicon& = FindWindowEx(buddywin&, inviteicon&, "_AOL_Icon", vbNullString)
        inviteicon& = FindWindowEx(buddywin&, inviteicon&, "_AOL_Icon", vbNullString)
        Call PostMessage(inviteicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(inviteicon&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        invitewin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy Chat")
        peoplebox& = FindWindowEx(invitewin&, 0&, "_AOL_Edit", vbNullString)
        tosaybox& = FindWindowEx(invitewin&, peoplebox&, "_AOL_Edit", vbNullString)
        roombox& = FindWindowEx(invitewin&, tosaybox&, "_AOL_Edit", vbNullString)
        sendicon& = FindWindowEx(invitewin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until invitewin& <> 0& And peoplebox& <> 0& And tosaybox& <> 0& And roombox& <> 0& And sendicon& <> 0&
    Call SendMessageByString(peoplebox&, WM_SETTEXT, 0&, buddies$)
    Call SendMessageByString(tosaybox&, WM_SETTEXT, 0&, tosay$)
    Call SendMessageByString(roombox&, WM_SETTEXT, 0&, room$)
    Call PostMessage(sendicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        itationwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Invitation from: " & getuser())
        goicon& = FindWindowEx(itationwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until itationwin& <> 0& And goicon& <> 0&
    If gotochat = True Then
        Call PostMessage(goicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(goicon&, WM_LBUTTONUP, 0&, 0&)
    ElseIf gotochat = False Then
        Call PostMessage(itationwin&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub chatignoreindex(index As Long, onoff As Boolean)
    Dim aol As Long, mdi As Long, chat As Long, list As Long
    Dim ignorewin As Long, checkbox As Long, screenname As String
    Dim checkif As CHECKVALUE
    If index& > getroomcount() Then Exit Sub
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    chat& = findroom()
    list& = FindWindowEx(chat&, 0&, "_AOL_Listbox", vbNullString)
    screenname$ = getlistitemtext(list&, index&)
    Call SendMessage(list&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(list&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do: DoEvents
        ignorewin& = FindWindowEx(mdi&, 0&, "AOL Child", screenname$)
        checkbox& = FindWindowEx(ignorewin&, 0&, "_AOL_Checkbox", vbNullString)
    Loop Until ignorewin& <> 0& And checkbox& <> 0&
    checkif = getcheckboxvalue(checkbox&)
    If checkif = cvUNCHECKED And onoff = True Then
        Call PostMessage(checkbox&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(checkbox&, WM_LBUTTONUP, 0&, 0&)
    ElseIf checkif = cvCHECKED And onoff = False Then
        Call PostMessage(checkbox&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(checkbox&, WM_LBUTTONUP, 0&, 0&)
    End If
    pause 0.3
    Call SendMessage(ignorewin&, WM_CLOSE, 0&, 0&)
End Sub

Public Function checkifalive(person As String) As Boolean
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim mailbutton As Long, mailwin As Long, personbox As Long, ccbox As Long
    Dim subjectbox As Long, messagebox As Long, sendbutton1 As Long, sendbutton2 As Long
    Dim sendbutton As Long, okwin As Long, okbutton As Long, maillist As Long
    Dim index As Long, indextext As String, nowin As Long, nobutton As Long
    Dim nobutton1 As Long, view As Long, viewtxt As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    mailbutton& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    mailbutton& = FindWindowEx(toolbar2&, mailbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(mailbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(mailbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
        personbox& = FindWindowEx(mailwin&, 0&, "_AOL_Edit", vbNullString)
        ccbox& = FindWindowEx(mailwin&, personbox&, "_AOL_Edit", vbNullString)
        subjectbox& = FindWindowEx(mailwin&, ccbox&, "_AOL_Edit", vbNullString)
        messagebox& = FindWindowEx(mailwin&, 0&, "RICHCNTL", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, 0&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton2& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton& = FindWindowEx(mailwin&, sendbutton2&, "_AOL_Icon", vbNullString)
    Loop Until mailwin& <> 0& And personbox& <> 0& And ccbox& <> 0& And subjectbox& <> 0& And messagebox& <> 0& And sendbutton& <> 0& And sendbutton& <> sendbutton1& And sendbutton1& <> sendbutton2& And sendbutton2& <> 0& And sendbutton1 <> 0&
    Call SendMessageByString(personbox&, WM_SETTEXT, 0&, Chr(34) & "," & person$)
    Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, "hello " & person$)
    Call SendMessageByString(messagebox&, WM_SETTEXT, 0&, "just checking if you still had aol or not :)")
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
        okwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
        view& = FindWindowEx(okwin&, 0&, "_AOL_View", vbNullString)
        okbutton& = FindWindowEx(okwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until mailwin& = 0& Or okwin& <> 0& And view& <> 0& And okbutton& <> 0&
    viewtxt$ = gettext(view&)
    Do: DoEvents
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        okwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Error")
        okbutton& = FindWindowEx(okwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until okbutton& = 0& And okwin& = 0&
    Call PostMessage(mailwin&, WM_CLOSE, 0&, 0&)
    Do: DoEvents
        nowin& = FindWindow("#32770", "America Online")
        nobutton1& = FindWindowEx(nowin&, 0&, "Button", vbNullString)
        nobutton& = FindWindowEx(nowin&, nobutton1&, "Button", vbNullString)
    Loop Until nowin& <> 0& And nobutton& <> 0&
    Do: DoEvents
        Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
        nobutton& = FindWindowEx(nowin&, 0&, "Button", "&No")
    Loop Until nobutton& = 0&
    If InStr(viewtxt$, person$ & " - This is not a known member.") <> 0& Then
        checkifalive = False
    ElseIf InStr(viewtxt$, person$ & " - This is not a known member.") = 0& Then
        checkifalive = True
    End If
End Function
Public Function checkifghost(person As String) As Boolean
    'thanks for the idea kidweb
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, dirwin As Long, searchbox As Long, checkbox1 As Long
    Dim checkbox As Long, searchbutton1 As Long, searchbutton As Long
    Dim resultswin As Long, snlist As Long, ha As String, index As Long
    Dim name As String, whereat As Long, checkedcheck As CHECKVALUE
    Dim okwin As Long, okbutton As Long
    person$ = LCase(person$)
    person$ = removechar(person$, " ")
    If checkifavailible(person$) = True Then
        checkifghost = False
        Exit Function
    End If
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 9&)
    Do: DoEvents
        dirwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
        searchbox& = FindWindowEx(dirwin&, 0&, "_AOL_Edit", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, 0&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, 0&, "_AOL_Icon", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
        searchbutton& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
    Loop Until dirwin& <> 0& And searchbox& <> 0& And checkbox& <> 0& And searchbutton& <> 0&
    Call SendMessageByString(searchbox&, WM_SETTEXT, 0&, person$)
    Call PostMessage(checkbox&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(checkbox&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(searchbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(searchbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        resultswin& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory Search Results")
        snlist& = FindWindowEx(resultswin&, 0&, "_AOL_Listbox", vbNullString)
        okwin& = FindWindow("#32770", "America Online")
    Loop Until resultswin& <> 0& And snlist& <> 0& Or okwin& <> 0&
    If okwin& <> 0& Then
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
        Do: DoEvents
            Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
            okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
        Loop Until okbutton& = 0&
        Call SendMessage(dirwin&, WM_CLOSE, 0&, 0&)
        checkifghost = False
        Exit Function
    End If
    Call waitforlisttoload(snlist&)
    For index& = 0& To SendMessage(snlist&, LB_GETCOUNT, 0&, 0&) - 1
        name$ = getlistitemtext(snlist&, index&)
        whereat& = InStr(2&, name$, Chr(9))
        name$ = Mid(name$, 2&, whereat& - 2)
        name$ = LCase(name$)
        name$ = removechar(name$, " ")
        If name$ = person$ Then
            checkifghost = True
            Exit For
        End If
        checkifghost = False
    Next index&
    Call SendMessage(dirwin&, WM_CLOSE, 0&, 0&)
    Call SendMessage(resultswin&, WM_CLOSE, 0&, 0&)
End Function
Public Function checkifmaster() As Boolean
    Dim aol As Long, mdi As Long, controlwin As Long, trybutton As Long
    Dim modal As Long, modalstatic As Long, modaltext As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://4344:1580.prntcon.12263709.564517913")
    Do: DoEvents
        controlwin& = FindWindowEx(mdi&, 0&, "AOL Child", " Parental Controls")
        trybutton& = FindWindowEx(controlwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until controlwin& <> 0& And trybutton& <> 0&
    Call PostMessage(trybutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(trybutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        modal& = FindWindow("_AOL_Modal", vbNullString)
        modalstatic& = FindWindowEx(modal&, 0&, "_AOL_Static", vbNullString)
    Loop Until modal& <> 0&
    modaltext$ = gettext(modalstatic&)
    If Left(modaltext$, Len(modaltext$) - 1) <> "Set Parental Controls" Then
        Call SendMessage(modal&, WM_CLOSE, 0&, 0&)
        Call SendMessage(controlwin&, WM_CLOSE, 0&, 0&)
        checkifmaster = False
    ElseIf Left(modaltext$, Len(modaltext$) - 1) = "Set Parental Controls" Then
        Call SendMessage(modal&, WM_CLOSE, 0&, 0&)
        Call SendMessage(controlwin&, WM_CLOSE, 0&, 0&)
        checkifmaster = True
    End If
End Function
Public Function countchar(thestring As String, thechar As String) As Long
    Dim whereat As Long
    whereat& = 0&
    Do
        whereat& = InStr(whereat& + 1, thestring$, thechar$)
        If whereat& = 0& Then Exit Do
        countchar& = countchar& + 1
    Loop
End Function
Public Sub createbuddygroup(newgroup As String)
    Dim newgroup2 As String, aol As Long, mdi As Long, buddywin As Long
    Dim setupbutton As Long, setupwin As Long, grouplist As Long, createbutton As Long
    Dim whereat As Long, groupindex As Long, newwin As Long, newbox As Long, groupname As String
    Dim savebutton1 As Long, savebutton As Long
    newgroup2$ = LCase(newgroup$)
    newgroup2$ = removechar(newgroup2$, " ")
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0& Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
        createbutton& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0& And createbutton& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&) - 1
        groupname$ = getlistitemtext(grouplist&, groupindex&)
        groupname$ = LCase(groupname$)
        groupname$ = removechar(groupname$, " ")
        whereat& = InStr(groupname$, Chr(9))
        groupname$ = Left(groupname$, whereat& - 1)
        If groupname$ = newgroup2$ Then
            Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
            Exit Sub
        End If
    Next groupindex&
    Call PostMessage(createbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(createbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        newwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Create a Buddy List Group")
        newbox& = FindWindowEx(newwin&, 0&, "_AOL_Edit", vbNullString)
        savebutton1& = FindWindowEx(newwin&, 0&, "_AOL_Icon", vbNullString)
        savebutton1& = FindWindowEx(newwin&, savebutton1&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(newwin&, savebutton1&, "_AOL_Icon", vbNullString)
    Loop Until newwin& <> 0& And newbox& <> 0& And savebutton& <> 0&
    Call SendMessageByString(newbox&, WM_SETTEXT, 0&, newgroup$)
    Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
    Call SendMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub createshortcut(theshortcutdirectory As String, theshortcutname As String, theshortcutexepath As String)
Dim retval As Long
retval& = CreateShellLink("", theshortcutname$, theshortcutexepath$, "")
Name "c:\windows\start menu\programs\" & theshortcutname$ & ".LNK" As theshortcutdirectory$ & "\" & theshortcutname$ & ".LNK"
End Sub
Public Sub deletebuddygroup(delgroup As String)
    Dim newgroup2 As String, aol As Long, mdi As Long, buddywin As Long
    Dim setupbutton As Long, setupwin As Long, grouplist As Long, deletebutton1 As Long
    Dim DeleteButton As Long, groupindex As Long, groupname As String, whereat As Long
    Dim modal As Long, okbutton As Long
    newgroup2$ = LCase(delgroup$)
    newgroup2$ = removechar(newgroup2$, " ")
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0& Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
        deletebutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        deletebutton1& = FindWindowEx(setupwin&, deletebutton1&, "_AOL_Icon", vbNullString)
        DeleteButton& = FindWindowEx(setupwin&, deletebutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0& And DeleteButton& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&) - 1
        groupname$ = getlistitemtext(grouplist&, groupindex&)
        groupname$ = LCase(groupname$)
        groupname$ = removechar(groupname$, " ")
        whereat& = InStr(groupname$, Chr(9))
        groupname$ = Left(groupname$, whereat& - 1)
        If groupname$ = newgroup2$ Then Exit For
    Next groupindex&
    If groupindex& = -1& Then Exit Sub
    Call PostMessage(grouplist&, LB_SETCURSEL, CLng(groupindex&), 0&)
    Call PostMessage(DeleteButton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(DeleteButton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        modal& = FindWindow("_AOL_Modal", vbNullString)
        okbutton& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until modal& <> 0& And okbutton& <> 0&
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub deletelocation(location As String)
    Dim aol As Long, mdi As Long, signonwin As Long, setupbutton As Long
    Dim setupbuttona As Long, setupwin As Long, setupbutton1 As Long, thestring As String
    Dim connectwin As Long, connecttabs As Long, connecttab As Long, index As Long
    Dim treelist As Long, deletelocationbutton1 As Long, deletelocationbutton As Long
    Dim yeswin As Long, yesbutton As Long
    location$ = LCase(location$)
    location$ = removechar(location$, " ")
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    signonwin& = findsignonwin()
    If signonwin& = 0& Then
        Call opensignonwin
        Do: DoEvents
            signonwin& = findsignonwin()
            setupbutton& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
            setupbuttona& = FindWindowEx(signonwin&, setupbutton&, "_AOL_Icon", vbNullString)
        Loop Until signonwin& <> 0& And setupbuttona& <> 0&
    End If
    setupbutton& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
    setupbuttona& = FindWindowEx(signonwin&, setupbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(setupbuttona&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(setupbuttona&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        setupwin& = FindWindow("_AOL_Modal", "AOL Setup")
        setupbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And setupbutton1& <> 0&
    Call PostMessage(setupbutton1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(setupbutton1&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        connectwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Connection Setup")
        connecttabs& = FindWindowEx(connectwin&, 0&, "_AOL_TabControl", vbNullString)
        connecttab& = FindWindowEx(connecttabs&, 0&, "_AOL_TabPage", vbNullString)
        treelist& = FindWindowEx(connecttab&, 0&, "_AOL_Tree", vbNullString)
        deletelocationbutton1& = FindWindowEx(connecttab&, 0&, "_AOL_Icon", vbNullString)
        deletelocationbutton1& = FindWindowEx(connecttab&, deletelocationbutton1&, "_AOL_Icon", vbNullString)
        deletelocationbutton1& = FindWindowEx(connecttab&, deletelocationbutton1&, "_AOL_Icon", vbNullString)
        deletelocationbutton& = FindWindowEx(connecttab&, deletelocationbutton1&, "_AOL_Icon", vbNullString)
    Loop Until connectwin& <> 0& And connecttabs& <> 0& And connecttab& <> 0& And deletelocationbutton& <> 0&
    For index& = 0& To gettreecount(treelist&)
        thestring$ = gettreeitemtext(treelist&, index&)
        thestring$ = removechar(thestring$, " ")
        thestring$ = Left(thestring$, Len(thestring$) - 1)
        thestring$ = LCase(thestring$)
        If thestring$ = location$ Then Exit For
    Next index&
    Call PostMessage(treelist&, LB_SETCURSEL, CLng(index&), 0&)
    Call PostMessage(deletelocationbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(deletelocationbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        yeswin& = FindWindow("#32770", "America Online")
        yesbutton& = FindWindowEx(yeswin&, 0&, "Button", vbNullString)
    Loop Until yeswin& <> 0& And yesbutton& <> 0&
    Call PostMessage(yesbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(yesbutton&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(connectwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function doublestring(thestring As String) As String
    Dim char As Long, first As String, newstring As String
    For char& = 1& To Len(thestring$)
        first$ = Mid(thestring$, char&, 1)
        newstring$ = newstring$ & first$ & first$
    Next char&
    doublestring$ = newstring$
End Function
Public Sub downloadlater(mailwin As Long)
    Dim downloadbutton As Long
    downloadbutton& = FindWindowEx(mailwin&, 0&, "_AOL_Icon", vbNullString)
    downloadbutton& = FindWindowEx(mailwin&, downloadbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(downloadbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(downloadbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function findflashmailwin() As Long
    Dim aol As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    findflashmailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
End Function
Public Function findopenmailwin() As Long
    Dim aol As Long, mdi As Long, child As Long, richcntl As Long
    Dim childtxt As String, richtxt As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    richcntl& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    childtxt$ = getcaption(child&)
    richtxt$ = gettext(richcntl&)
    If Left(richtxt$, Len(childtxt$) + 6) = "Subj:" & Chr(9) & childtxt$ Then
        findopenmailwin& = child&
        Exit Function
    Else
        Do: DoEvents
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            richcntl& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            childtxt$ = getcaption(child&)
            richtxt$ = gettext(richcntl&)
            If Left(richtxt$, Len(childtxt$) + 6) = "Subj:" & Chr(9) & childtxt$ Then
                findopenmailwin& = child&
                Exit Function
            End If
        Loop Until child& = 0&
        findopenmailwin& = child&
    End If
    
End Function

Public Function findsendwin() As Long
    Dim aol As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    findsendwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
End Function
Public Function findsignonwin() As Long
    Dim aol As Long, mdi As Long, child As Long, combo1 As Long
    Dim combo2 As Long, static1 As Long, static2 As Long, static3 As Long
    Dim static4 As Long, edit1 As Long, icon1 As Long, icon2 As Long
    Dim icon3 As Long, icon4 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    combo1& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
    combo2& = FindWindowEx(child&, combo1&, "_AOL_Combobox", vbNullString)
    static1& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    static2& = FindWindowEx(child&, static1&, "_AOL_Static", vbNullString)
    static3& = FindWindowEx(child&, static2&, "_AOL_Static", vbNullString)
    static4& = FindWindowEx(child&, static3&, "_AOL_Static", vbNullString)
    edit1& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
    icon1& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    icon2& = FindWindowEx(child&, icon1&, "_AOL_Icon", vbNullString)
    icon3& = FindWindowEx(child&, icon2&, "_AOL_Icon", vbNullString)
    icon4& = FindWindowEx(child&, icon3&, "_AOL_Icon", vbNullString)
    If combo1& <> 0& And combo2 <> 0& And static1& <> 0& And static2& <> 0& And static3& <> 0& And static4& <> 0& And edit1& <> 0& And icon1& <> 0& And icon2& <> 0& And icon3& <> 0& And icon4& <> 0& Then
        findsignonwin& = child&
    Else
        Do: DoEvents
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            combo1& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
            combo2& = FindWindowEx(child&, combo1&, "_AOL_Combobox", vbNullString)
            static1& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            static2& = FindWindowEx(child&, static1&, "_AOL_Static", vbNullString)
            static3& = FindWindowEx(child&, static2&, "_AOL_Static", vbNullString)
            static4& = FindWindowEx(child&, static3&, "_AOL_Static", vbNullString)
            edit1& = FindWindowEx(child&, 0&, "_AOL_Edit", vbNullString)
            icon1& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            icon2& = FindWindowEx(child&, icon1&, "_AOL_Icon", vbNullString)
            icon3& = FindWindowEx(child&, icon2&, "_AOL_Icon", vbNullString)
            icon4& = FindWindowEx(child&, icon3&, "_AOL_Icon", vbNullString)
            If combo1& <> 0& And combo2 <> 0& And static1& <> 0& And static2& <> 0& And static3& <> 0& And static4& <> 0& And edit1& <> 0& And icon1& <> 0& And icon2& <> 0& And icon3& <> 0& And icon4& <> 0& Then
                findsignonwin& = child&
                Exit Function
            End If
        Loop Until child& = 0&
        Exit Function
    End If
    findsignonwin& = child&
End Function
Public Function findwelcomewin() As Long
    Dim aol As Long, mdi As Long, child As Long, childtxt As String
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    childtxt$ = getcaption(child&)
    If InStr(childtxt$, "Welcome, ") <> 0& Then
        findwelcomewin& = child&
        Exit Function
    Else
        Do: DoEvents
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            childtxt$ = getcaption(child&)
            If InStr(childtxt$, "Welcome, ") <> 0& Then
                findwelcomewin& = child&
                Exit Function
            End If
        Loop Until child& = 0&
    End If
End Function
Public Sub formexit(frm As Form, formexita As FRMEXIT)
    If formexita = feBOTTOM Then
        Do
            frm.Top = frm.Top + 200
        Loop Until frm.Top > Screen.Height
    ElseIf formexita = feBOTTOMLEFT Then
        Do
            frm.Top = frm.Top + 200
            frm.Left = frm.Left - 200
        Loop Until frm.Top > Screen.Height And frm.Left + frm.Width < 0
    ElseIf formexita = feBOTTOMRIGHT Then
        Do
            frm.Top = frm.Top + 200
            frm.Left = frm.Left + 200
        Loop Until frm.Top > Screen.Height And frm.Left > Screen.Width
    ElseIf formexita = feLEFT Then
        Do
            frm.Left = frm.Left - 200
        Loop Until frm.Left + frm.Width < 0
    ElseIf formexita = feRIGHT Then
        Do
            frm.Left = frm.Left + 200
        Loop Until frm.Left > Screen.Width
    ElseIf formexita = feTOP Then
        Do
            frm.Top = frm.Top - 200
        Loop Until frm.Top + frm.Height < 0
    ElseIf formexita = feTOPLEFT Then
        Do
            frm.Left = frm.Left - 200
            frm.Top = frm.Top - 200
        Loop Until frm.Left + frm.Width < 0 And frm.Top + frm.Height < 0
    ElseIf formexita = feTOPRIGHT Then
        Do
            frm.Left = frm.Left + 200
            frm.Top = frm.Top - 200
        Loop Until frm.Top + frm.Height < 0 And frm.Left > Screen.Width
    End If
End Sub
Public Sub formposition(frm As Form, position As FORMPOS)
    If position = fpBOTTOMLEFT Then
        frm.Left = 0
        frm.Top = Screen.Height - frm.Height
    ElseIf position = fpBOTTOMRIGHT Then
        frm.Left = Screen.Width - frm.Width
        frm.Top = Screen.Height - frm.Height
    ElseIf position = fpCENTER Then
        frm.Left = Screen.Width / 2 - frm.Width / 2
        frm.Top = Screen.Height / 2 - frm.Height / 2
    ElseIf position = fpTOPLEFT Then
        frm.Left = 0
        frm.Top = 0
    ElseIf position = fpTOPRIGHT Then
        frm.Left = Screen.Width - frm.Width
        frm.Top = 0
    End If
End Sub


Public Function getbuddycount(closewin As Boolean) As Long
    Dim aol As Long, mdi As Long, buddywin As Long, setupbutton As Long
    Dim setupwin As Long, grouplist As Long, group As String, whereat As Long
    Dim groupindex As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0& Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&) - 1
        group$ = getlistitemtext(grouplist&, groupindex&)
        whereat& = InStr(group$, Chr(9))
        getbuddycount& = getbuddycount& + Right(group$, Len(group$) - whereat&)
    Next groupindex&
    If closewin = True Then Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Function
Public Function getcheckboxvalue(checkbox As Long) As CHECKVALUE
    Dim x As Long
    x& = SendMessage(checkbox&, BM_GETCHECK, 0&, 0&)
    If x& = 0& Then
        getcheckboxvalue = cvUNCHECKED
    ElseIf x& = 1& Then
        getcheckboxvalue = cvCHECKED
    End If
End Function
Function getfromini(appname As String, keyname As String, filename As String) As String
   Dim retstr As String
   retstr$ = String(255, Chr(0))
   getfromini = Left(retstr$, GetPrivateProfileString(appname$, ByVal keyname$, "", retstr, Len(retstr), filename$))
End Function
Public Sub addallbuddystocontrol(addto As Control, addgroups As Boolean)
    Dim aol As Long, mdi As Long, buddylistwin As Long, whereat As Long
    Dim setupbutton As Long, setupwin As Long, editbutton1 As Long
    Dim editbutton As Long, grouplist As Long, groupcount As Long
    Dim editwin As Long, groupname As String, snlist As Long
    Dim index As Long, getcount1 As Long, getcount2 As Long, getcount3 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddylistwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddylistwin& = 0& Then
        Call keyword("buddy list")
    Else
        setupbutton& = FindWindowEx(buddylistwin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddylistwin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddylistwin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        If setupwin& = 0& Then setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy List")
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until setupwin& <> 0& And editbutton& <> 0& And grouplist& <> 0&
    groupcount& = SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        groupname$ = getlistitemtext(grouplist&, 0&)
        whereat& = InStr(groupname$, Chr(9))
        groupname$ = Left(groupname$, whereat& - 1)
        editwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & groupname$)
        snlist& = FindWindowEx(editwin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until editwin& <> 0& And snlist& <> 0&
    For index& = 1 To groupcount&
        groupname$ = getlistitemtext(grouplist&, index& - 1)
        whereat& = InStr(groupname$, Chr(9))
        groupname$ = Left(groupname$, whereat& - 1)
        If addgroups = True Then addto.AddItem ("  " & groupname$)
        Do: DoEvents
            editwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & groupname$)
            snlist& = FindWindowEx(editwin&, 0&, "_AOL_Listbox", vbNullString)
        Loop Until editwin& <> 0& And snlist& <> 0&
        Do: DoEvents
            getcount1& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
            pause 0.2
            getcount2& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
            pause 0.2
            getcount3& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
        Loop Until getcount1& = getcount2& And getcount2& = getcount3&
        Call addlisttocontrol(snlist&, addto)
        Call SendMessage(editwin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(grouplist&, LB_SETCURSEL, CLng(index&), 0&)
        Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Next index&
    Call SendMessage(editwin&, WM_CLOSE, 0&, 0&)
    Call SendMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addbuddystocontrol(addto As Control, addgroups As Boolean)
    On Error Resume Next
    Dim aol As Long, mdi As Long, buddywin As Long
    Dim snlist As Long, getcount1 As Long, getcount2 As Long
    Dim getcount3 As Long, getcount As Long, index As Long
    Dim whereat As Long, snitem As String
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy view")
    End If
    Do: DoEvents
        buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
        snlist& = FindWindowEx(buddywin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until buddywin& <> 0& And snlist& <> 0&
    Do: DoEvents
        getcount1& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
        pause 0.2
        getcount2& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
        pause 0.2
        getcount3& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
    Loop Until getcount1& = getcount2& And getcount2& = getcount3&
    getcount& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
    For index& = 0& To getcount& - 1
        snitem$ = getlistitemtext(snlist&, index&)
        snitem$ = removechar(snitem$, " ")
        snitem$ = removechar(snitem$, " ")
        snitem$ = removechar(snitem$, " ")
        If addgroups = True Then
            whereat& = InStr(snitem$, "/")
            If whereat& <> 0& Then
                whereat& = InStr(snitem$, "(")
                snitem$ = "  " & Left(snitem$, whereat& - 1)
            End If
        End If
        whereat& = InStr(snitem$, "(")
        If whereat& <> 1& And whereat& <> 0& Then
            snitem$ = ""
        ElseIf whereat& <> 0& Then
            snitem$ = Mid(snitem$, whereat& + 1, Len(snitem$) - 2)
        End If
        whereat& = InStr(snitem$, "*")
        If whereat& <> 0& Then
            snitem$ = Left(snitem$, whereat& - 1)
        End If
        If snitem$ <> "" Then addto.AddItem snitem$
    Next index&
End Sub
Public Sub addmailtocontrol(thecontrol As Control, themailtype As MAILTYPE, themailtext As MAILTEXT)
    Dim aol As Long, mdi As Long, mailwin As Long, mailtabs As Long, mailtab As Long, maillist As Long
    Dim listcount As Long, index As Long, indexlength As Long, thestring As String, mailtab1 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    If themailtype = mtNEW Or themailtype = mtOLD Or themailtype = mtSENT Then
        If themailtype = mtNEW Then
            Call openmailbox(mtNEW)
            Do: DoEvents
                mailwin& = findmailwin()
                mailtabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
                mailtab& = FindWindowEx(mailtabs&, 0&, "_AOL_TabPage", vbNullString)
                maillist& = FindWindowEx(mailtab&, 0&, "_AOL_Tree", vbNullString)
            Loop Until mailwin& <> 0& And mailtabs& <> 0& And mailtab& <> 0& And maillist& <> 0&
            Call killmailad(mailwin&)
            Call waitforlisttoload(maillist&)
            listcount& = getlistcount(maillist&)
        ElseIf themailtype = mtOLD Then
            Call openmailbox(mtOLD)
            Do: DoEvents
                mailwin& = findmailwin()
                mailtabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
                mailtab1& = FindWindowEx(mailtabs&, 0&, "_AOL_TabPage", vbNullString)
                mailtab& = FindWindowEx(mailtabs&, mailtab1&, "_AOL_TabPage", vbNullString)
                maillist& = FindWindowEx(mailtab&, 0&, "_AOL_Tree", vbNullString)
            Loop Until mailwin& <> 0& And mailtabs& <> 0& And mailtab& <> 0& And maillist& <> 0&
            Call killmailad(mailwin&)
            Call waitforlisttoload(maillist&)
            listcount& = getlistcount(maillist&)
        ElseIf themailtype = mtSENT Then
            Call openmailbox(mtSENT)
            Do: DoEvents
                mailwin& = findmailwin()
                mailtabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
                mailtab1& = FindWindowEx(mailtabs&, 0&, "_AOL_TabPage", vbNullString)
                mailtab1& = FindWindowEx(mailtabs&, mailtab1&, "_AOL_TabPage", vbNullString)
                mailtab& = FindWindowEx(mailtabs&, mailtab1&, "_AOL_TabPage", vbNullString)
                maillist& = FindWindowEx(mailtab&, 0&, "_AOL_Tree", vbNullString)
            Loop Until mailwin& <> 0& And mailtabs& <> 0& And mailtab& <> 0& And maillist& <> 0&
            Call killmailad(mailwin&)
            Call waitforlisttoload(maillist&)
            listcount& = getlistcount(maillist&)
        End If
    ElseIf themailtype = mtFLASH Then
        Call openmailbox(mtFLASH)
        Do: DoEvents
            mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
            maillist& = FindWindowEx(mailwin&, 0&, "_AOL_Tree", vbNullString)
        Loop Until mailwin& <> 0& And maillist& <> 0&
        Call waitforlisttoload(maillist&)
        listcount& = getlistcount(maillist&)
    End If
    For index& = 0 To listcount& - 1
        indexlength& = SendMessage(maillist&, LB_GETTEXTLEN, index&, 0&)
        thestring$ = String(indexlength& + 1, 0)
        Call SendMessageByString(maillist&, LB_GETTEXT, index&, thestring$)
        If themailtext = mtDATE Then
            thestring$ = getmaildate(thestring$)
        ElseIf themailtext = mtSENDER Then
            thestring$ = getmailsender(thestring$)
        ElseIf themailtext = mtSUBJECT Then
            thestring$ = getmailsubject(thestring$)
        End If
        If InStr(thestring$, " ") = 1 Then thestring$ = Right(thestring$, Len(thestring$) - 1)
        thecontrol.AddItem (thestring$)
    Next index&
    Call SendMessage(mailwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addstringtocontrol(thestring As String, thecontrol As Control, separator As String, lcasetext As Boolean, removespaces As Boolean)
    Dim whereat As Long, theitem As String
    If thestring$ = "" Then Exit Sub
    whereat& = InStr(thestring$, separator$)
    thecontrol.AddItem Left(thestring$, whereat& - 1)
    thestring$ = Right(thestring$, Len(thestring$) - whereat&)
    Do: DoEvents
        whereat& = InStr(thestring$, separator$)
        If whereat& = 0& Then Exit Do
        theitem$ = Left(thestring$, whereat& - 1)
        theitem$ = Right(theitem$, Len(theitem$) - Len(separator$) + 1)
        If lcasetext = True Then theitem$ = LCase(theitem$)
        If removespaces = True Then theitem$ = removechar(theitem$, " ")
        thecontrol.AddItem (theitem$)
        thestring$ = Right(thestring$, Len(thestring$) - whereat&)
    Loop Until whereat& = 0&
    thestring$ = Right(thestring$, Len(thestring$) - Len(separator$) + 1)
    thecontrol.AddItem (thestring$)
End Sub
Public Sub chatignoretext(person As String, exactsn As Boolean, onoff As Boolean)
    Dim aol As Long, mdi As Long, chat As Long, list As Long, sthread As Long
    Dim mthread As Long, cprocess As Long, itmhold As Long, rbytes As Long
    Dim screenname As String, psnhold As Long, ignorewin As Long, checkbox As Long
    Dim index As Long, checkif As CHECKVALUE, curpos As POINTAPI
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    chat& = findroom()
    list& = FindWindowEx(chat&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = GetWindowThreadProcessId(list&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(list&, LB_GETCOUNT, 0&, 0&) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = SendMessage(list, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
            Call CopyMemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If exactsn = False Then
                If InStr(LCase(screenname$), LCase(person$)) <> 0& Then Exit For
            End If
            If LCase(screenname$) = LCase(person$) Then Exit For
        Next index&
        Call CloseHandle(mthread&)
    End If
    Call SendMessage(list&, LB_SETCURSEL, index&, 0&)
    Call PostMessage(list&, WM_LBUTTONDBLCLK, 0&, 0&)
    Do: DoEvents
        ignorewin& = FindWindowEx(mdi&, 0&, "AOL Child", screenname$)
        checkbox& = FindWindowEx(ignorewin&, 0&, "_AOL_Checkbox", vbNullString)
    Loop Until ignorewin& <> 0& And checkbox& <> 0&
    checkif = getcheckboxvalue(checkbox&)
    If checkif = cvUNCHECKED And onoff = True Then
        Call PostMessage(checkbox&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(checkbox&, WM_LBUTTONUP, 0&, 0&)
    ElseIf checkif = cvCHECKED And onoff = False Then
        Call PostMessage(checkbox&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(checkbox&, WM_LBUTTONUP, 0&, 0&)
    End If
    pause 0.3
    Call SendMessage(ignorewin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub clickaoliconmenu(icon As Long, itemnum As Long)
    'parts taken from dos32.bas
    Dim smod As Long, winvis As Long, dothis As Long
    Dim curpos As POINTAPI
    Call GetCursorPos(curpos)
    Call SetCursorPos(Screen.Width, Screen.Height)
    Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
    Do
        smod& = FindWindow("#32768", vbNullString)
        winvis& = IsWindowVisible(smod&)
    Loop Until winvis& = 1
    For dothis& = 1 To itemnum&
        Call PostMessage(smod&, WM_KEYDOWN, VK_DOWN, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_DOWN, 0&)
    Next dothis&
    Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
    Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
    Call SetCursorPos(curpos.x, curpos.Y)
End Sub
Public Sub deletemail(maillist As Long, index As Long, themailtype As MAILTYPE)
    Dim listtab As Long, listtabs As Long, listwin As Long, delbutton As Long
    If themailtype = mtNEW Or themailtype = mtOLD Or themailtype = mtSENT Then
        listtab& = GetParent(maillist&)
        If listtab& = 0& Then Exit Sub
        listtabs& = GetParent(listtab&)
        listwin& = GetParent(listtabs&)
    ElseIf themailtype = mtFLASH Then
        listwin& = GetParent(maillist&)
        If listwin& = 0& Then Exit Sub
    End If
    Call PostMessage(maillist&, LB_SETCURSEL, CLng(index&), 0&)
    delbutton& = FindWindowEx(listwin&, 0&, "_AOL_Icon", vbNullString)
    delbutton& = FindWindowEx(listwin&, delbutton&, "_AOL_Icon", vbNullString)
    delbutton& = FindWindowEx(listwin&, delbutton&, "_AOL_Icon", vbNullString)
    delbutton& = FindWindowEx(listwin&, delbutton&, "_AOL_Icon", vbNullString)
    delbutton& = FindWindowEx(listwin&, delbutton&, "_AOL_Icon", vbNullString)
    delbutton& = FindWindowEx(listwin&, delbutton&, "_AOL_Icon", vbNullString)
    delbutton& = FindWindowEx(listwin&, delbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(delbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(delbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function findforwardwin() As Long
    Dim aol As Long, mdi As Long, child As Long, childtxt As String
    Dim iffwd As String, static1 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    static1& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
    childtxt$ = getcaption(child&)
    iffwd$ = Left(childtxt$, 5)
    If iffwd$ <> "Write Mail" And static1& <> 0& Then
        findforwardwin& = child&
        Exit Function
    Else
        Do: DoEvents
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            static1& = FindWindowEx(child&, 0&, "_AOL_Static", "Send Now")
            childtxt$ = getcaption(child&)
            iffwd$ = Left(childtxt$, 5)
            If iffwd$ <> "Write Mail" And static1& <> 0& Then
                findforwardwin& = child&
                Exit Function
            End If
        Loop Until child& = 0&
        findforwardwin& = child&
    End If
End Function
Public Function findmailwin() As Long
    Dim aol As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    findmailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
End Function
Public Sub forwardmail(mailsubject As String, sendto As String, message As String, removefwd As Boolean)
    Dim aol As Long, mdi As Long, mailwin As Long, fwdbutton As Long
    Dim fwdwin As Long, sendtobox As Long, subjectbox1 As Long, subjectbox As Long
    Dim messagebox As Long, sendbutton1 As Long, sendbutton As Long
    Dim subject As String, sendbutton2 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", mailsubject$)
    If mailwin& = 0& Then Exit Sub
    fwdbutton& = FindWindowEx(mailwin&, 0&, "_AOL_Icon", vbNullString)
    fwdbutton& = FindWindowEx(mailwin&, fwdbutton&, "_AOL_Icon", vbNullString)
    fwdbutton& = FindWindowEx(mailwin&, fwdbutton&, "_AOL_Icon", vbNullString)
    fwdbutton& = FindWindowEx(mailwin&, fwdbutton&, "_AOL_Icon", vbNullString)
    fwdbutton& = FindWindowEx(mailwin&, fwdbutton&, "_AOL_Icon", vbNullString)
    fwdbutton& = FindWindowEx(mailwin&, fwdbutton&, "_AOL_Icon", vbNullString)
    fwdbutton& = FindWindowEx(mailwin&, fwdbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(fwdbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(fwdbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        fwdwin& = findforwardwin()
        sendtobox& = FindWindowEx(fwdwin&, 0&, "_AOL_Edit", vbNullString)
        subjectbox1& = FindWindowEx(fwdwin&, sendtobox&, "_AOL_Edit", vbNullString)
        subjectbox& = FindWindowEx(fwdwin&, subjectbox1&, "_AOL_Edit", vbNullString)
        messagebox& = FindWindowEx(fwdwin&, 0&, "RICHCNTL", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, 0&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(fwdwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton2& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton& = FindWindowEx(mailwin&, sendbutton2&, "_AOL_Icon", vbNullString)
    Loop Until fwdwin& <> 0& And sendtobox& <> 0& And subjectbox& <> 0& And messagebox& <> 0& And sendbutton& <> 0& And sendbutton& <> sendbutton1& And sendbutton1& <> sendbutton2& And sendbutton2& <> 0& And sendbutton1 <> 0&
    Call SendMessageByString(sendtobox&, WM_SETTEXT, 0&, sendto$)
    Call SendMessageByString(messagebox&, WM_SETTEXT, 0&, message$)
    If removefwd = True Then
        subject$ = gettext(subjectbox&)
        subject$ = Right(subject$, Len(subject$) - 5)
        Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, subject$)
    End If
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(mailwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function getchatname() As String
    Dim chatroom As Long
    chatroom& = findroom()
    If chatroom& = 0& Then getchatname$ = "Not in chat."
    getchatname$ = gettext(chatroom&)
End Function
Public Function getlocationcount(signonwin As Long) As Long
    Dim locationcombo As Long
    locationcombo& = FindWindowEx(signonwin&, 0&, "_AOL_Combobox", vbNullString)
    locationcombo& = FindWindowEx(signonwin&, locationcombo&, "_AOL_Combobox", vbNullString)
    getlocationcount& = SendMessage(locationcombo&, CB_GETCOUNT, 0&, 0&)
End Function

Public Function getlocationindex(signonwin As Long, getwhat As String, lcasetext As Boolean, trimspaces As Boolean) As Long
    Dim locationcombo As Long, index As Long, rlist As Long
    Dim sthread As Long, mthread As Long, screenname As String
    Dim itmhold As Long, psnhold As Long, rbytes As Long, cprocess As Long
    locationcombo& = FindWindowEx(signonwin&, 0&, "_AOL_Combobox", vbNullString)
    locationcombo& = FindWindowEx(signonwin&, locationcombo&, "_AOL_Combobox", vbNullString)
    rlist& = locationcombo&
    sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, CB_GETCOUNT, 0&, 0&) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = SendMessage(rlist, CB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
            Call CopyMemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If trimspaces = True Then
                screenname$ = replacestring(screenname$, " ", "")
                screenname$ = replacestring(screenname$, " ", "")
                screenname$ = replacestring(screenname$, " ", "")
            End If
            If lcasetext = True Then screenname$ = LCase(screenname$)
            If screenname$ = getwhat$ Then
                getlocationindex& = index&
                Call CloseHandle(mthread&)
                Exit Function
            End If
        Next index&
        getlocationindex& = -1&
        Call CloseHandle(mthread&)
    End If
End Function
Public Function getlocationtext(signonwin As Long, index As Long) As String
    Dim locationcombo As Long, rlist As Long, cprocess As Long
    Dim sthread As Long, mthread As Long, screenname As String
    Dim itmhold As Long, psnhold As Long, rbytes As Long
    locationcombo& = FindWindowEx(signonwin&, 0&, "_AOL_Combobox", vbNullString)
    locationcombo& = FindWindowEx(signonwin&, locationcombo&, "_AOL_Combobox", vbNullString)
    rlist& = locationcombo&
    sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        screenname$ = String$(4, vbNullChar)
        itmhold& = SendMessage(rlist, CB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmhold& = itmhold& + 24
        Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
        Call CopyMemory(psnhold&, ByVal screenname$, 4)
        psnhold& = psnhold& + 6
        screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        getlocationtext$ = screenname$
        Call CloseHandle(mthread&)
    End If
End Function
Public Function getmailcount(mailtype1 As MAILTYPE) As Long
    Dim aol As Long, mdi As Long, mailwin As Long, list As Long
    Dim tabs As Long, tab1 As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    If mailtype1 = mtFLASH Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
        list& = FindWindowEx(mailwin&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtNEW Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtOLD Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtSENT Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    End If
    getmailcount& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
    Call SendMessage(mailwin&, WM_CLOSE, 0&, 0&)
End Function
Public Function getmaildate(thetext As String) As String
    Dim whereat As Long
    whereat& = InStr(thetext$, Chr(9))
    thetext$ = Left(thetext$, whereat& - 1)
    getmaildate$ = thetext$
End Function
Public Function getmailindex(mailtype1 As MAILTYPE, theline As String, getwhat As MAILTEXT, removespaces As Boolean, lcasetext As Boolean) As Long
    Dim aol As Long, mdi As Long, mailwin As Long, tabs As Long, newtab As Long
    Dim list As Long, ad As Long, getmailtext As String, tab1 As Long
    Dim index As Long, indexlength As Long, thestring As String
    Dim getcount1 As Long, getcount2 As Long, getcount3 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    If mailtype1 = mtFLASH Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
        list& = FindWindowEx(mailwin&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtNEW Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtOLD Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtSENT Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    End If
    Do: DoEvents
        getcount1& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
        pause 0.2
        getcount2& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
        pause 0.2
        getcount3& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
    Loop Until getcount1& = getcount2& And getcount2& = getcount3&
    For index& = 0& To SendMessage(list&, LB_GETCOUNT, 0&, 0&) - 1
        indexlength& = SendMessage(list&, LB_GETTEXTLEN, index&, 0&)
        thestring$ = String(indexlength& + 1, 0)
        Call SendMessageByString(list&, LB_GETTEXT, index&, thestring$)
        If getwhat = mtDATE Then
           thestring$ = getmaildate(thestring$)
        ElseIf getwhat = mtSENDER Then
           thestring$ = getmailsender(thestring$)
        ElseIf getwhat = mtSUBJECT Then
           thestring$ = getmailsubject(thestring$)
        End If
        If removespaces = True Then thestring$ = removechar(thestring$, " ")
        If lcasetext = True Then thestring$ = LCase(thestring$)
        If thestring$ = theline$ Then
            getmailindex& = index&
            Exit Function
        End If
    Next index&
End Function
Public Function getmailtext(mailtype1 As MAILTYPE, getwhat As MAILTEXT, index As Long) As String
    Dim aol As Long, mdi As Long, mailwin As Long, list As Long
    Dim indexlength As Long, thestring As String, tabs As Long, tab1 As Long
    Dim getcount1 As Long, getcount2 As Long, getcount3 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    If mailtype1 = mtFLASH Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
        list& = FindWindowEx(mailwin&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtNEW Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtOLD Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    ElseIf mailtype1 = mtSENT Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Online Mailbox")
        tabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        tab1& = FindWindowEx(tabs&, 0&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        tab1& = FindWindowEx(tabs&, tab1&, "_AOL_TabPage", vbNullString)
        list& = FindWindowEx(tab1&, 0&, "_AOL_Tree", vbNullString)
    End If
    indexlength& = SendMessage(list&, LB_GETTEXTLEN, index&, 0&)
    thestring$ = String(indexlength& + 1, 0)
    Call SendMessageByString(list&, LB_GETTEXT, index&, thestring$)
    If getwhat = mtDATE Then
       getmailtext$ = getmaildate(thestring$)
    ElseIf getwhat = mtSENDER Then
       getmailtext$ = getmailsender(thestring$)
    ElseIf getwhat = mtSUBJECT Then
       getmailtext$ = getmailsubject(thestring$)
    ElseIf getwhat = mtALL Then
        getmailtext$ = thestring$
    End If
End Function
Public Function getmailsender(thetext As String) As String
    Dim whereat As Long, whereat1 As Long
    whereat& = InStr(thetext$, Chr(9))
    whereat1& = InStr(whereat& + 1, thetext$, Chr(9))
    thetext$ = Mid(thetext$, whereat& + 1, whereat1& - whereat& - 1)
    getmailsender$ = thetext$
End Function
Public Function getmailsubject(thetext As String) As String
    Dim whereat As Long
    whereat& = InStr(thetext$, Chr(9))
    whereat& = InStr(whereat& + 1, thetext$, Chr(9))
    thetext$ = Mid(thetext$, whereat& + 1, Len(thetext$) - whereat& - 1)
    getmailsubject$ = thetext$
End Function
Public Function findmaillist(themailtype As MAILTYPE) As Long
    Dim aol As Long, mdi As Long, mailwin As Long, mailtabs As Long, mailtab As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    If themailtype = mtNEW Or themailtype = mtOLD Or themailtype = mtSENT Then
        mailwin& = findmailwin()
        If mailwin& = 0& Then Exit Function
        mailtabs& = FindWindowEx(mailwin&, 0&, "_AOL_TabControl", vbNullString)
        If themailtype = mtNEW Then
            mailtab& = FindWindowEx(mailtabs&, 0&, "_AOL_TabPage", vbNullString)
            findmaillist& = FindWindowEx(mailtab&, 0&, "_AOL_Tree", vbNullString)
            Exit Function
        ElseIf themailtype = mtOLD Then
            mailtab& = FindWindowEx(mailtabs&, 0&, "_AOL_TabPage", vbNullString)
            mailtab& = FindWindowEx(mailtabs&, mailtab&, "_AOL_TabPage", vbNullString)
            findmaillist& = FindWindowEx(mailtab&, 0&, "_AOL_Tree", vbNullString)
            Exit Function
        ElseIf themailtype = mtSENT Then
            mailtab& = FindWindowEx(mailtabs&, 0&, "_AOL_TabPage", vbNullString)
            mailtab& = FindWindowEx(mailtabs&, mailtab&, "_AOL_TabPage", vbNullString)
            mailtab& = FindWindowEx(mailtabs&, mailtab&, "_AOL_TabPage", vbNullString)
            findmaillist& = FindWindowEx(mailtab&, 0&, "_AOL_Tree", vbNullString)
            Exit Function
        End If
    ElseIf themailtype = mtFLASH Then
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Incoming/Saved Mail")
        If mailwin& = 0& Then Exit Function
        findmaillist& = FindWindowEx(mailwin&, 0&, "_AOL_Tree", vbNullString)
        Exit Function
    End If
End Function
Public Function getprofile(screenname As String) As String
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, profilewin As Long, snbox As Long, okbutton As Long
    Dim prowin As Long, profile As Long, protxt1 As String, protxt2 As String
    Dim protxt3 As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 11)
    Do: DoEvents
        profilewin& = FindWindowEx(mdi&, 0&, "AOL Child", "Get a Member's Profile")
        snbox& = FindWindowEx(profilewin&, 0&, "_AOL_Edit", vbNullString)
        okbutton& = FindWindowEx(profilewin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until profilewin& <> 0& And snbox& <> 0& And okbutton& <> 0&
    Call SendMessageByString(snbox&, WM_SETTEXT, 0&, screenname$)
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        prowin& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Profile")
        profile& = FindWindowEx(prowin&, 0&, "_AOL_View", vbNullString)
        Do: DoEvents
            protxt1$ = gettext(profile&)
            pause 0.2
            protxt2$ = gettext(profile&)
            pause 0.2
            protxt3$ = gettext(profile&)
        Loop Until protxt1$ = protxt2$ And protxt2$ = protxt3$
    Loop Until prowin& <> 0& And profile& <> 0&
    getprofile$ = gettext(profile&)
    Call SendMessage(profilewin&, WM_CLOSE, 0&, 0&)
    Call SendMessage(prowin&, WM_CLOSE, 0&, 0&)
End Function
Public Function getroomcount() As Long
    Dim room As Long, list As Long
    room& = findroom()
    list& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    getroomcount& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function getsnfrommemdir(thestring As String) As String
    Dim whereat As Long
    whereat& = InStr(2&, thestring$, Chr(9))
    getsnfrommemdir$ = Mid(thestring$, 2&, whereat& - 2)
End Function

Public Function getstringlinecount(thestring As String) As Long
    Dim tempstring As String, whereat As Long, linenum As Long
    tempstring$ = thestring$
    whereat& = InStr(tempstring$, Chr(13))
    If whereat& = 0& Then
        getstringlinecount& = 1
        Exit Function
    End If
    linenum& = 1
    Do: DoEvents
        whereat& = InStr(whereat& + 1, tempstring$, Chr(13))
        If whereat& = 0& Then Exit Do
        linenum& = linenum& + 1
    Loop Until whereat& = 0&
    getstringlinecount& = linenum&
End Function
Public Function getstringlinetext(thestring As String, index As Long, usedwithgetstringlineindex As Boolean) As String
    Dim tempstring As String, whereat As Long, line As Long
    If index& > getstringlinecount(thestring$) Then Exit Function
    tempstring$ = removechar(thestring$, Chr(10))
    If usedwithgetstringlineindex = False Then index& = index& - 1
    For line& = 0& To index& - 1
        whereat& = InStr(tempstring$, Chr(13))
        If whereat& = 0& Then Exit For
        tempstring$ = Right(tempstring$, Len(tempstring$) - whereat&)
    Next line&
    whereat& = InStr(tempstring$, Chr(13))
    getstringlinetext$ = Left(tempstring$, whereat& - 1)
End Function
Public Function getinfo(getwhat As INFO) As String
    Dim modal As Long, b1 As Long, b2 As Long, b3 As Long, b4 As Long
    Dim b5 As Long, win As Long, txt As Long, ok1 As Long, ok As Long
    Call runmenu(4, 10)
    Do: DoEvents
        modal& = FindWindow("_AOL_Modal", vbNullString)
        b1& = FindWindowEx(modal&, 0&, "_AOL_Button", vbNullString)
        b2& = FindWindowEx(modal&, b1&, "_AOL_Button", vbNullString)
        b3& = FindWindowEx(modal&, b2&, "_AOL_Button", vbNullString)
        b4& = FindWindowEx(modal&, b3&, "_AOL_Button", vbNullString)
        b5& = FindWindowEx(modal&, b4&, "_AOL_Button", vbNullString)
        ok& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until modal& <> 0& And b5& <> 0& And ok& <> 0&
    If getwhat = iCONNECTION Then
        Call PostMessage(b4&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(b4&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            win& = FindWindow("_AOL_Modal", "Connectivity Summary")
            txt& = FindWindowEx(win&, 0&, "RICHCNTL", vbNullString)
            ok1& = FindWindowEx(win&, 0&, "_AOL_Icon", vbNullString)
        Loop Until win& <> 0& And txt& <> 0& And ok1& <> 0&
        Call waitfortxttoload(txt&)
    ElseIf getwhat = iERRORMSG Then
        Call PostMessage(b3&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(b3&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            win& = FindWindow("_AOL_Modal", "Error Messages")
            txt& = FindWindowEx(win&, 0&, "RICHCNTL", vbNullString)
            ok1& = FindWindowEx(win&, 0&, "_AOL_Icon", vbNullString)
        Loop Until win& <> 0& And txt& <> 0& And ok1& <> 0&
        Call waitfortxttoload(txt&)
    ElseIf getwhat = iPPPINFO Then
        Call PostMessage(b5&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(b5&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            win& = FindWindow("_AOL_Modal", "PPP Status Information")
            txt& = FindWindowEx(win&, 0&, "_AOL_Static", vbNullString)
            ok1& = FindWindowEx(win&, 0&, "_AOL_Icon", vbNullString)
        Loop Until win& <> 0& And txt& <> 0& And ok1& <> 0&
    ElseIf getwhat = iSYSTEMINFO Then
        Call PostMessage(b1&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(b1&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            win& = FindWindow("_AOL_Modal", "System Information")
            txt& = FindWindowEx(win&, 0&, "_AOL_Static", vbNullString)
            ok1& = FindWindowEx(win&, 0&, "_AOL_Icon", vbNullString)
        Loop Until win& <> 0& And txt& <> 0& And ok1& <> 0&
    End If
    getinfo$ = gettext(txt&)
    Call PostMessage(ok1&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ok1&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(ok&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ok&, WM_LBUTTONUP, 0&, 0&)
End Function
Public Function gettreecount(tree As Long) As Long
    gettreecount& = SendMessage(tree&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function gettreeitemindex(tree As Long, item As String, trimspaces As Boolean, lcasetext As Boolean) As Long
    Dim index As Long, indexlength As Long, thestring As String
    For index& = 0& To SendMessage(tree&, LB_GETCOUNT, 0&, 0&) - 1
        indexlength& = SendMessage(tree&, LB_GETTEXTLEN, index&, 0&)
        thestring$ = String(indexlength& + 1, 0)
        Call SendMessageByString(tree&, LB_GETTEXT, index&, thestring$)
        If trimspaces = True Then thestring$ = removechar(thestring$, " ")
        If lcasetext = True Then thestring$ = LCase(thestring$)
        If thestring$ = item$ Then
            gettreeitemindex& = index&
            Exit Function
        End If
    Next index&
End Function
Public Function gettreeitemtext(tree As Long, index As Long) As String
    Dim indexlength As Long
    indexlength& = SendMessage(tree&, LB_GETTEXTLEN, index&, 0&)
    gettreeitemtext$ = String(indexlength& + 1, 0)
    Call SendMessageByString(tree&, LB_GETTEXT, index&, gettreeitemtext$)
End Function
Public Function htmlstring(thestring As String, bold As Boolean, italic As Boolean, underline As Boolean, color As String, font As String) As String
    Dim sbold As String, sitalic As String, sunderline As String
    sbold$ = "</b>"
    sitalic$ = "</i>"
    sunderline$ = "</u>"
    If bold = True Then sbold$ = "<b>"
    If italic = True Then sitalic$ = "<i>"
    If underline = True Then sunderline$ = "<u>"
    If font$ <> "" Then font$ = "face=" & Chr(34) & font$ & Chr(34)
    If color$ <> "" Then
        If Left(color$, 1) <> "#" Then color$ = "#" & color$
        color$ = "color=" & Chr(34) & color$ & Chr(34)
    End If
    htmlstring$ = sbold$ & sitalic$ & sunderline$ & "<font " & font$ & " " & color$ & ">" & thestring$
End Function
Public Sub ignoremail(index As Long)
    Dim maillist As Long, listtab As Long, listtabs As Long
    Dim listwin As Long, ignorebutton As Long
    maillist& = findmaillist(mtNEW)
    listtab& = GetParent(maillist&)
    If listtab& = 0& Then Exit Sub
    listtabs& = GetParent(listtab&)
    listwin& = GetParent(listtabs&)
    Call PostMessage(maillist&, LB_SETCURSEL, CLng(index&), 0&)
    ignorebutton& = FindWindowEx(listwin&, 0&, "_AOL_Icon", vbNullString)
    ignorebutton& = FindWindowEx(listwin&, ignorebutton&, "_AOL_Icon", vbNullString)
    ignorebutton& = FindWindowEx(listwin&, ignorebutton&, "_AOL_Icon", vbNullString)
    ignorebutton& = FindWindowEx(listwin&, ignorebutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(ignorebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ignorebutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub killmailad(mailwin As Long)
    Dim ad As Long
    ad& = FindWindowEx(mailwin&, 0&, "_AOL_Image", vbNullString)
    Call SendMessage(ad&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub killwait()
    Dim modal As Long, okbutton As Long
    Call runmenu(4, 10)
    Do: DoEvents
        modal& = FindWindow("_AOL_Modal", vbNullString)
        okbutton& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until modal& <> 0& And okbutton& <> 0&
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub macroscroll(macro As String, duration As Long, font As String)
    Dim whereat As Long, linesep As String, theline As String, thechar As Integer
    whereat& = 0
    linesep$ = Chr(13)
    Do: DoEvents
        whereat& = InStr(macro$, linesep$)
        If whereat& = 0& Then Exit Do
        theline$ = Left(macro$, whereat& - 1)
        macro$ = Right(macro$, Len(macro$) - Len(theline$) - 2)
        Call sendchat("<font face=" & font & ">" & theline$)
        pause duration&
    Loop Until macro$ = ""
    If macro$ = "" Or macro$ = Chr(13) Or macro$ = Chr(10) Then Exit Sub
    Call sendchat(macro$)
End Sub
Public Sub notifyaolchat(persontotos As String, tosviolation As String)
    Dim aol As Long, mdi As Long, chatwin As Long, notifybutton As Long
    Dim toswin As Long, categorycombo As Long, datetimebox As Long, roomnamebox As Long
    Dim personbox As Long, violationbox As Long, sendbutton1 As Long, sendbutton As Long
    Dim roomname As String, whereat As Long, roomcategory As String, okwin As Long
    Dim okbutton As Long, index As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    chatwin& = findroom()
    notifybutton& = FindWindowEx(chatwin&, 0&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(chatwin&, notifybutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(notifybutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(notifybutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        toswin& = FindWindowEx(mdi&, 0&, "AOL Child", "Notify AOL")
        categorycombo& = FindWindowEx(toswin&, 0&, "_AOL_Combobox", vbNullString)
        datetimebox& = FindWindowEx(toswin&, 0&, "_AOL_Edit", vbNullString)
        roomnamebox& = FindWindowEx(toswin&, datetimebox&, "_AOL_Edit", vbNullString)
        personbox& = FindWindowEx(toswin&, roomnamebox&, "_AOL_Edit", vbNullString)
        violationbox& = FindWindowEx(toswin&, personbox&, "_AOL_Edit", vbNullString)
        sendbutton1& = FindWindowEx(toswin&, 0&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(toswin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(toswin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(toswin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton& = FindWindowEx(toswin&, sendbutton1&, "_AOL_Icon", vbNullString)
    Loop Until toswin& <> 0& And categorycombo& <> 0& And datetimebox& <> 0& And personbox& <> 0& And violationbox& <> 0& And sendbutton& <> 0&
    roomname$ = getchatname()
    whereat& = InStr(roomname$, " - ")
    roomcategory$ = Left(roomname$, whereat& - 1)
    roomname$ = Right(roomname$, Len(roomname$) - Len(roomcategory$) - 3)
    Select Case roomcategory$
        Case "Times Square": index& = 0&
        Case "Arts and Entertainment": index& = 1&
        Case "Friends": index& = 2&
        Case "Life": index& = 3&
        Case "News, Sports, and Finance": index& = 4&
        Case "Places": index& = 5&
        Case "Romance": index& = 6&
        Case "Special Interests": index& = 7&
        Case "Germany": index& = 8&
        Case "UK Experience": index& = 9&
        Case "France": index& = 10&
        Case "Canada": index& = 11&
        Case "Conference Room": index& = 12&
        Case "Kids Room": index& = 13&
        Case "Teen Room": index& = 14&
        Case "Japan": index& = 15&
        Case "Other": index& = 16&
        Case Else: index& = 16&
    End Select
    Call PostMessage(categorycombo&, CB_SETCURSEL, index&, 0&)
    Call SendMessageByString(datetimebox&, WM_SETTEXT, 0&, Now)
    Call SendMessageByString(roomnamebox&, WM_SETTEXT, 0&, roomname$)
    Call SendMessageByString(personbox&, WM_SETTEXT, 0&, persontotos$)
    Call SendMessageByString(violationbox&, WM_SETTEXT, 0&, tosviolation$)
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
    Loop Until okwin& <> 0& And okbutton& <> 0&
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(toswin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub notifyaolim(imwin As Long, toswhat As String)
    Dim aol As Long, mdi As Long, notifybutton As Long, modalwin As Long
    Dim tosbox As Long, sendbutton As Long, thankswin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    notifybutton& = FindWindowEx(imwin&, 0&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    notifybutton& = FindWindowEx(imwin&, notifybutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(notifybutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(notifybutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        modalwin& = FindWindow("_AOL_Modal", "Notify AOL")
        tosbox& = FindWindowEx(modalwin&, 0&, "_AOL_Edit", vbNullString)
        sendbutton& = FindWindowEx(modalwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until modalwin& <> 0& And tosbox& <> 0& And sendbutton& <> 0&
    Call SendMessageByString(tosbox&, WM_SETTEXT, 0&, toswhat$)
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        thankswin& = FindWindowEx(mdi&, 0&, "AOL Child", "Thank you from C.A.T.")
    Loop Until thankswin& <> 0&
    Call SendMessage(thankswin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub openmail(maillist As Long, index As Long, themailtype As MAILTYPE)
    Dim listtab As Long, listtabs As Long, listwin As Long, readbutton As Long
    If themailtype = mtNEW Or themailtype = mtOLD Or themailtype = mtSENT Then
        listtab& = GetParent(maillist&)
        If listtab& = 0& Then Exit Sub
        listtabs& = GetParent(listtab&)
        listwin& = GetParent(listtabs&)
    ElseIf themailtype = mtFLASH Then
        listwin& = GetParent(maillist&)
        If listwin& = 0& Then Exit Sub
    End If
    Call PostMessage(maillist&, LB_SETCURSEL, CLng(index&), 0&)
    readbutton& = FindWindowEx(listwin&, 0&, "_AOL_Icon", vbNullString)
    Call PostMessage(readbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(readbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub openprofile(screenname As String)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, profilewin As Long, snbox As Long, okbutton As Long
    Dim prowin As Long, profile As Long, protxt1 As String, protxt2 As String
    Dim protxt3 As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 11)
    Do: DoEvents
        profilewin& = FindWindowEx(mdi&, 0&, "AOL Child", "Get a Member's Profile")
        snbox& = FindWindowEx(profilewin&, 0&, "_AOL_Edit", vbNullString)
        okbutton& = FindWindowEx(profilewin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until profilewin& <> 0& And snbox& <> 0& And okbutton& <> 0&
    Call SendMessageByString(snbox&, WM_SETTEXT, 0&, screenname$)
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub opensignonwin()
    Dim aol As Long, amenu As Long, smenu As Long, sid As Long
    Dim sstring As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    amenu& = GetMenu(aol&)
    smenu& = GetSubMenu(amenu&, 3&)
    sid& = GetMenuItemID(smenu&, 0&)
    sstring$ = String$(16, " ")
    Call GetMenuString(smenu&, sid&, sstring$, 16&, 1&)
    If Left(sstring$, Len(sstring$) - 1) = "&Sign On Screen" Then
        Call SendMessageLong(aol&, WM_COMMAND, sid&, 0&)
    End If
End Sub
Public Sub playwav(sound As String)
    If Len(Dir(sound$)) Then Call sndPlaySound(sound$, SND_FLAG)
End Sub
Public Sub privateroom(room As String)
    Call keyword("aol://2719:2-2-" & room$)
End Sub
Public Sub removebuddy(screenname As String, group As String)
    Dim aol As Long, mdi As Long, buddywin As Long, setupbutton As Long
    Dim setupwin As Long, grouplist As Long, editbutton As Long, groupindex As Long
    Dim groupwin As Long, addbutton As Long, removebutton As Long, savebutton As Long
    Dim buddylist As Long, addbox1 As Long, buddylistindex As Long, addbox As Long
    Dim grouptext As String, whereat As Long, editbutton1 As Long, errorstatic1 As Long
    Dim errorstatic As Long, screennameindex As Long, statictxt As String
    group$ = LCase(group$)
    group$ = removechar(group$, " ")
    screenname$ = LCase(screenname$)
    screenname$ = removechar(screenname$, " ")
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0& Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        If setupwin& = 0& Then setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy List")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0& And editbutton& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&)
        grouptext$ = getlistitemtext(grouplist&, groupindex&)
        whereat& = InStr(grouptext$, Chr(9))
        grouptext$ = Left(grouptext$, whereat& - 1)
        grouptext$ = LCase(grouptext$)
        grouptext$ = removechar(grouptext$, " ")
        If grouptext$ = group$ Then Exit For
    Next groupindex&
    If groupindex& = -1& Then
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    grouptext$ = getlistitemtext(grouplist&, groupindex&)
    grouptext$ = Left(grouptext$, whereat& - 1)
    Call PostMessage(grouplist&, LB_SETCURSEL, CLng(groupindex&), 0&)
    Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        groupwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & grouptext$)
        addbox1& = FindWindowEx(groupwin&, 0&, "_AOL_Edit", vbNullString)
        addbox& = FindWindowEx(groupwin&, addbox1&, "_AOL_Edit", vbNullString)
        buddylist& = FindWindowEx(groupwin&, 0&, "_AOL_Listbox", vbNullString)
        addbutton& = FindWindowEx(groupwin&, 0&, "_AOL_Icon", vbNullString)
        removebutton& = FindWindowEx(groupwin&, addbutton&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(groupwin&, removebutton&, "_AOL_Icon", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, 0&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
    Loop Until groupwin& <> 0& And addbox& <> 0& And buddylist& <> 0& And addbutton& <> 0& And removebutton& <> 0& And savebutton& <> 0& And errorstatic& <> 0&
    Call waitforlisttoload(buddylist&)
    buddylistindex& = getlistitemindex(buddylist&, screenname$, True, True)
    If buddylistindex& = -1& Then
        Call PostMessage(groupwin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    Call PostMessage(buddylist&, LB_SETCURSEL, CLng(buddylistindex&), 0&)
    Call PostMessage(removebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(removebutton&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub removemultiplebuddies(buddies As String, separator As String, group As String)
    Dim numofbuddies As Long, buddiesnum As Long, whereat As Long, aol As Long
    Dim mdi As Long, setupwin As Long, grouplist As Long, editbutton1 As Long, editbutton As Long
    Dim groupindex As Long, grouptext As String, groupwin As Long, buddylist As Long
    Dim addbox1 As Long, addbox As Long, addbutton As Long, removebutton As Long
    Dim savebutton As Long, errorstatic1 As Long, errorstatic As Long, originalcount As Long
    Dim buddynum As Long, buddy As String, buddythere As Long, addboxtxt As String
    Dim numofoldbuddies As Long, groupnum As String, lbcount1 As String
    group$ = LCase(group$)
    group$ = removechar(group$, " ")
    numofbuddies& = countchar(buddies$, separator$)
    numofoldbuddies& = getbuddycount(False)
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        If setupwin& = 0& Then setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy List")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0& And editbutton& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&)
        grouptext$ = getlistitemtext(grouplist&, groupindex&)
        whereat& = InStr(grouptext$, Chr(9))
        grouptext$ = Left(grouptext$, whereat& - 1)
        grouptext$ = LCase(grouptext$)
        grouptext$ = removechar(grouptext$, " ")
        If grouptext$ = group$ Then Exit For
    Next groupindex&
    If groupindex& = -1& Then
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    grouptext$ = getlistitemtext(grouplist&, groupindex&)
    groupnum$ = Right(grouptext$, Len(grouptext$) - whereat&)
    grouptext$ = Left(grouptext$, whereat& - 1)
    Call PostMessage(grouplist&, LB_SETCURSEL, CLng(groupindex&), 0&)
    Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        groupwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & grouptext$)
        buddylist& = FindWindowEx(groupwin&, 0&, "_AOL_Listbox", vbNullString)
        addbox1& = FindWindowEx(groupwin&, 0&, "_AOL_Edit", vbNullString)
        addbox& = FindWindowEx(groupwin&, addbox1&, "_AOL_Edit", vbNullString)
        addbutton& = FindWindowEx(groupwin&, 0&, "_AOL_Icon", vbNullString)
        removebutton& = FindWindowEx(groupwin&, addbutton&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(groupwin&, removebutton&, "_AOL_Icon", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, 0&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic1& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
        errorstatic& = FindWindowEx(groupwin&, errorstatic1&, "_AOL_Static", vbNullString)
    Loop Until groupwin& <> 0& And buddylist& <> 0& And addbutton& <> 0& And removebutton& <> 0& And savebutton& <> 0& And errorstatic& <> 0& And addbox& <> 0&
    Do: DoEvents
        lbcount1$ = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&)
    Loop Until lbcount1$ = groupnum$
    originalcount& = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&)
    For buddynum& = 0& To numofbuddies& - 1
        whereat& = 0&
        Do: DoEvents
            whereat& = InStr(whereat& + 1, buddies$, separator$)
            If whereat& = 0& Then Exit For
            buddy$ = Left(buddies$, whereat& - Len(separator$) + 1)
            buddies$ = Right(buddies$, Len(buddies$) - Len(buddy$) - Len(separator$))
            buddy$ = LCase(buddy$)
            buddy$ = removechar(buddy$, " ")
            buddythere& = getlistitemindex(buddylist&, buddy$, True, True)
            If buddythere& <> -1& Then Exit Do
        Loop Until buddythere& <> -1&
        Call PostMessage(buddylist&, LB_SETCURSEL, CLng(buddythere&), 0&)
        Call PostMessage(removebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(removebutton&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            addboxtxt$ = gettext(addbox&)
        Loop Until addboxtxt$ = buddy$
    Next buddynum&
    If originalcount& = SendMessage(buddylist&, LB_GETCOUNT, 0&, 0&) Then
        Call PostMessage(groupwin&, WM_CLOSE, 0&, 0&)
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
    End If
    If buddies$ = "" Or buddies$ = separator$ Then
        Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    buddythere& = getlistitemindex(buddylist&, buddies$, True, True)
    If buddythere& = -1& Then
        Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
        Call waitforok
        Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
        Exit Sub
    End If
    Call PostMessage(buddylist&, LB_SETCURSEL, CLng(buddythere&), 0&)
    Call PostMessage(removebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(removebutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        addboxtxt$ = gettext(addbox&)
    Loop Until addboxtxt$ = buddies$
    Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)

End Sub
Public Sub renamebuddygroup(group As String, newgroup As String)
    Dim group2 As String, aol As Long, mdi As Long, buddywin As Long
    Dim setupbutton As Long, setupwin As Long, grouplist As Long, editbutton1 As Long
    Dim editbutton As Long, groupindex As Long, groupname As String, whereat As Long
    Dim editwin As Long, namebox As Long, savebutton1 As Long, savebutton As Long
    group2$ = LCase(group$)
    group2$ = removechar(group2$, " ")
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0& Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0& Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And grouplist& <> 0& And editbutton& <> 0&
    For groupindex& = 0& To SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&) - 1
        groupname$ = getlistitemtext(grouplist&, groupindex&)
        groupname$ = LCase(groupname$)
        groupname$ = removechar(groupname$, " ")
        whereat& = InStr(groupname$, Chr(9))
        groupname$ = Left(groupname$, whereat& - 1)
        If groupname$ = group2$ Then Exit For
    Next groupindex&
    If groupindex& = -1& Then Exit Sub
    groupname$ = getlistitemtext(grouplist&, groupindex&)
    groupname$ = Left(groupname$, whereat& - 1)
    Call PostMessage(grouplist&, LB_SETCURSEL, CLng(groupindex&), 0&)
    Call PostMessage(editbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(editbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        editwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit List " & groupname$)
        namebox& = FindWindowEx(editwin&, 0&, "_AOL_Edit", vbNullString)
        savebutton1& = FindWindowEx(editwin&, 0&, "_AOL_Icon", vbNullString)
        savebutton1& = FindWindowEx(editwin&, savebutton1&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(editwin&, savebutton1&, "_AOL_Icon", vbNullString)
    Loop Until editwin& <> 0& And namebox& <> 0& And savebutton& <> 0&
    Call SendMessageByString(namebox&, WM_SETTEXT, 0&, newgroup$)
    Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function reversestring(thestring As String) As String
    Dim char As Long, first As String, newstring As String
    For char& = 1& To Len(thestring$)
        first$ = Mid(thestring$, char&, 1)
        newstring$ = first$ & newstring$
    Next char&
End Function
Public Sub roombustfast(room As String)
    Dim aol As Long, mdi As Long, roomname As String, okwin As Long
    Dim okbutton As Long, chatwin As Long, toolbar1 As Long, toolbar2 As Long
    Dim combo As Long, editwin As Long, modal As Long, button As Long
    Dim chatwintxt As String, formwin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    roombustcount& = 0&
    roombuststop = False
    roomname$ = gettext(findroom())
    roomname$ = removechar(roomname$, " ")
    roomname$ = LCase(roomname$)
    room$ = removechar(room$, " ")
    room$ = LCase(room$)
    If roomname$ = room$ Then Exit Sub
    Do: DoEvents
        If roombuststop = True Then Exit Do
        toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
        toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
        combo& = FindWindowEx(toolbar2&, 0&, "_AOL_Combobox", vbNullString)
        editwin& = FindWindowEx(combo&, 0&, "Edit", vbNullString)
        Call SendMessageByString(editwin&, WM_SETTEXT, 0&, "aol://2719:2-2-" & room$)
        Call SendMessageLong(editwin&, WM_CHAR, VK_SPACE, 0&)
        Call SendMessageLong(editwin&, WM_CHAR, VK_RETURN, 0&)
        okwin& = FindWindow("#32770", "America Online")
        If okwin& <> 0& Then
            roombustcount& = roombustcount& + 1
            okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
            Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
        chatwin& = findroom()
        chatwintxt$ = gettext(chatwin&)
        chatwintxt$ = LCase(chatwintxt$)
        chatwintxt$ = removechar(chatwintxt$, " ")
        formwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Form")
    Loop Until chatwintxt$ = room$ Or formwin& <> 0&
    Call runmenu(4, 10)
    Do: DoEvents
        modal& = FindWindow("_AOL_Modal", vbNullString)
        button& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until modal& <> 0& And button& <> 0&
    Do: DoEvents
        Call PostMessage(button&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(button&, WM_LBUTTONUP, 0&, 0&)
        button& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
    Loop Until button& = 0&
End Sub
Public Sub roombustslow(room As String)
Dim aol As Long, mdi As Long, roomname As String, okwin As Long
    Dim okbutton As Long, chatwin As Long, toolbar1 As Long, toolbar2 As Long
    Dim combo As Long, editwin As Long, chatwintxt As String, formwin As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    roombustcount& = 0&
    roombuststop = False
    roomname$ = gettext(findroom())
    roomname$ = removechar(roomname$, " ")
    roomname$ = LCase(roomname$)
    room$ = removechar(room$, " ")
    room$ = LCase(room$)
    If roomname$ = room$ Then Exit Sub
    Do: DoEvents
        If roombuststop = True Then Exit Do
        toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
        toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
        combo& = FindWindowEx(toolbar2&, 0&, "_AOL_Combobox", vbNullString)
        editwin& = FindWindowEx(combo&, 0&, "Edit", vbNullString)
        Call SendMessageByString(editwin&, WM_SETTEXT, 0&, "aol://2719:2-2-" & room$)
        Call SendMessageLong(editwin&, WM_CHAR, VK_SPACE, 0&)
        Call SendMessageLong(editwin&, WM_CHAR, VK_RETURN, 0&)
        Do: DoEvents
            okwin& = FindWindow("#32770", "America Online")
            chatwin& = findroom()
            chatwintxt$ = gettext(chatwin&)
            chatwintxt$ = LCase(chatwintxt$)
            chatwintxt$ = removechar(chatwintxt$, " ")
            formwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Form")
        Loop Until chatwintxt$ = room$ Or okwin& <> 0& Or formwin& <> 0&
        If okwin& <> 0& Then
            roombustcount& = roombustcount& + 1
            Do: DoEvents
                okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
                Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
            Loop Until okbutton& = 0&
        End If
    Loop Until chatwintxt$ = room$ Or formwin& <> 0&
End Sub
Public Sub roomrun(room As String)
    Dim aol As Long, mdi As Long, roomruncount As Long, roomname As String
    Dim toolbar1 As Long, toolbar2 As Long, combo As Long, editwin As Long
    Dim okwin As Long, chatwin As Long, okbutton As Long, roomrunna As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    roomruncount& = 0&
    roomrunstop = False
    roomname$ = gettext(findroom())
    roomname$ = removechar(roomname$, " ")
    roomname$ = LCase(roomname$)
    Do: DoEvents
        If roomrunstop = True Then Exit Sub
        toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
        toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
        combo& = FindWindowEx(toolbar2&, 0&, "_AOL_Combobox", vbNullString)
        editwin& = FindWindowEx(combo&, 0&, "Edit", vbNullString)
        If room$ & roomruncount& = roomname$ Then roomruncount& = roomruncount& + 1
        roomrunna$ = "aol://2719:2-2-" & room$ & roomruncount&
        If roomruncount& = 0& Then roomrunna$ = "aol://2719:2-2-" & room$
        Call SendMessageByString(editwin&, WM_SETTEXT, 0&, roomrunna$)
        Call SendMessageLong(editwin&, WM_CHAR, VK_SPACE, 0&)
        Call SendMessageLong(editwin&, WM_CHAR, VK_RETURN, 0&)
        Do: DoEvents
            okwin& = FindWindow("#32770", "America Online")
            chatwin& = FindWindowEx(mdi&, 0&, "AOL Child", room$ & roomruncount&)
        Loop Until chatwin& <> 0& Or okwin& <> 0&
        If okwin& <> 0& Then
            roomruncount& = roomruncount& + 1
            okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
            Do: DoEvents
                Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
                Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
                okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
            Loop Until okbutton& = 0&
        End If
    Loop Until chatwin& <> 0&
End Sub
Public Sub runflashsession()
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, setupwin As Long, checkbox1 As Long, checkbox2 As Long
    Dim checkbox3 As Long, checkbox4 As Long, checkbox5 As Long, checkbox6 As Long
    Dim runicon1 As Long, runicon As Long, runwin As Long, signoff As Long
    Dim beginbutton As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 10)
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Automatic AOL")
        checkbox1& = FindWindowEx(setupwin&, 0&, "_AOL_Checkbox", vbNullString)
        checkbox2& = FindWindowEx(setupwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox3& = FindWindowEx(setupwin&, checkbox2&, "_AOL_Checkbox", vbNullString)
        checkbox4& = FindWindowEx(setupwin&, checkbox3&, "_AOL_Checkbox", vbNullString)
        checkbox5& = FindWindowEx(setupwin&, checkbox4&, "_AOL_Checkbox", vbNullString)
        checkbox6& = FindWindowEx(setupwin&, checkbox5&, "_AOL_Checkbox", vbNullString)
        runicon1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        runicon& = FindWindowEx(setupwin&, runicon1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And checkbox6& <> 0& And runicon& <> 0&
    Call PostMessage(checkbox1&, BM_SETCHECK, False, 0&)
    Call PostMessage(checkbox2&, BM_SETCHECK, True, 0&)
    Call PostMessage(checkbox3&, BM_SETCHECK, False, 0&)
    Call PostMessage(checkbox4&, BM_SETCHECK, False, 0&)
    Call PostMessage(checkbox5&, BM_SETCHECK, False, 0&)
    Call PostMessage(checkbox6&, BM_SETCHECK, False, 0&)
    Call PostMessage(runicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(runicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        runwin& = FindWindow("_AOL_Modal", "Run Automatic AOL Now")
        signoff& = FindWindowEx(runwin&, 0&, "_AOL_Checkbox", vbNullString)
        beginbutton& = FindWindowEx(runwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until runwin& <> 0& And signoff& <> 0& And beginbutton& <> 0&
    Call SendMessage(signoff&, BM_SETCHECK, False, 0&)
    Call PostMessage(beginbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(beginbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub searchmemberdir(searchfor As String, onlyon As Boolean)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, dirwin As Long, searchbox As Long, checkbox1 As Long
    Dim checkbox As Long, searchbutton1 As Long, searchbutton As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 9&)
    Do: DoEvents
        dirwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Member Directory")
        searchbox& = FindWindowEx(dirwin&, 0&, "_AOL_Edit", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, 0&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox1& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        checkbox& = FindWindowEx(dirwin&, checkbox1&, "_AOL_Checkbox", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, 0&, "_AOL_Icon", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
        searchbutton1& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
        searchbutton& = FindWindowEx(dirwin&, searchbutton1&, "_AOL_Icon", vbNullString)
    Loop Until dirwin& <> 0& And searchbox& <> 0& And checkbox& <> 0& And searchbutton& <> 0&
    Call SendMessageByString(searchbox&, WM_SETTEXT, 0&, searchfor$)
    If onlyon = True Then
        Call PostMessage(checkbox&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(checkbox&, WM_LBUTTONUP, 0&, 0&)
    End If
    Call PostMessage(searchbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(searchbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub

Public Sub sendmailattachment(person As String, subject As String, message As String, filepath As String)
    Dim slashcount As Long, indexx As Long, whereat As Long, whereattemp As Long
    Dim whereat1 As Long, folders As String, folder As String, file As String
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long, writeicon As Long
    Dim writewin As Long, sendtobox As Long, subjectbox1 As Long, subjectbox As Long
    Dim messagebox As Long, attach1 As Long, sendbutton1 As Long, sendbutton2 As Long
    Dim sendbutton As Long, attachwin As Long, attachbutton As Long, okbutton1 As Long
    Dim okbutton As Long, openwin As Long, filebox As Long, combo As Long
    Dim openbutton1 As Long, openbutton As Long
    slashcount& = countchar(filepath$, "\")
    For indexx& = 0& To slashcount& - 1
        whereat& = InStr(whereat& + 1, filepath$, "\")
    Next indexx&
    Do: DoEvents
        whereattemp& = InStr(whereattemp& + 1, filepath$, "\")
        If whereattemp& = whereat& Then Exit Do
        whereat1& = whereattemp&
    Loop
    folders$ = Left(filepath$, whereat&)
    folder$ = Mid(filepath$, whereat1& + 1, whereat& - whereat1& - 1)
    file$ = Right(filepath$, Len(filepath$) - whereat&)
    If Len(Dir(filepath$)) Then
        aol& = FindWindow("AOL Frame25", vbNullString)
        mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
        toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
        toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
        writeicon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
        writeicon& = FindWindowEx(toolbar2&, writeicon&, "_AOL_Icon", vbNullString)
        Call PostMessage(writeicon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(writeicon&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            writewin& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
            sendtobox& = FindWindowEx(writewin&, 0&, "_AOL_Edit", vbNullString)
            subjectbox1& = FindWindowEx(writewin&, sendtobox&, "_AOL_Edit", vbNullString)
            subjectbox& = FindWindowEx(writewin&, subjectbox1&, "_AOL_Edit", vbNullString)
            messagebox& = FindWindowEx(writewin&, 0&, "RICHCNTL", vbNullString)
            attach1& = FindWindowEx(writewin&, 0&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            attach1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            sendbutton1& = FindWindowEx(writewin&, attach1&, "_AOL_Icon", vbNullString)
            sendbutton2& = FindWindowEx(writewin&, sendbutton1&, "_AOL_Icon", vbNullString)
            sendbutton& = FindWindowEx(writewin&, sendbutton2&, "_AOL_Icon", vbNullString)
        Loop Until writewin& <> 0& And sendtobox& <> 0& And subjectbox1& <> 0& And subjectbox& <> 0& And messagebox& <> 0& And sendbutton& <> 0& And sendbutton& <> sendbutton1& And sendbutton1& <> sendbutton2& And sendbutton2& <> 0& And sendbutton1 <> 0&
        Call SendMessageByString(sendtobox&, WM_SETTEXT, 0&, person$)
        Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, subject$)
        Call SendMessageByString(messagebox&, WM_SETTEXT, 0&, message$)
        Call PostMessage(sendbutton2&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(sendbutton2&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            attachwin& = FindWindow("_AOL_Modal", "Attachments")
            attachbutton& = FindWindowEx(attachwin&, 0&, "_AOL_Icon", vbNullString)
            okbutton1& = FindWindowEx(attachwin&, attachbutton&, "_AOL_Icon", vbNullString)
            okbutton& = FindWindowEx(attachwin&, okbutton1&, "_AOL_Icon", vbNullString)
        Loop Until attachwin& <> 0& And attachbutton& <> 0& And okbutton& <> 0&
        Call PostMessage(attachbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(attachbutton&, WM_LBUTTONUP, 0&, 0&)
        Do: DoEvents
            openwin& = FindWindow("#32770", "Attach")
            filebox& = FindWindowEx(openwin&, 0&, "Edit", vbNullString)
            combo& = FindWindowEx(openwin&, 0&, "ComboBox", vbNullString)
            openbutton1& = FindWindowEx(openwin&, 0&, "Button", vbNullString)
            openbutton& = FindWindowEx(openwin&, openbutton1&, "Button", vbNullString)
        Loop Until openwin& <> 0& And filebox& <> 0& And combo& <> 0& And openbutton& <> 0&
        Call SendMessageByString(filebox&, WM_SETTEXT, 0&, folders$)
        Do: DoEvents
            Call SendMessageByString(filebox&, WM_SETTEXT, 0&, folders$)
            Call PostMessage(openbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(openbutton&, WM_LBUTTONUP, 0&, 0&)
        Loop Until LCase(gettext(combo&)) = LCase(folder$)
        Call SendMessageByString(filebox&, WM_SETTEXT, 0&, file$)
        Do: DoEvents
            Call PostMessage(openbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(openbutton&, WM_LBUTTONUP, 0&, 0&)
            openwin& = FindWindow("#32770", "Browse")
        Loop Until openwin& = 0&
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Else
        Exit Sub
    End If
End Sub
Public Sub sendmobilecommpage(pagerid As String, message As String)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, choicewin As Long, picon1 As Long, picon As Long
    Dim sendwin As Long, idbox As Long, messagebox As Long, sendicon As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 8&)
    Do: DoEvents
        choicewin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Page")
        picon1& = FindWindowEx(choicewin&, 0&, "_AOL_Icon", vbNullString)
        picon1& = FindWindowEx(choicewin&, picon1&, "_AOL_Icon", vbNullString)
        picon1& = FindWindowEx(choicewin&, picon1&, "_AOL_Icon", vbNullString)
        picon& = FindWindowEx(choicewin&, picon1&, "_AOL_Icon", vbNullString)
    Loop Until choicewin& <> 0& And picon& <> 0&
    Call PostMessage(picon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(picon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        sendwin& = FindWindowEx(mdi&, 0&, "AOL Child", UCase(getuser()) & "'s Send a Page")
        idbox& = FindWindowEx(sendwin&, 0&, "_AOL_Edit", vbNullString)
        messagebox& = FindWindowEx(sendwin&, idbox&, "_AOL_Edit", vbNullString)
        sendicon& = FindWindowEx(sendwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until sendwin& <> 0& And idbox& <> 0& And messagebox& <> 0& And sendicon& <> 0&
    Call SendMessageByString(idbox&, WM_SETTEXT, 0&, pagerid$)
    Call SendMessageByString(messagebox&, WM_SETTEXT, 0&, message$)
    Call PostMessage(sendicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendicon&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(sendwin&, WM_CLOSE, 0&, 0&)
    Call SendMessage(choicewin&, WM_CLOSE, 0&, 0&)
End Sub

Public Sub setgeneralprefs(showchannels As Boolean, eventsound As Boolean, chatsound As Boolean)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long, icon As Long
    Dim prefwin As Long, genicon As Long, genwin As Long, channelscheck As Long
    Dim eventsoundcheck1 As Long, eventsoundcheck As Long, chatsoundcheck As Long
    Dim okicon As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 3)
    Do: DoEvents
        prefwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Preferences")
        genicon& = FindWindowEx(prefwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until prefwin& <> 0& And genicon& <> 0&
    Call PostMessage(genicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(genicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        genwin& = FindWindow("_AOL_Modal", "General Preferences")
        channelscheck& = FindWindowEx(genwin&, 0&, "_AOL_Checkbox", vbNullString)
        eventsoundcheck1& = FindWindowEx(genwin&, channelscheck&, "_AOL_Checkbox", vbNullString)
        eventsoundcheck1& = FindWindowEx(genwin&, eventsoundcheck1&, "_AOL_Checkbox", vbNullString)
        eventsoundcheck1& = FindWindowEx(genwin&, eventsoundcheck1&, "_AOL_Checkbox", vbNullString)
        eventsoundcheck1& = FindWindowEx(genwin&, eventsoundcheck1&, "_AOL_Checkbox", vbNullString)
        eventsoundcheck1& = FindWindowEx(genwin&, eventsoundcheck1&, "_AOL_Checkbox", vbNullString)
        eventsoundcheck& = FindWindowEx(genwin&, eventsoundcheck1&, "_AOL_Checkbox", vbNullString)
        chatsoundcheck& = FindWindowEx(genwin&, eventsoundcheck&, "_AOL_Checkbox", vbNullString)
        okicon& = FindWindowEx(genwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until genwin& <> 0& And channelscheck& <> 0& And eventsoundcheck& <> 0& And chatsoundcheck& <> 0&
    Call PostMessage(channelscheck&, BM_SETCHECK, showchannels, 0&)
    Call PostMessage(eventsoundcheck&, BM_SETCHECK, eventsound, 0&)
    Call PostMessage(chatsoundcheck&, BM_SETCHECK, chatsound, 0&)
    Call PostMessage(okicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okicon&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(prefwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub setmailprefs()
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim myaolicon As Long, prefwin As Long, mailicon1 As Long, mailicon As Long
    Dim mailprefwin As Long, confirm As Long, closemail As Long, spcheck1 As Long
    Dim spcheck As Long, okicon As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    myaolicon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    myaolicon& = FindWindowEx(toolbar2&, myaolicon&, "_AOL_Icon", vbNullString)
    myaolicon& = FindWindowEx(toolbar2&, myaolicon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(myaolicon&, 7&)
    Do: DoEvents
        mailprefwin& = FindWindow("_AOL_Modal", "Mail Preferences")
        confirm& = FindWindowEx(mailprefwin&, 0&, "_AOL_Checkbox", vbNullString)
        closemail& = FindWindowEx(mailprefwin&, confirm&, "_AOL_Checkbox", vbNullString)
        spcheck1& = FindWindowEx(mailprefwin&, closemail&, "_AOL_Checkbox", vbNullString)
        spcheck1& = FindWindowEx(mailprefwin&, spcheck1&, "_AOL_Checkbox", vbNullString)
        spcheck1& = FindWindowEx(mailprefwin&, spcheck1&, "_AOL_Checkbox", vbNullString)
        spcheck& = FindWindowEx(mailprefwin&, spcheck1&, "_AOL_Checkbox", vbNullString)
        okicon& = FindWindowEx(mailprefwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until mailprefwin& <> 0& And confirm& <> 0& And closemail& <> 0& And spcheck1& <> 0& And spcheck& <> 0& And okicon& <> 0&
    Call PostMessage(confirm&, BM_SETCHECK, False, 0&)
    Call PostMessage(closemail&, BM_SETCHECK, True, 0&)
    Call PostMessage(spcheck&, BM_SETCHECK, False, 0&)
    Call PostMessage(okicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okicon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub formsetaolparent(frm As Form)
    'thanks neo for the idea
    Dim aol As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call SetParent(frm.hwnd, mdi&)
End Sub
Public Sub setprofile(name As String, location As String, birthday As String, yoursex As SEX, maritalstatus As String, hobbies As String, computer As String, occupation As String, quote As String)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim myaol As Long, profilewin As Long, nbox As Long, lbox As Long
    Dim bbox As Long, mbox As Long, hbox As Long, cbox As Long, obox As Long, qbox As Long
    Dim msex As Long, fsex As Long, nsex As Long, updatebutton As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    myaol& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    myaol& = FindWindowEx(toolbar2&, myaol&, "_AOL_Icon", vbNullString)
    myaol& = FindWindowEx(toolbar2&, myaol&, "_AOL_Icon", vbNullString)
    myaol& = FindWindowEx(toolbar2&, myaol&, "_AOL_Icon", vbNullString)
    myaol& = FindWindowEx(toolbar2&, myaol&, "_AOL_Icon", vbNullString)
    myaol& = FindWindowEx(toolbar2&, myaol&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(myaol&, 4)
    Do: DoEvents
        profilewin& = FindWindowEx(mdi&, 0&, "AOL Child", "Edit Your Online Profile")
        nbox& = FindWindowEx(profilewin&, 0&, "_AOL_Edit", vbNullString)
        lbox& = FindWindowEx(profilewin&, nbox&, "_AOL_Edit", vbNullString)
        bbox& = FindWindowEx(profilewin&, lbox&, "_AOL_Edit", vbNullString)
        mbox& = FindWindowEx(profilewin&, bbox&, "_AOL_Edit", vbNullString)
        hbox& = FindWindowEx(profilewin&, mbox&, "_AOL_Edit", vbNullString)
        cbox& = FindWindowEx(profilewin&, hbox&, "_AOL_Edit", vbNullString)
        obox& = FindWindowEx(profilewin&, cbox&, "_AOL_Edit", vbNullString)
        qbox& = FindWindowEx(profilewin&, obox&, "_AOL_Edit", vbNullString)
        msex& = FindWindowEx(profilewin&, 0&, "_AOL_Checkbox", vbNullString)
        fsex& = FindWindowEx(profilewin&, msex&, "_AOL_Checkbox", vbNullString)
        nsex& = FindWindowEx(profilewin&, fsex&, "_AOL_Checkbox", vbNullString)
        updatebutton& = FindWindowEx(profilewin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until profilewin& <> 0& And qbox& <> 0& And nsex& <> 0& And updatebutton& <> 0&
    Call SendMessageByString(nbox&, WM_SETTEXT, 0&, name$)
    Call SendMessageByString(lbox&, WM_SETTEXT, 0&, location$)
    Call SendMessageByString(bbox&, WM_SETTEXT, 0&, birthday$)
    Call SendMessageByString(mbox&, WM_SETTEXT, 0&, maritalstatus$)
    Call SendMessageByString(hbox&, WM_SETTEXT, 0&, hobbies$)
    Call SendMessageByString(cbox&, WM_SETTEXT, 0&, computer$)
    Call SendMessageByString(qbox&, WM_SETTEXT, 0&, quote$)
    If yoursex = sMALE Then
        Call PostMessage(msex&, BM_SETCHECK, True, 0&)
    ElseIf yoursex = sFEMALE Then
        Call PostMessage(fsex&, BM_SETCHECK, True, 0&)
    ElseIf yoursex = sNEITHER Then
        Call PostMessage(nsex&, BM_SETCHECK, True, 0&)
    End If
    Call PostMessage(updatebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(updatebutton&, WM_LBUTTONUP, 0&, 0&)
    Call waitforok
End Sub
Public Sub signoff()
    Call runmenu(3&, 1&)
End Sub
Public Function spacedstring(thestring As String, separator As String) As String
    Dim char As Long, first As String, newstring As String
    For char& = 1& To Len(thestring$)
        first$ = Mid(thestring$, char&, 1)
        newstring$ = newstring$ & separator$ & first$
    Next char&
    spacedstring$ = Right(newstring$, Len(newstring$) - Len(separator$))
End Function
Public Sub pause(duration As Long)
    Dim current As Long
    current = Timer
    Do Until Timer - current >= duration
        DoEvents
    Loop
End Sub
Public Sub openmailbox(mailtype1 As MAILTYPE)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim mailbutton As Long, smod As Long, winvis As Long
    Dim curpos As POINTAPI
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    mailbutton& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    mailbutton& = FindWindowEx(toolbar2&, mailbutton&, "_AOL_Icon", vbNullString)
    mailbutton& = FindWindowEx(toolbar2&, mailbutton&, "_AOL_Icon", vbNullString)
    If mailtype1 = mtFLASH Then
        Call GetCursorPos(curpos)
        Call SetCursorPos(Screen.Width, Screen.Height)
        Call PostMessage(mailbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(mailbutton&, WM_LBUTTONUP, 0&, 0&)
        Do
            smod& = FindWindow("#32768", vbNullString)
            winvis& = IsWindowVisible(smod&)
        Loop Until winvis& = 1
        Call PostMessage(smod&, WM_KEYDOWN, VK_UP, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_UP, 0&)
        Call PostMessage(smod&, WM_KEYDOWN, VK_RIGHT, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_RIGHT, 0&)
        Call PostMessage(smod&, WM_KEYDOWN, VK_RETURN, 0&)
        Call PostMessage(smod&, WM_KEYUP, VK_RETURN, 0&)
        Call SetCursorPos(curpos.x, curpos.Y)
    ElseIf mailtype1 = mtNEW Then
        Call clickaoliconmenu(mailbutton&, 2&)
    ElseIf mailtype1 = mtOLD Then
        Call clickaoliconmenu(mailbutton&, 4&)
    ElseIf mailtype1 = mtSENT Then
        Call clickaoliconmenu(mailbutton&, 5&)
    End If
End Sub

Public Sub addbuddygroupstocontrol(addto As Control)
    Dim aol As Long, mdi As Long, buddylistwin As Long, setupbutton As Long
    Dim setupwin As Long, editbutton1 As Long, editbutton As Long, grouplist As Long
    Dim index As Long, item As String, groupcount As Long, whereat As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddylistwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddylistwin& = 0& Then
        Call keyword("buddy list")
    Else
        setupbutton& = FindWindowEx(buddylistwin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddylistwin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddylistwin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy Lists")
        If setupwin& = 0& Then setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", getuser() & "'s Buddy List")
        editbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        editbutton& = FindWindowEx(setupwin&, editbutton1&, "_AOL_Icon", vbNullString)
        grouplist& = FindWindowEx(setupwin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until setupwin& <> 0& And editbutton& <> 0& And grouplist& <> 0&
    groupcount& = SendMessage(grouplist&, LB_GETCOUNT, 0&, 0&)
    For index& = 0& To groupcount& - 1&
        item$ = getlistitemtext(grouplist&, index&)
        whereat& = InStr(item$, Chr(9))
        item$ = Left(item$, whereat& - 1)
        addto.AddItem item$
    Next index&
    Call SendMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub addroomtocontrol(addto As Control)
    'parts taken from dos's addroomtolist sub
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long, index As Long
    Dim screenname As String, itmhold As Long, psnhold As Long
    Dim rbytes As Long, cprocess As Long, room As Long
    room& = findroom()
    rlist& = FindWindowEx(room&, 0&, "_AOL_Listbox", vbNullString)
    sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
            Call CopyMemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            addto.AddItem screenname$
        Next index&
        Call CloseHandle(mthread&)
    End If
End Sub
Public Sub buddyblock(person As String, onoff As Boolean)
    Dim user As String, aol As Long, mdi As Long
    Dim buddywin As Long, setupbutton As Long, setupwin As Long
    Dim ppbutton1 As Long, ppbutton As Long, ppwin As Long
    Dim blockbox As Long, addbutton1 As Long, addbutton As Long
    Dim savebutton As Long, okwin As Long, okbutton As Long
    Dim removebutton As Long, snlist As Long, whereat As Long
    user$ = getuser()
    person$ = LCase(replacestring(person$, " ", ""))
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0 Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0 Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", user$ & "'s Buddy Lists")
        ppbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And ppbutton& <> 0&
    Call PostMessage(ppbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ppbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        ppwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Privacy Preferences")
        blockbox& = FindWindowEx(ppwin&, 0&, "_AOL_Edit", vbNullString)
        addbutton1& = FindWindowEx(ppwin&, 0&, "_AOL_Icon", vbNullString)
        addbutton& = FindWindowEx(ppwin&, addbutton1&, "_AOL_Icon", vbNullString)
        removebutton& = FindWindowEx(ppwin&, addbutton&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(ppwin&, removebutton&, "_AOL_Icon", vbNullString)
        snlist& = FindWindowEx(ppwin&, 0&, "_AOL_Listbox", vbNullString)
    Loop Until ppwin& <> 0& And blockbox& <> 0& And addbutton& <> 0& And removebutton& <> 0& And savebutton& <> 0& And snlist& <> 0&
    If onoff = True Then
        whereat& = getlistitemindex(snlist&, person$, False, False)
        Call SendMessageByString(blockbox&, WM_SETTEXT, 0&, person$)
        Call PostMessage(addbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(addbutton&, WM_LBUTTONUP, 0&, 0&)
    ElseIf onoff = False Then
        whereat& = getlistitemindex(snlist&, person$, False, False)
        If whereat& <> 0& Then
            Call PostMessage(snlist&, LB_SETCURSEL, CLng(whereat&), 0&)
            Call PostMessage(removebutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(removebutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
    Do: DoEvents
        Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
        ppwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Privacy Preferences")
    Loop Until ppwin& = 0&
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
    Loop Until okwin& <> 0& And okbutton& <> 0&
    Do: DoEvents
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        okwin& = FindWindow("#32770", "America Online")
    Loop Until okwin& = 0&
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub buddylist(openbuddylist As Boolean)
    Dim aol As Long, mdi As Long, buddywin As Long
    If openbuddylist = True Then
        Call keyword("buddy view")
    Else
        aol& = FindWindow("AOL Frame25", "America  Online")
        mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
        buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
        Call SendMessage(buddywin&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Function checkifavailible(person As String) As Boolean
    Dim aol As Long, mdi As Long, imwin As Long, availible1 As Long
    Dim availible As Long, okwin As Long, okbutton As Long
    Dim okmsg1a As Long, okmsg1 As Long, okmsg As String, tobox As Long
    Dim availible2 As Long, imtext As Long, static1 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://9293:" & person$)
    Do: DoEvents
        imwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
        imtext& = FindWindowEx(imwin&, 0&, "RICHCNTL", vbNullString)
        tobox& = FindWindowEx(imwin&, 0&, "_AOL_Edit", vbNullString)
        static1& = FindWindowEx(imwin&, 0&, "_AOL_Static", "To:")
        availible1& = FindWindowEx(imwin&, 0&, "_AOL_Icon", vbNullString)
        availible1& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible1& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible1& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible1& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible1& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible1& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible1& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible2& = FindWindowEx(imwin&, availible1&, "_AOL_Icon", vbNullString)
        availible& = FindWindowEx(imwin&, availible2&, "_AOL_Icon", vbNullString)
    Loop Until imwin& <> 0& And availible& <> 0& And imtext& <> 0& And tobox& <> 0& And static1& <> 0& And availible1& <> availible2& And availible2& <> availible& And availible1& <> 0& And availible2& <> 0&
    Call PostMessage(availible&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(availible&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
        okmsg1a& = FindWindowEx(okwin&, 0&, "Static", vbNullString)
        okmsg1& = FindWindowEx(okwin&, okmsg1a&, "Static", vbNullString)
    Loop Until okwin& <> 0& And okbutton& <> 0& And okmsg1& <> 0&
    okmsg$ = gettext(okmsg1&)
    Do: DoEvents
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
    Loop Until okwin& = 0& And okbutton& = 0&
    Call SendMessage(imwin&, WM_CLOSE, 0&, 0&)
    If LCase(okmsg$) = LCase(person$ & " is online and able to receive Instant Messages.") Then
        checkifavailible = True
    Else
        checkifavailible = False
    End If
End Function
Public Sub clearchatwin()
    Dim chatwin As Long, chattxt As Long
    chatwin& = findroom()
    chattxt& = FindWindowEx(chatwin&, 0&, "RICHCNTL", vbNullString)
    Call SendMessage(chattxt&, WM_CLEAR, 0&, 0&)
    Call SendMessageByString(chattxt&, WM_SETTEXT, 0&, "")
End Sub
Public Sub closewin(win As Long)
    Call SendMessage(win&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub dblclickicon(icon As Long)
    Call PostMessage(icon&, WM_LBUTTONDBLCLK, 0&, 0&)
End Sub
Public Function findaimim() As Long
    Dim aol As Long, mdi As Long, childwin As Long
    Dim childwincaption As String, aimimwin As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    childwin& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    childwincaption$ = getcaption(childwin&)
    If InStr(childwincaption$, "Instant Message From: ") <> 0& Then
        aimimwin& = childwin&
    Else
        Do: DoEvents
            childwin& = FindWindowEx(mdi&, childwin&, "AOL Child", vbNullString)
            childwincaption$ = getcaption(childwin&)
            If InStr(childwincaption$, "Instant Message From: ") <> 0& Then
                aimimwin& = childwin&
                Exit Do
            End If
        Loop Until childwin& = 0&
    End If
    findaimim& = childwin&
End Function
Public Function findbuddylist() As Long
    Dim aol As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    findbuddylist& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
End Function
Public Function findim() As Long
    Dim aol As Long, mdi As Long, childwin As Long
    Dim childwincaption As String, imwin As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    childwin& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    childwincaption$ = getcaption(childwin&)
    If LCase(childwincaption$) Like LCase(">Instant Message From: *") Or LCase(childwincaption$) Like LCase("  Instant Message From: *") Or LCase(childwincaption$) Like LCase("  Instant Message To: *") Then
        imwin& = childwin&
    Else
        Do: DoEvents
            childwin& = FindWindowEx(mdi&, childwin&, "AOL Child", vbNullString)
            childwincaption$ = getcaption(childwin&)
            If LCase(childwincaption$) Like LCase(">Instant Message From: *") Or LCase(childwincaption$) Like LCase("  Instant Message From: *") Or LCase(childwincaption$) Like LCase("  Instant Message To: *") Then
                imwin& = childwin&
                Exit Do
            End If
        Loop Until childwin& = 0&
    End If
    findim& = childwin&
End Function
Public Function getlistitemindex(list As Long, item As String, trimspaces As Boolean, lcasetext As Boolean) As Long
    'parts of this were taken from dos's addroomtolistbox sub
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long, index As Long
    Dim screenname As String, itmhold As Long, psnhold As Long
    Dim rbytes As Long, cprocess As Long
    rlist& = list&
    sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        For index& = 0& To SendMessage(rlist&, LB_GETCOUNT, 0&, 0&) - 1
            screenname$ = String$(4, vbNullChar)
            itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
            itmhold& = itmhold& + 24
            Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
            Call CopyMemory(psnhold&, ByVal screenname$, 4)
            psnhold& = psnhold& + 6
            screenname$ = String$(16, vbNullChar)
            Call ReadProcessMemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
            screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
            If trimspaces = True Then
                screenname$ = replacestring(screenname$, " ", "")
                screenname$ = replacestring(screenname$, " ", "")
                screenname$ = replacestring(screenname$, " ", "")
            End If
            If lcasetext = True Then screenname$ = LCase(screenname$)
            If screenname$ = item$ Then
                getlistitemindex& = index&
                Call CloseHandle(mthread&)
                Exit Function
            End If
        Next index&
        getlistitemindex& = -1&
        Call CloseHandle(mthread&)
    End If
End Function
Public Function findroom() As Long
    Dim aol As Long, mdi As Long, child As Long, rich2 As Long
    Dim Rich As Long, aollist As Long, aolicon2 As Long, aolicon3 As Long
    Dim aolicon As Long, aolstatic As Long, aolicon4 As Long, aolicon5 As Long
    Dim aolicon6 As Long, aolicon7 As Long, aolicon8 As Long, aolicon9 As Long
    Dim aolicon10 As Long, aolicon11 As Long, aolicon12 As Long, aolicon13 As Long
    Dim aolcombobox As Long, aolglyph As Long, aolstatic2 As Long, aolimage As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
    rich2& = FindWindowEx(child&, Rich&, "RICHCNTL", vbNullString)
    aollist& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
    aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
    aolicon2& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
    aolicon3& = FindWindowEx(child&, aolicon2&, "_AOL_Icon", vbNullString)
    aolicon4& = FindWindowEx(child&, aolicon3&, "_AOL_Icon", vbNullString)
    aolicon5& = FindWindowEx(child&, aolicon4&, "_AOL_Icon", vbNullString)
    aolicon6& = FindWindowEx(child&, aolicon5&, "_AOL_Icon", vbNullString)
    aolicon7& = FindWindowEx(child&, aolicon6&, "_AOL_Icon", vbNullString)
    aolicon8& = FindWindowEx(child&, aolicon7&, "_AOL_Icon", vbNullString)
    aolicon9& = FindWindowEx(child&, aolicon8&, "_AOL_Icon", vbNullString)
    aolicon10& = FindWindowEx(child&, aolicon9&, "_AOL_Icon", vbNullString)
    aolicon11& = FindWindowEx(child&, aolicon10&, "_AOL_Icon", vbNullString)
    aolicon12& = FindWindowEx(child&, aolicon11&, "_AOL_Icon", vbNullString)
    aolicon13& = FindWindowEx(child&, aolicon12&, "_AOL_Icon", vbNullString)
    aolcombobox& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
    aolstatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
    aolstatic2& = FindWindowEx(child&, aolstatic&, "_AOL_Static", vbNullString)
    aolglyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
    aolimage& = FindWindowEx(child&, 0&, "_AOL_Image", vbNullString)
    If aolcombobox& <> 0& And Rich& <> 0& And rich2& <> 0& And aolglyph& <> 0& And aollist& <> 0& And aolicon& <> 0& And aolicon2& <> 0& And aolicon3& <> 0& And aolicon4& <> 0& And aolicon5& <> 0& And aolicon6& <> 0& And aolicon7& <> 0& And aolicon8& <> 0& And aolicon9& <> 0& And aolicon10& <> 0& And aolicon11& <> 0& And aolicon12& <> 0& And aolicon13& <> 0& And aolstatic& <> 0& And aolstatic2& <> 0& Then
        findroom& = child&
    Else
        Do: DoEvents
            child& = FindWindowEx(mdi&, child&, "AOL Child", vbNullString)
            Rich& = FindWindowEx(child&, 0&, "RICHCNTL", vbNullString)
            rich2& = FindWindowEx(child&, Rich&, "RICHCNTL", vbNullString)
            aollist& = FindWindowEx(child&, 0&, "_AOL_Listbox", vbNullString)
            aolicon& = FindWindowEx(child&, 0&, "_AOL_Icon", vbNullString)
            aolicon2& = FindWindowEx(child&, aolicon&, "_AOL_Icon", vbNullString)
            aolicon3& = FindWindowEx(child&, aolicon2&, "_AOL_Icon", vbNullString)
            aolicon4& = FindWindowEx(child&, aolicon3&, "_AOL_Icon", vbNullString)
            aolicon5& = FindWindowEx(child&, aolicon4&, "_AOL_Icon", vbNullString)
            aolicon6& = FindWindowEx(child&, aolicon5&, "_AOL_Icon", vbNullString)
            aolicon7& = FindWindowEx(child&, aolicon6&, "_AOL_Icon", vbNullString)
            aolicon8& = FindWindowEx(child&, aolicon7&, "_AOL_Icon", vbNullString)
            aolicon9& = FindWindowEx(child&, aolicon8&, "_AOL_Icon", vbNullString)
            aolicon10& = FindWindowEx(child&, aolicon9&, "_AOL_Icon", vbNullString)
            aolicon11& = FindWindowEx(child&, aolicon10&, "_AOL_Icon", vbNullString)
            aolicon12& = FindWindowEx(child&, aolicon11&, "_AOL_Icon", vbNullString)
            aolicon13& = FindWindowEx(child&, aolicon12&, "_AOL_Icon", vbNullString)
            aolcombobox& = FindWindowEx(child&, 0&, "_AOL_Combobox", vbNullString)
            aolstatic& = FindWindowEx(child&, 0&, "_AOL_Static", vbNullString)
            aolstatic2& = FindWindowEx(child&, aolstatic&, "_AOL_Static", vbNullString)
            aolglyph& = FindWindowEx(child&, 0&, "_AOL_Glyph", vbNullString)
            aolimage& = FindWindowEx(child&, 0&, "_AOL_Image", vbNullString)
            If aolcombobox& <> 0& And Rich& <> 0& And rich2& <> 0& And aolglyph& <> 0& And aollist& <> 0& And aolicon& <> 0& And aolicon2& <> 0& And aolicon3& <> 0& And aolicon4& <> 0& And aolicon5& <> 0& And aolicon6& <> 0& And aolicon7& <> 0& And aolicon8& <> 0& And aolicon9& <> 0& And aolicon10& <> 0& And aolicon11& <> 0& And aolicon12& <> 0& And aolicon13& <> 0& And aolstatic& <> 0& And aolstatic2& <> 0& Then
                findroom& = child&
                Exit Function
            End If
        Loop Until child& = 0&
        Exit Function
    End If
    findroom& = child&
End Function
Public Sub clickicon(icon As Long)
    Call PostMessage(icon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(icon&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Function getlistitemtext(list As Long, index As Long) As String
    'parts of this were taken from dos's addroomtolistbox sub
    On Error Resume Next
    Dim rlist As Long, sthread As Long, mthread As Long
    Dim screenname As String, itmhold As Long, psnhold As Long
    Dim rbytes As Long, cprocess As Long
    rlist& = list&
    sthread& = GetWindowThreadProcessId(rlist&, cprocess&)
    mthread& = OpenProcess(PROCESS_READ Or RIGHTS_REQUIRED, False, cprocess&)
    If mthread& Then
        screenname$ = String$(4, vbNullChar)
        itmhold& = SendMessage(rlist, LB_GETITEMDATA, ByVal CLng(index&), ByVal 0&)
        itmhold& = itmhold& + 24
        Call ReadProcessMemory(mthread&, itmhold&, screenname$, 4, rbytes&)
        Call CopyMemory(psnhold&, ByVal screenname$, 4)
        psnhold& = psnhold& + 6
        screenname$ = String$(16, vbNullChar)
        Call ReadProcessMemory(mthread&, psnhold&, screenname$, Len(screenname$), rbytes&)
        screenname$ = Left$(screenname$, InStr(screenname$, vbNullChar) - 1)
        getlistitemtext$ = screenname$
        Call CloseHandle(mthread&)
    End If
End Function
Public Sub hidewin(win As Long, onoff As Boolean)
    If onoff = True Then
        Call ShowWindow(win&, SW_HIDE)
    ElseIf onoff = False Then
        Call ShowWindow(win&, SW_SHOW)
    End If
End Sub
Public Sub idlebot()
    Dim palette As Long, modal As Long, palettebutton As Long, modalbutton As Long
    idlebotstop = False
    Do: DoEvents
        If idlebotstop = True Then Exit Sub
        palette& = FindWindow("_AOL_Palette", vbNullString)
        modal& = FindWindow("_AOL_Modal", vbNullString)
        palettebutton& = FindWindowEx(palette&, 0&, "_AOL_Icon", vbNullString)
        modalbutton& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
        If palettebutton& <> 0& Or modalbutton& <> 0& Then
            Call PostMessage(palettebutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(palettebutton&, WM_LBUTTONUP, 0&, 0&)
            Call PostMessage(modalbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(modalbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    Loop
End Sub
Public Sub runmenu(topmenu As Long, submenu As Long)
    'taken from dos32.bas
    Dim aol As Long, amenu As Long, smenu As Long, mnid As Long
    Dim mval As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    amenu& = GetMenu(aol&)
    smenu& = GetSubMenu(amenu&, topmenu&)
    mnid& = GetMenuItemID(smenu&, submenu&)
    Call SendMessageLong(aol&, WM_COMMAND, mnid&, 0&)
End Sub
Public Sub runmenubystring(searchstring As String)
    'taken from dos32.bas
    Dim aol As Long, amenu As Long, mcount As Long
    Dim lookfor As Long, smenu As Long, scount As Long
    Dim looksub As Long, sid As Long, sstring As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    amenu& = GetMenu(aol&)
    mcount& = GetMenuItemCount(amenu&)
    For lookfor& = 0& To mcount& - 1
        smenu& = GetSubMenu(amenu&, lookfor&)
        scount& = GetMenuItemCount(smenu&)
        For looksub& = 0 To scount& - 1
            sid& = GetMenuItemID(smenu&, looksub&)
            sstring$ = String$(100, " ")
            Call GetMenuString(smenu&, sid&, sstring$, 100&, 1&)
            If InStr(LCase(sstring$), LCase(searchstring$)) Then
                Call SendMessageLong(aol&, WM_COMMAND, sid&, 0&)
                Exit Sub
            End If
        Next looksub&
    Next lookfor&
End Sub
Public Sub aimaccept(accept As Boolean)
    Dim aol As Long, mdi As Long
    Dim aimimwin As Long, okbutton As Long, nobutton As Long
    Dim okwin As Long, okwbutton As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    aimimwin& = findaimim()
    If aimimwin& <> 0& Then
        okbutton& = FindWindowEx(aimimwin&, 0&, "_AOL_Icon", vbNullString)
        nobutton& = FindWindowEx(aimimwin&, okbutton&, "_AOL_Icon", vbNullString)
        If accept = True Then
            Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        Else
            Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
            Do: DoEvents
                okwin& = FindWindow("#32770", "America Online")
                okwbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
            Loop Until okwin& <> 0& And okwbutton& <> 0&
            Call PostMessage(okwbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okwbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
End Sub
Public Function imfrom(imwin As Long) As String
    Dim imwincaption As String, whereat As Long
    imwincaption$ = getcaption(imwin&)
    whereat& = InStr(imwincaption$, ":")
    imfrom$ = Right(imwincaption$, Len(imwincaption$) - whereat& - 1)
End Function
Public Sub imson(onoff As Boolean)
    If onoff = True Then
        Call sendim("$im_on", "GpX says hi ;)")
    ElseIf onoff = False Then
        Call sendim("$im_off", "GpX says hi ;)")
    End If
End Sub
Public Sub keyword(kw As String)
    Dim aol As Long, toolbar1 As Long, toolbar2 As Long
    Dim combo As Long, comboedit As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    combo& = FindWindowEx(toolbar2&, 0&, "_AOL_Combobox", vbNullString)
    comboedit& = FindWindowEx(combo&, 0&, "Edit", vbNullString)
    Call SendMessageByString(comboedit&, WM_SETTEXT, 0&, kw$)
    Call SendMessageLong(comboedit&, WM_CHAR, VK_SPACE, 0&)
    Call SendMessageLong(comboedit&, WM_CHAR, VK_RETURN, 0&)
End Sub
Public Function lastchatline() As String
    Dim chatwin As Long, chattxt1 As Long, chattxt As String
    Dim enter1 As Long, enter2 As Long
    chatwin& = findroom()
    chattxt1& = FindWindowEx(chatwin&, 0&, "RICHCNTL", vbNullString)
    chattxt$ = gettext(chattxt1&)
    enter1& = InStr(chattxt$, Chr(13))
    Do: DoEvents
        If enter1& <> 0& Then enter2& = enter1&
        enter1& = InStr(enter1& + 1, chattxt$, Chr(13))
    Loop Until enter1& = 0&
    lastchatline$ = Right(chattxt$, Len(chattxt$) - enter2&)
End Function
Public Function gettext(windowhandle As Long) As String
    'thanks dos
    'www.hider.com/dos
    Dim buffer As String, textlength As Long
    textlength& = SendMessage(windowhandle&, WM_GETTEXTLENGTH, 0&, 0&)
    buffer$ = String(textlength&, 0&)
    Call SendMessageByString(windowhandle&, WM_GETTEXT, textlength& + 1, buffer$)
    gettext$ = buffer$
End Function
Public Function lastchatlinemsg() As String
    Dim lastline As String, tab1 As Long
    lastline$ = lastchatline()
    tab1& = InStr(lastline$, Chr(9))
    lastchatlinemsg$ = Right(lastline$, Len(lastline$) - tab1&)
End Function
Public Function lastchatlinesn() As String
    Dim lastline As String, tab1 As Long
    lastline$ = lastchatline()
    tab1& = InStr(lastline$, Chr(9))
    lastchatlinesn$ = Left(lastline$, tab1& - 3)
End Function
Public Function lastim(imwin As Long) As String
    Dim aol As Long, mdi As Long, childwin As Long
    Dim childwincaption As String, enter2 As Long
    Dim imtxt1 As Long, imtxt As String, enter1 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    imtxt1& = FindWindowEx(imwin&, 0&, "RICHCNTL", vbNullString)
    imtxt$ = gettext(imtxt1&)
    enter1& = InStr(imtxt$, Chr(13))
    Do: DoEvents
        If enter1& <> 0& Then enter2& = enter1&
        enter1& = InStr(enter1& + 1, imtxt$, Chr(13))
    Loop Until enter1& = 0&
    lastim$ = Mid(imtxt$, enter2& + 3, Len(imtxt$) - enter2& - 3)
End Function
Public Function lastimmsg(imwin As Long) As String
    Dim lastimtxt As String, tab1 As Long
    lastimtxt$ = lastim(imwin&)
    tab1& = InStr(lastimtxt$, Chr(9))
    lastimmsg$ = Mid(lastimtxt$, tab1& + 2, Len(lastimtxt$) - tab1& + 1)
End Function
Public Function lastimsn(imwin As Long) As String
    Dim lastimtxt As String, tab1 As Long
    lastimtxt$ = lastim(imwin&)
    tab1& = InStr(lastimtxt$, Chr(9))
    lastimsn$ = Left(lastimtxt$, tab1& - 2)
End Function
Public Sub getfontlist(list As Control)
    Dim x As Long
    list.Clear
    For x = 0 To Screen.FontCount - 1
            list.AddItem Screen.Fonts(x)
    Next x
End Sub
Public Function locatemember(person As String) As String
    Dim aol As Long, mdi As Long, okwin As Long
    Dim childwin As Long, childwincaption As String
    Dim locatewin As Long, locatemsg1 As Long, locatemsg As String
    Dim okbutton As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://3548:" & person$)
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        If okwin& <> 0& Then Exit Do
        childwin& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
        childwincaption$ = getcaption(childwin&)
        If LCase(childwincaption$) = LCase("locate " & person$) Then
            locatewin& = childwin&
        Else
            Do: DoEvents
                childwin& = FindWindowEx(mdi&, childwin&, "AOL Child", vbNullString)
                childwincaption$ = getcaption(childwin&)
                If LCase(childwincaption$) = LCase("locate " & person$) Then
                    locatewin& = childwin&
                    Exit Do
                End If
                okwin& = FindWindow("#32770", "America Online")
            Loop Until childwin& = 0& Or okwin& <> 0&
        End If
    Loop Until locatewin& <> 0& Or okwin& <> 0&
    If locatewin& <> 0& Then
        locatemsg1& = FindWindowEx(locatewin&, 0&, "_AOL_Static", vbNullString)
        locatemsg$ = gettext(locatemsg1&)
        If LCase(locatemsg$) = LCase(person$ & " is online, but not in a chat area.") Then
            locatemember$ = "Not in a chat."
        ElseIf LCase(locatemsg$) = LCase(person$ & " is online, but in a private room.") Then
            locatemember$ = "Private room."
        ElseIf LCase(locatemsg$) Like LCase(person$ & " is in chat room *") Then
            locatemember$ = Right(locatemsg$, Len(locatemsg$) - Len(person$ & " is in chat room "))
        End If
        Call SendMessage(locatewin&, WM_CLOSE, 0&, 0&)
    ElseIf okwin& <> 0& Then
        okbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
        Do: DoEvents
            Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
            okwin& = FindWindow("#32770", "America Online")
            okbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
        Loop Until okwin& = 0& And okbutton& = 0&
        locatemember$ = "Not signed on."
    End If
End Function
Public Sub formontop(formname As Form, onoff As Boolean)
    If onoff = True Then
        Call SetWindowPos(formname.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
    Else
        Call SetWindowPos(formname.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
    End If
End Sub
Public Function removehtml(thestring As String, returns As Boolean) As String
    Dim roomtext As String, takeout As String, replacewith As String
    Dim whereat As Long, lefttext As String, righttext As String
    Dim takeout2 As String, whereat1 As Long, whereat2 As Long
    Dim takeout3 As String, takeout4 As String
    roomtext$ = thestring$
    If returns = True Then
        takeout$ = "<br>"
        takeout2$ = "<Br>"
        takeout3$ = "<bR>"
        takeout4$ = "<BR>"
        replacewith$ = Chr(13) & Chr(10)
        whereat& = 0&
        Do
            whereat& = InStr(whereat& + 1, roomtext$, takeout$)
            If whereat& = 0& Then
                whereat& = InStr(whereat& + 1, roomtext$, takeout2$)
                If whereat& = 0& Then
                    whereat& = InStr(whereat& + 1, roomtext$, takeout3$)
                    If whereat& = 0& Then
                        whereat& = InStr(whereat& + 1, roomtext$, takeout4$)
                        If whereat& = 0& Then
                            Exit Do
                        End If
                    End If
                End If
            End If
            lefttext$ = Left(roomtext$, whereat& - 1)
            righttext$ = Mid(roomtext$, whereat& + 4, Len(roomtext$))
            roomtext$ = lefttext$ & replacewith$ & righttext$
        Loop
    End If
    takeout$ = "<"
    takeout2$ = ">"
    whereat& = 0&
    whereat1& = 0&
    whereat2& = 0&
    Do
        whereat1& = InStr(whereat1& + 1, roomtext$, takeout$)
        If whereat1& = 0& Then Exit Do
        whereat2& = InStr(whereat2& + 1, roomtext$, takeout2$)
        whereat& = whereat2& - whereat1&
        lefttext$ = Left(roomtext$, whereat1& - 1)
        righttext$ = Mid(roomtext$, whereat2& + 1, Len(roomtext$) - whereat& + 1)
        roomtext$ = lefttext$ & righttext$
        whereat& = 0&
        whereat1& = 0&
        whereat2& = 0&
    Loop
    removehtml$ = Left(roomtext$, Len(roomtext$) - 2)
End Function
Public Function replacestring(thestring As String, takeout As String, replacewith As String) As String
    Dim tempstring As String, whereat As Long, lefttext As String
    Dim righttext As String
    tempstring$ = thestring$
    whereat& = 0&
    Do
        whereat& = InStr(whereat& + 1, tempstring$, takeout$)
        If whereat& = 0& Then Exit Do
        lefttext$ = Left(tempstring$, whereat& - 1)
        righttext$ = Mid(tempstring$, whereat& + Len(takeout$), Len(tempstring$))
        tempstring$ = lefttext$ & replacewith$ & righttext$
    Loop
    replacestring$ = tempstring$
End Function
Public Sub resethistorycombo()
    Dim aol As Long, toolbar1 As Long, toolbar2 As Long
    Dim combo As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    combo& = FindWindowEx(toolbar2&, 0&, "_AOL_Combobox", vbNullString)
    Call SendMessage(combo&, CB_RESETCONTENT, 0&, 0&)
End Sub
Public Sub imrespond(respondwith As String, acceptaim As Boolean)
    Dim aol As Long, mdi As Long, aimimwin As Long, yesbutton As Long
    Dim nobutton As Long, okwin As Long, okbutton As Long
    Dim childwin As Long, childwincaption As String, whereat As Long
    Dim imwin As Long, imtxt As Long, sendbutton As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    aimimwin& = findaimim()
    If aimimwin& <> 0& Then
        yesbutton& = FindWindowEx(aimimwin&, 0&, "_AOL_Icon", vbNullString)
        nobutton& = FindWindowEx(aimimwin&, yesbutton&, "_AOL_Icon", vbNullString)
        If acceptaim = True Then
            Call PostMessage(yesbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(yesbutton&, WM_LBUTTONUP, 0&, 0&)
        Else
            Call PostMessage(nobutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(nobutton&, WM_LBUTTONUP, 0&, 0&)
            Do: DoEvents
                okwin& = FindWindow("#32770", "America Online")
                okbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
            Loop Until okwin& <> 0& And okbutton& <> 0&
            Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        End If
    End If
    childwin& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    childwincaption$ = getcaption(childwin&)
    whereat& = InStr(childwincaption$, "Instant Message")
    If whereat& = 0& Then
        Do: DoEvents
            childwin& = FindWindowEx(mdi&, childwin&, "AOL Child", vbNullString)
            childwincaption$ = getcaption(childwin&)
            whereat& = InStr(childwincaption$, "Instant Message")
            If whereat& <> 0& Then
                imwin& = childwin&
                Exit Do
            End If
        Loop Until childwin& = 0&
    Else
        imwin& = childwin&
    End If
    imtxt& = FindWindowEx(imwin&, 0&, "RICHCNTL", vbNullString)
    imtxt& = FindWindowEx(imwin&, imtxt&, "RICHCNTL", vbNullString)
    sendbutton& = FindWindowEx(imwin&, 0&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(imwin&, sendbutton&, "_AOL_Icon", vbNullString)
    Call SendMessageByString(imtxt&, WM_SETTEXT, 0&, respondwith$)
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Call SendMessage(imwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub selectcomboitem(combo As Long, index As Long)
    Call PostMessage(combo&, CB_SETCURSEL, CLng(index&), 0&)
End Sub
Public Sub selectlistitem(list As Long, index As Long)
    Call PostMessage(list&, LB_SETCURSEL, CLng(index&), 0&)
End Sub
Public Sub sendchat(message As String)
    Dim chatwin As Long, chatbox As Long, sendbutton As Long
    Dim temptext As String, txt As String
    chatwin& = findroom()
    chatbox& = FindWindowEx(chatwin&, 0&, "RICHCNTL", vbNullString)
    chatbox& = FindWindowEx(chatwin&, chatbox&, "RICHCNTL", vbNullString)
    sendbutton& = FindWindowEx(chatwin&, 0&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(chatwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(chatwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(chatwin&, sendbutton&, "_AOL_Icon", vbNullString)
    sendbutton& = FindWindowEx(chatwin&, sendbutton&, "_AOL_Icon", vbNullString)
    temptext$ = gettext(chatbox&)
    Call SendMessageLong(chatbox&, EM_SETSEL, 0&, Len(temptext$))
    Call SendMessageByString(chatbox&, EM_REPLACESEL, 0&, message$)
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        txt$ = gettext(chatbox&)
    Loop Until txt$ = ""
    Call SendMessageByString(chatbox&, WM_SETTEXT, 0&, temptext$)
End Sub
Public Sub sendim(person As String, message As String)
    Dim aol As Long, mdi As Long, imwin As Long, messagebox As Long
    Dim sendbutton As Long, sendbutton1 As Long, ok As Long, okbutton As Long
    Dim okwin As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call keyword("aol://9293:" & person$)
    Do: DoEvents
        imwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
        messagebox& = FindWindowEx(imwin&, 0&, "RICHCNTL", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, 0&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton& = FindWindowEx(imwin&, sendbutton1&, "_AOL_Icon", vbNullString)
    Loop Until imwin& <> 0& And messagebox& <> 0& And sendbutton& <> 0&
    Call SendMessageByString(messagebox&, WM_SETTEXT, 0&, message$)
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        imwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Send Instant Message")
    Loop Until okwin& <> 0& Or imwin& = 0&
    If okwin& <> 0& Then
        okbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        Call PostMessage(imwin&, WM_CLOSE, 0&, 0&)
    End If
End Sub
Public Sub sendmail(person As String, subject As String, message As String)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim writeicon As Long, mailwin As Long, personbox As Long, ccbox As Long
    Dim subjectbox As Long, messagebox As Long, sendbutton As Long, sendbutton1 As Long
    Dim sendbutton2 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    writeicon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    writeicon& = FindWindowEx(toolbar2&, writeicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(writeicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(writeicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
        personbox& = FindWindowEx(mailwin&, 0&, "_AOL_Edit", vbNullString)
        ccbox& = FindWindowEx(mailwin&, personbox&, "_AOL_Edit", vbNullString)
        subjectbox& = FindWindowEx(mailwin&, ccbox&, "_AOL_Edit", vbNullString)
        messagebox& = FindWindowEx(mailwin&, 0&, "RICHCNTL", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, 0&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton2& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton& = FindWindowEx(mailwin&, sendbutton2&, "_AOL_Icon", vbNullString)
    Loop Until mailwin& <> 0& And personbox& <> 0& And ccbox& <> 0& And subjectbox& <> 0& And messagebox& <> 0& And sendbutton& <> 0& And sendbutton& <> sendbutton1& And sendbutton1& <> sendbutton2& And sendbutton2& <> 0& And sendbutton1 <> 0&
    Call SendMessageByString(personbox&, WM_SETTEXT, 0&, person$)
    Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, subject$)
    Call SendMessageByString(messagebox&, WM_SETTEXT, 0&, message$)
    Do: DoEvents
        Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
    Loop Until mailwin& = 0&
End Sub
Public Sub ghost(onoff As Boolean)
    Dim aol As Long, mdi As Long, buddywin As Long
    Dim setupbutton As Long, setupwin As Long, ppbutton1 As Long
    Dim ppbutton As Long, ppwin As Long, blockall1 As Long
    Dim blockalloff As Long, blockallon As Long, blockiandb1 As Long
    Dim blockiandb As Long, savebutton1 As Long, savebutton As Long
    Dim okwin As Long, okbutton As Long, user As String
    user$ = getuser()
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    buddywin& = FindWindowEx(mdi&, 0&, "AOL Child", "Buddy List Window")
    If buddywin& = 0 Then
        Call keyword("buddy list")
    ElseIf buddywin& <> 0 Then
        setupbutton& = FindWindowEx(buddywin&, 0&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        setupbutton& = FindWindowEx(buddywin&, setupbutton&, "_AOL_Icon", vbNullString)
        Call PostMessage(setupbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(setupbutton&, WM_LBUTTONUP, 0&, 0&)
    End If
    Do: DoEvents
        setupwin& = FindWindowEx(mdi&, 0&, "AOL Child", user$ & "'s Buddy Lists")
        ppbutton1& = FindWindowEx(setupwin&, 0&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton1& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
        ppbutton& = FindWindowEx(setupwin&, ppbutton1&, "_AOL_Icon", vbNullString)
    Loop Until setupwin& <> 0& And ppbutton& <> 0&
    Call PostMessage(ppbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(ppbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        ppwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Privacy Preferences")
        blockall1& = FindWindowEx(ppwin&, 0&, "_AOL_Checkbox", vbNullString)
        blockall1& = FindWindowEx(ppwin&, blockall1&, "_AOL_Checkbox", vbNullString)
        blockalloff& = FindWindowEx(ppwin&, blockall1&, "_AOL_Checkbox", vbNullString)
        blockall1& = FindWindowEx(ppwin&, blockalloff&, "_AOL_Checkbox", vbNullString)
        blockallon& = FindWindowEx(ppwin&, blockall1&, "_AOL_Checkbox", vbNullString)
        blockiandb1& = FindWindowEx(ppwin&, blockallon&, "_AOL_Checkbox", vbNullString)
        blockiandb& = FindWindowEx(ppwin&, blockiandb1&, "_AOL_Checkbox", vbNullString)
        savebutton1& = FindWindowEx(ppwin&, 0&, "_AOL_Icon", vbNullString)
        savebutton1& = FindWindowEx(ppwin&, savebutton1&, "_AOL_Icon", vbNullString)
        savebutton1& = FindWindowEx(ppwin&, savebutton1&, "_AOL_Icon", vbNullString)
        savebutton& = FindWindowEx(ppwin&, savebutton1&, "_AOL_Icon", vbNullString)
    Loop Until ppwin& <> 0& And blockallon& <> 0& And blockalloff& <> 0& And blockiandb& <> 0& And savebutton& <> 0&
    If onoff = True Then
        Call PostMessage(blockallon&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(blockallon&, WM_LBUTTONUP, 0&, 0&)
    ElseIf onoff = False Then
        Call PostMessage(blockalloff&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(blockalloff&, WM_LBUTTONUP, 0&, 0&)
    End If
    Call PostMessage(blockiandb&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(blockiandb&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(savebutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
    Loop Until okwin& <> 0& And okbutton& <> 0&
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(setupwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Function controltostring(list As Control, separator As String) As String
    Dim thestring As String, index As Long
    For index& = 0& To list.listcount - 1
        thestring$ = thestring$ & list.list(index&) & separator$
    Next index&
    thestring$ = Left(thestring$, Len(thestring$) - Len(separator$))
    controltostring$ = thestring$
End Function
Public Function getuser() As String
    Dim aol As Long, mdi As Long, childwin As Long
    Dim childwintxt As String
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    childwin& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
    childwintxt$ = getcaption(childwin&)
    Do
        If InStr(childwintxt$, "Welcome, ") Then Exit Do
        childwin& = FindWindowEx(mdi&, childwin&, "AOL Child", vbNullString)
        childwintxt$ = getcaption(childwin&)
    Loop Until childwin& = 0&
    If childwin& = 0& Then
        getuser$ = "Not Online"
    ElseIf childwin& <> 0& Then
        getuser$ = Mid(childwintxt$, 10, InStr(childwintxt$, "!") - 10)
    End If
End Function
Public Function getcaption(windowhandle As Long) As String
    'thanks dos
    'www.hider.com/dos
    Dim buffer As String, textlength As Long
    textlength& = GetWindowTextLength(windowhandle&)
    buffer$ = String(textlength&, 0&)
    Call GetWindowText(windowhandle&, buffer$, textlength& + 1)
    getcaption$ = buffer$
End Function
Public Sub sendmailbcc(people As String, subject As String, message As String)
    Call sendmail("(" & getuser() & "," & people$ & ")", subject$, message$)
End Sub
Public Sub sendmailcc(people As String, subject As String, message As String)
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim writeicon As Long, mailwin As Long, personbox As Long, ccbox As Long
    Dim subjectbox As Long, messagebox As Long, sendbutton As Long, sendbutton1 As Long
    Dim sendbutton2 As Long
    aol& = FindWindow("AOL Frame25", "America  Online")
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    writeicon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    writeicon& = FindWindowEx(toolbar2&, writeicon&, "_AOL_Icon", vbNullString)
    Call PostMessage(writeicon&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(writeicon&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        mailwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Write Mail")
        personbox& = FindWindowEx(mailwin&, 0&, "_AOL_Edit", vbNullString)
        ccbox& = FindWindowEx(mailwin&, personbox&, "_AOL_Edit", vbNullString)
        subjectbox& = FindWindowEx(mailwin&, ccbox&, "_AOL_Edit", vbNullString)
        messagebox& = FindWindowEx(mailwin&, 0&, "RICHCNTL", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, 0&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton1& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton2& = FindWindowEx(mailwin&, sendbutton1&, "_AOL_Icon", vbNullString)
        sendbutton& = FindWindowEx(mailwin&, sendbutton2&, "_AOL_Icon", vbNullString)
    Loop Until mailwin& <> 0& And personbox& <> 0& And ccbox& <> 0& And subjectbox& <> 0& And messagebox& <> 0& And sendbutton& <> 0& And sendbutton& <> sendbutton1& And sendbutton1& <> sendbutton2& And sendbutton2& <> 0& And sendbutton1 <> 0&
    Call SendMessageByString(personbox&, WM_SETTEXT, 0&, getuser())
    Call SendMessageByString(ccbox&, WM_SETTEXT, 0&, people$)
    Call SendMessageByString(subjectbox&, WM_SETTEXT, 0&, subject$)
    Call SendMessageByString(messagebox&, WM_SETTEXT, 0&, message$)
    Call PostMessage(sendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sendbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub setchatprefs(setnotify As Boolean, setsound As Boolean)
    Dim room As Long, prefbutton As Long, chatprefwin As Long, okbutton As Long
    Dim notifya As Long, notifyb As Long, sound1 As Long, sound As Long, aol As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    room& = findroom()
    prefbutton& = FindWindowEx(room&, 0&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    prefbutton& = FindWindowEx(room&, prefbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(prefbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(prefbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        chatprefwin& = FindWindowEx(aol&, 0&, "_AOL_Modal", "Chat Preferences")
        notifya& = FindWindowEx(chatprefwin&, 0&, "_AOL_Checkbox", vbNullString)
        notifyb& = FindWindowEx(chatprefwin&, notifya&, "_AOL_Checkbox", vbNullString)
        sound1& = FindWindowEx(chatprefwin&, notifyb&, "_AOL_Checkbox", vbNullString)
        sound1& = FindWindowEx(chatprefwin&, sound1&, "_AOL_Checkbox", vbNullString)
        sound& = FindWindowEx(chatprefwin&, sound1&, "_AOL_Checkbox", vbNullString)
        okbutton& = FindWindowEx(chatprefwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until chatprefwin& <> 0& And notifya& <> 0& And notifyb& <> 0& And sound& <> 0& And okbutton& <> 0&
    Call PostMessage(notifya&, BM_SETCHECK, setnotify, 0&)
    Call PostMessage(notifyb&, BM_SETCHECK, setnotify, 0&)
    Call PostMessage(sound&, BM_SETCHECK, setsound, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub settext(win As Long, txt As String)
    Call SendMessageByString(win&, WM_SETTEXT, 0&, txt$)
End Sub
Public Function signedon() As Boolean
    Dim aol As Long, amenu As Long, smenu As Long, sid As Long, sstring As String
    aol& = FindWindow("AOL Frame25", vbNullString)
    amenu& = GetMenu(aol&)
    smenu& = GetSubMenu(amenu&, 3&)
    sid& = GetMenuItemID(smenu&, 0&)
    sstring$ = String$(16, " ")
    Call GetMenuString(smenu&, sid&, sstring$, 16&, 1&)
    If Left(sstring$, Len(sstring$) - 1) = "&Sign On Screen" Then signedon = False
End Function
Public Function htmllink(url As String, text As String) As String
    htmllink$ = "<a href=" & Chr(34) & url$ & Chr(34) & ">" & text$ & "</a>"
End Function
Public Function removechar(thestring As String, char As String) As String
    removechar$ = replacestring(thestring$, char$, "")
End Function
Public Sub signon(screenname As String, password As String)
    Dim aol As Long, mdi As Long, signonwin As Long, namelist As Long
    Dim passwordbox As Long, signonbutton As Long, snindex As Long
    Dim index As Long, indexlength As Long, thestring As String
    Dim rlist As Long, sthread As Long, mthread As Long, itmhold As Long
    Dim psnhold As Long, rbytes As Long, cprocess As Long, combocount As Long
    Dim guestwin As Long, guestsn As Long, guestpw As Long, guestok As Long
    Dim child As Long, signonbutton1 As Long, errorwin As Long, okbutton As Long
    Dim guestcancel As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    signonwin& = findsignonwin()
    If signonwin& = 0& Then
        Do: DoEvents
            child& = FindWindowEx(mdi&, 0&, "AOL Child", vbNullString)
            Call SendMessage(child&, WM_CLOSE, 0&, 0&)
        Loop Until child& = 0&
        Call runmenu(3&, 0&)
    End If
    Do: DoEvents
        signonwin& = findsignonwin()
        namelist& = FindWindowEx(signonwin&, 0&, "_AOL_Combobox", vbNullString)
        passwordbox& = FindWindowEx(signonwin&, 0&, "_AOL_Edit", vbNullString)
        signonbutton1& = FindWindowEx(signonwin&, 0&, "_AOL_Icon", vbNullString)
        signonbutton1& = FindWindowEx(signonwin&, signonbutton1&, "_AOL_Icon", vbNullString)
        signonbutton1& = FindWindowEx(signonwin&, signonbutton1&, "_AOL_Icon", vbNullString)
        signonbutton& = FindWindowEx(signonwin&, signonbutton1&, "_AOL_Icon", vbNullString)
    Loop Until signonwin& <> 0& And namelist& <> 0& And passwordbox& <> 0& And signonbutton& <> 0&
    combocount& = SendMessage(namelist&, CB_GETCOUNT, 0&, 0&)
    Call PostMessage(namelist&, CB_SETCURSEL, combocount& - 1, 0&)
    Call PostMessage(signonbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(signonbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        guestwin& = FindWindow("_AOL_Modal", vbNullString)
        guestsn& = FindWindowEx(guestwin&, 0&, "_AOL_Edit", vbNullString)
        guestpw& = FindWindowEx(guestwin&, guestsn&, "_AOL_Edit", vbNullString)
        guestok& = FindWindowEx(guestwin&, 0&, "_AOL_Icon", vbNullString)
        guestcancel& = FindWindowEx(guestwin&, guestok&, "_AOL_Icon", vbNullString)
    Loop Until guestwin& <> 0& And guestsn& <> 0& And guestpw& <> 0& And guestok& <> 0&
    Call SendMessageByString(guestsn&, WM_SETTEXT, 0&, screenname$)
    Call SendMessageByString(guestpw&, WM_SETTEXT, 0&, password$)
    Call PostMessage(guestok&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(guestok&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        guestwin& = FindWindow("_AOL_Modal", vbNullString)
        errorwin& = FindWindow("#32770", "America Online")
    Loop Until guestwin& = 0& Or errorwin& <> 0&
    If errorwin& <> 0& Then
        okbutton& = FindWindowEx(errorwin&, 0&, "Button", vbNullString)
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        Call PostMessage(guestcancel&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(guestcancel&, WM_LBUTTONUP, 0&, 0&)
    End If
End Sub
Public Sub startdownloadlater()
    Dim aol As Long, mdi As Long, toolbar1 As Long, toolbar2 As Long
    Dim icon As Long, dlwin As Long, dlbutton1 As Long, dlbutton As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    toolbar1& = FindWindowEx(aol&, 0&, "AOL Toolbar", vbNullString)
    toolbar2& = FindWindowEx(toolbar1&, 0&, "_AOL_Toolbar", vbNullString)
    icon& = FindWindowEx(toolbar2&, 0&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    icon& = FindWindowEx(toolbar2&, icon&, "_AOL_Icon", vbNullString)
    Call clickaoliconmenu(icon&, 4&)
    Do: DoEvents
        dlwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Download Manager")
        dlbutton1& = FindWindowEx(dlwin&, 0&, "_AOL_Icon", vbNullString)
        dlbutton& = FindWindowEx(dlwin&, dlbutton1&, "_AOL_Icon", vbNullString)
    Loop Until dlwin& <> 0& And dlbutton& <> 0&
    Call PostMessage(dlbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(dlbutton&, WM_LBUTTONUP, 0&, 0&)
    Call PostMessage(dlwin&, WM_CLOSE, 0&, 0&)
End Sub
Public Sub switchscreenname(index As Long, password As String)
    Dim aol As Long, mdi As Long, switchwin As Long, snlist As Long
    Dim switchbutton As Long, listcount As Long
    Dim sthread As Long, mthread As Long, thesn As String, itmhold As Long
    Dim psnhold As Long, rbytes As Long, swin2 As Long, sok2 As Long
    Dim swin As Long, sok As Long, spw As Long, cprocess As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call runmenu(3&, 0&)
    Do: DoEvents
        switchwin& = FindWindowEx(mdi&, 0&, "AOL Child", "Switch Screen Names")
        snlist& = FindWindowEx(switchwin&, 0&, "_AOL_Listbox", vbNullString)
        switchbutton& = FindWindowEx(switchwin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until switchwin& <> 0& And snlist& <> 0& And switchbutton& <> 0&
    listcount& = SendMessage(snlist&, LB_GETCOUNT, 0&, 0&)
    Call PostMessage(snlist&, LB_SETCURSEL, CLng(index&), 0&)
    Call PostMessage(switchbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(switchbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        swin2& = FindWindow("_AOL_Modal", "Switch Screen Name")
        sok2& = FindWindowEx(swin2&, 0&, "_AOL_Icon", vbNullString)
    Loop Until swin2& <> 0& And sok2& <> 0&
    Call PostMessage(sok2&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sok2&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        swin& = FindWindow("_AOL_Modal", "Switch Screen Name")
        spw& = FindWindowEx(swin&, 0&, "_AOL_Edit", vbNullString)
        sok& = FindWindowEx(swin&, 0&, "_AOL_Icon", vbNullString)
    Loop Until swin& <> 0& And spw& <> 0& And sok& <> 0&
    Call SendMessageByString(spw&, WM_SETTEXT, 0&, password$)
    Call PostMessage(sok&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(sok&, WM_LBUTTONUP, 0&, 0&)
End Sub
Public Sub unsendmail(index As Long)
    Dim maillist As Long, listtab As Long, listtabs As Long, listwin As Long
    Dim unsendbutton As Long, modal As Long, yesbutton1 As Long
    Dim yesbutton As Long, okwin As Long, okbutton As Long
    maillist& = findmaillist(mtSENT)
    listtab& = GetParent(maillist&)
    If listtab& = 0& Then Exit Sub
    listtabs& = GetParent(listtab&)
    listwin& = GetParent(listtabs&)
    Call PostMessage(maillist&, LB_SETCURSEL, CLng(index&), 0&)
    unsendbutton& = FindWindowEx(listwin&, 0&, "_AOL_Icon", vbNullString)
    unsendbutton& = FindWindowEx(listwin&, unsendbutton&, "_AOL_Icon", vbNullString)
    unsendbutton& = FindWindowEx(listwin&, unsendbutton&, "_AOL_Icon", vbNullString)
    unsendbutton& = FindWindowEx(listwin&, unsendbutton&, "_AOL_Icon", vbNullString)
    unsendbutton& = FindWindowEx(listwin&, unsendbutton&, "_AOL_Icon", vbNullString)
    Call PostMessage(unsendbutton&, WM_LBUTTONDOWN, 0&, 0&)
    Call PostMessage(unsendbutton&, WM_LBUTTONUP, 0&, 0&)
    Do: DoEvents
        modal& = FindWindow("_AOL_Modal", vbNullString)
        yesbutton1& = FindWindowEx(modal&, 0&, "_AOL_Icon", vbNullString)
        yesbutton& = FindWindowEx(modal&, yesbutton1&, "_AOL_Icon", vbNullString)
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
    Loop Until modal& <> 0& And yesbutton& <> 0& Or okwin& <> 0& And okbutton& <> 0&
    If modal& <> 0& And yesbutton& <> 0& Then
        Call PostMessage(yesbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(yesbutton&, WM_LBUTTONUP, 0&, 0&)
    ElseIf okwin& <> 0& And okbutton& <> 0& Then
        Do: DoEvents
            Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
            Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
            okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
        Loop Until okbutton& = 0&
    End If
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
    Loop Until okwin& <> 0& And okbutton& <> 0&
    Do: DoEvents
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        okbutton& = FindWindowEx(okwin&, 0&, "Button", vbNullString)
    Loop Until okbutton& = 0&
End Sub


Public Sub waitforlisttoload(list As Long)
    'thanks for the idea izekial
    Dim getcount1 As Long, getcount2 As Long, getcount3 As Long
    Do
        getcount1& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
        pause 0.8
        getcount2& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
        pause 0.8
        getcount3& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
    Loop Until getcount1& = getcount2& And getcount2& = getcount3&
End Sub
Public Sub waitforok()
    Dim okwin As Long, okbutton As Long
    Do: DoEvents
        okwin& = FindWindow("#32770", "America Online")
        okbutton& = FindWindowEx(okwin&, 0&, "Button", "OK")
        Exit Sub
    Loop Until okwin& <> 0& And okbutton& <> 0&
    Do
        Call PostMessage(okbutton&, WM_LBUTTONDOWN, 0&, 0&)
        Call PostMessage(okbutton&, WM_LBUTTONUP, 0&, 0&)
        okwin& = FindWindow("#32770", "America Online")
    Loop Until okwin& = 0&
End Sub
Public Function getstringlineindex(thestring As String, theline As String, removespaces As Boolean, lcasetext As Boolean) As Long
    Dim tempstring As String, whereat As Long, index As Long, linestring As String, linecount As Long, tempstring2 As String, line As Long, index2 As Long
    tempstring$ = thestring$
    whereat& = InStr(tempstring$, Chr(13))
    If whereat& = 0& Then Exit Function
    For index& = 0& To getstringlinecount(thestring$) - 1
        linestring$ = getstringlinetext(tempstring$, index&, True)
        If lcasetext = True Then linestring$ = LCase(linestring$)
        If removespaces = True Then
            linestring$ = removechar(linestring$, " ")
            linestring$ = removechar(linestring$, " ")
            linestring$ = removechar(linestring$, " ")
        End If
        If linestring$ = theline$ Then
            getstringlineindex& = index& + 1
            Exit For
        End If
    Next index&
End Function
Public Function getlistcount(list As Long) As Long
    getlistcount& = SendMessage(list&, LB_GETCOUNT, 0&, 0&)
End Function
Public Function getcombocount(combo As Long) As Long
    getcombocount& = SendMessage(combo&, CB_GETCOUNT, 0&, 0&)
End Function
Public Function getimtext(imwin As Long) As String
    Dim imtext As Long
    imtext& = FindWindowEx(imwin&, 0&, "RICHCNTL", vbNullString)
    getimtext$ = gettext(imtext&)
End Function
Public Function getchattext() As String
    Dim chattext As Long
    chattext& = FindWindowEx(findroom(), 0&, "RICHCNTL", vbNullString)
    getchattext$ = gettext(chattext&)
End Function
Public Sub waitfortxttoload(list As Long)
    'thanks for the idea izekial
    Dim gettxt1 As String, gettxt2 As String, gettxt3 As String
    Do: DoEvents
        gettxt1$ = gettext(list&)
        pause 0.8
        gettxt2$ = gettext(list&)
        pause 0.8
        gettxt3$ = gettext(list&)
    Loop Until gettxt1$ = gettxt2$ And gettxt2$ = gettxt3$
End Sub
Public Sub winontop(win As Long, onoff As Boolean)
    If onoff = True Then
        Call SetWindowPos(win&, HWND_TOPMOST, 0&, 0&, 0&, 0&, Flags)
    Else
        Call SetWindowPos(win&, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, Flags)
    End If
End Sub
Public Sub winsetaolparent(win As Long)
    'thanks neo for the idea
    Dim aol As Long, mdi As Long
    aol& = FindWindow("AOL Frame25", vbNullString)
    mdi& = FindWindowEx(aol&, 0&, "MDIClient", vbNullString)
    Call SetParent(win&, mdi&)
End Sub
Public Function writetoini(section As String, key As String, keyvalue As String, directory As String)
    Call WritePrivateProfileString(section$, UCase$(key$), keyvalue$, directory$)
End Function

