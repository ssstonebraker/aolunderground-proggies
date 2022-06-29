Attribute VB_Name = "KnK"
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_RBUTTONUP = &H205
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global inipath
Global UserAOL
Public fMainForm As frmServerhelp
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC = (-4)

Private Declare Function GetTickCount Lib "User" () As Long
Private pBar As New CProgressBar
Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmServerhelp
    Load frmServerhelp

'    frmSplash.pBar.Min = 0
'    If frmSplash.pBar.Value <> 0 Then
'    frmSplash.pBar.Value = 0
'    End If
'    frmSplash.pBar.Max = 10
    
'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    frmServerhelp.framRunning.Top = 0
    frmServerhelp.framRunning.Left = 5000
    frmServerhelp.frmSetup.Top = 2040
    frmServerhelp.frmSetup.Left = 5000
    
'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    sendlist$ = GetFromINI("Number4", "sendlist", App.Path + "\KnK4Life.ini")
    If sendlist$ = "" Then
        R% = WritePrivateProfileString("Number4", "ascii", "", App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendlist", "Empty", App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendstatus", "Empty", App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "sendthanx", "Empty", App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "find", "Empty", App.Path + "\KnK4Life.ini")
        R% = WritePrivateProfileString("Number4", "send", "Empty", App.Path + "\KnK4Life.ini")
    End If

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    frmServerhelp.Text8.Text = GetFromINI("Number4", "ascii", App.Path + "\KnK4Life.ini")
    frmServerhelp.Text4.Text = GetFromINI("Number4", "sendlist", App.Path + "\KnK4Life.ini")
    frmServerhelp.Text5.Text = GetFromINI("Number4", "sendstatus", App.Path + "\KnK4Life.ini")
    frmServerhelp.Text6.Text = GetFromINI("Number4", "sendthanx", App.Path + "\KnK4Life.ini")
    frmServerhelp.Text7.Text = GetFromINI("Number4", "find", App.Path + "\KnK4Life.ini")
    frmServerhelp.Text9.Text = GetFromINI("Number4", "send", App.Path + "\KnK4Life.ini")

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    If sendlist$ <> "Empty" Then
        frmServerhelp.number4.Caption = frmServerhelp.Text8.Text + "SN" + " " + frmServerhelp.Text4.Text
    End If

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    If Dir(App.Path + "\find2.lst") = "" Then
        Call SaveComboBox(App.Path + "\find2.lst", frmServerhelp.fndBin)
    End If

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    For i = 0 To 10000
        frmServerhelp.lstNumbers.AddItem i
        frmServerhelp.lstNumbers.ListIndex = 0
    Next i

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    If Dir(App.Path + "\room.lst") = "" Then
        Call SaveComboBox(App.Path + "\room.lst", frmServerhelp.fndBin)
    End If

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    If Dir(App.Path + "\find.lst") = "" Then
        Call SaveComboBox(App.Path + "\find.lst", frmServerhelp.fndBin)
    End If
        Call LoadComboBox(App.Path + "\find.lst", frmServerhelp.fndBin)

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    frmServerhelp.Width = 3480
    frmServerhelp.Height = 2430

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    If frmServerhelp.Text8 = "" Then
    End If
    If frmServerhelp.Text8 <> "" Then
        frmServerhelp.number4.Caption = frmServerhelp.Text8.Text + "SN" + " " + frmServerhelp.Text4.Text
    End If

'frmSplash.pBar.Value = frmSplash.pBar.Value + 1
    Unload frmSplash
    frmServerhelp.Show

End Sub

Sub Setupform()
    frmServerhelp.frmSetup.Top = 2040
    frmServerhelp.frmSetup.Left = 5000
    frmServerhelp.Width = 3480
    frmServerhelp.Height = 2430
    frmServerhelp.file.Enabled = False
    frmServerhelp.tools.Enabled = False
    frmServerhelp.option.Enabled = False
    frmServerhelp.namess.Enabled = False
    frmServerhelp.btnList.Visible = False
    frmServerhelp.btnStatus.Visible = False
    frmServerhelp.btnFind.Visible = False
    frmServerhelp.btnThrough.Visible = Fasle
    frmServerhelp.btnNext.Visible = Fasle
    frmServerhelp.btnClearbin.Visible = False
    frmServerhelp.btnRequest.Visible = False
    frmServerhelp.svrName.Visible = False
    frmServerhelp.fndBin.Visible = False
    frmServerhelp.fndBin.Visible = False
    frmServerhelp.framRunning.Top = 0
    frmServerhelp.framRunning.Left = 0

End Sub
Sub Endsetups()
    frmServerhelp.file.Enabled = True
    frmServerhelp.tools.Enabled = True
    frmServerhelp.option.Enabled = True
    frmServerhelp.namess.Enabled = True
    frmServerhelp.btnList.Visible = True
    frmServerhelp.btnStatus.Visible = True
    frmServerhelp.btnFind.Visible = True
    frmServerhelp.btnThrough.Visible = True
    frmServerhelp.btnNext.Visible = True
    frmServerhelp.btnClearbin.Visible = True
    frmServerhelp.btnRequest.Visible = True
    frmServerhelp.svrName.Visible = True
    frmServerhelp.fndBin.Visible = True
    frmServerhelp.fndBin.Visible = True
    frmServerhelp.framRunning.Top = 0
    frmServerhelp.framRunning.Left = 5000

End Sub



Public Function GetFromINI(AppName$, KeyName$, FileName$) As String
    Dim RetStr As String
    RetStr = String(255, Chr(0))
    GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
        'To write to an ini type this
        'R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\KnK.ini")

        'To read do this
        'Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
        'If Color$ = "bbb" Then

        '*Note* an .ini must be in the the same foder as the prog with these examples
        'For more info read the ini_Help.txt that was included with this
End Function

Public Sub DoChatStuff(strSN As String, strSaid As String, blnRTF As Boolean)
    Dim lngSpot As Long
    If strSN$ <> "" And strSaid$ <> "" Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 8
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelUnderline = False
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelColor = vbBlue
        If frmChat.txtChat.Text = "" Then
            frmChat.txtChat.SelText = strSN$ & ":" & Chr(9)
        Else
            frmChat.txtChat.SelText = vbCrLf & strSN$ & ":" & Chr(9)
        End If
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 10
        frmChat.txtChat.SelBold = False
        frmChat.txtChat.SelColor = vbBlack
        lngSpot& = Len(frmChat.txtChat.Text)
        If blnRTF = True Then
            frmChat.txtChat.SelRTF = strSaid$
        Else
            frmChat.txtChat.SelText = strSaid$
        End If
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelStart = lngSpot&
        frmChat.txtChat.SelLength = Len(frmChat.txtChat.Text) - lngSpot&
        frmChat.txtChat.SelHangingIndent = 1400
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
    End If
End Sub

Public Sub DoChatStuff2(strSN As String, strSaid As String, blnRTF As Boolean)
    Dim lngSpot As Long
    If strSN$ <> "" And strSaid$ <> "" Then
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 8
        frmChat.txtChat.SelBold = True
        frmChat.txtChat.SelUnderline = False
        frmChat.txtChat.SelItalic = False
        frmChat.txtChat.SelColor = vbRed
        If frmChat.txtChat.Text = "" Then
            frmChat.txtChat.SelText = strSN$ & ":" & Chr(9)
        Else
            frmChat.txtChat.SelText = vbCrLf & strSN$ & ":" & Chr(9)
        End If
        frmChat.txtChat.SelFontName = "Arial"
        frmChat.txtChat.SelFontSize = 10
        frmChat.txtChat.SelBold = False
        frmChat.txtChat.SelColor = vbBlack
        lngSpot& = Len(frmChat.txtChat.Text)
        If blnRTF = True Then
            frmChat.txtChat.SelRTF = strSaid$
        Else
            frmChat.txtChat.SelText = strSaid$
        End If
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
        frmChat.txtChat.SelStart = lngSpot&
        frmChat.txtChat.SelLength = Len(frmChat.txtChat.Text) - lngSpot&
        frmChat.txtChat.SelHangingIndent = 1400
        frmChat.txtChat.SelStart = Len(frmChat.txtChat.Text)
    End If
End Sub

Public Sub RandomStuff()
    Dim intPhrase As Integer, intPerson As Integer
    Dim intReply As Integer
    Randomize
    intPhrase% = Int((6 * Rnd) + 1)
'intPerson% = Int(frmChat.lstNames.ListCount * Rnd)
    intReply% = Int((6 * Rnd) + 1)
    Select Case intPhrase
        Case 1
           ' Call DoChatStuff("PeaceX101", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\fswiss Arial;}{\f4\froman Times New Roman;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red0\green128\blue128;\red128\green128\blue0;}\deflang1033\pard\plain\f4\fs24\cf2\ul Welcome \plain\f4\fs24\cf3\ul " & frmChat.lstNames.List(intPerson%) & "\plain\f2\fs20\par }", True)
            Call Pause(0.4)
            Select Case intReply%
                Case 1
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "turn that damn bot off peace", False)
                Case 2
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "30 years old and still programming a welcome bot peace?", False)
                Case 3
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "stfu peace before i reach through your screen and...", False)
                Case 4
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "peace, you are so gay", False)
                Case 5
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "is peace a bot?", False)
                Case 6
             '       Call DoChatStuff(frmChat.lstNames.List(intPerson%), "thanks peace. i couldn't call this room lame if you weren't here", False)
            End Select
        Case 2
        '    Call DoChatStuff("Izekial83", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\froman Times New Roman;}{\f4\fswiss Arial;}{\f5\froman Arial;}{\f6\fswiss Tahoma;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red128\green128\blue0;\red0\green128\blue128;\red0\green0\blue128;}\deflang1033\pard\plain\f6\fs24\cf4\b ^\plain\f3\fs24\cf4\b  izekial's ga\plain\f3\fs24\cf1\b y prog\plain\f2\fs20\par }", True)
            Call Pause(0.2)
            Call DoChatStuff("Izekial83", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\froman Times New Roman;}{\f4\fswiss Arial;}{\f5\froman Arial;}{\f6\fswiss Tahoma;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red128\green128\blue0;\red0\green128\blue128;\red0\green0\blue128;}\deflang1033\pard\plain\f6\fs24\cf4\b ^\plain\f3\fs24\cf4\b  100% copied\plain\f3\fs24\cf1\b  code\plain\f2\fs20\par }", True)
            Call Pause(0.2)
            Call DoChatStuff("Izekial83", "{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\fswiss Arial;}{\f3\froman Times New Roman;}{\f4\fswiss Arial;}{\f5\froman Arial;}{\f6\fswiss Tahoma;}{\f7\fswiss Tahoma;}}{\colortbl\red0\green0\blue0;\red0\green0\blue255;\red128\green128\blue0;\red0\green128\blue128;\red0\green0\blue128;}\deflang1033\pard\plain\f7\fs24\cf4\b ^\plain\f3\fs24\cf4\b  coded by ize\plain\f3\fs24\cf1\b kial83\plain\f2\fs20\par }", True)
        Case 3
            Call DoChatStuff("MacroBoy", "PReSs 555 eF JeW WaNnA gOin a PHat kNEw gRwP", False)
            Call Pause(0.2)
            Call DoChatStuff("MacroBoy", "555", False)
        Case 4
            Call DoChatStuff("MaGuSHaVoK", "Wanna Make A Prog Wit Me??? Anybody??? Please???", False)
        Case 5
            Call DoChatStuff("It Be Mi", "this room blows", True)
            Call Pause(0.4)
          '  Call DoChatStuff(frmChat.lstNames.List(intPerson%), "right on mi", False)
        Case 6
            'do nothing
    End Select
End Sub


