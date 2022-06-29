Attribute VB_Name = "KnK"
'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global inipath
Global UserAOL
Public fMainForm As Form1

Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
(ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)

Sub Main()
    frmSplash.show
    frmSplash.Refresh
    Set fMainForm = New Form1
    Load Form1
'frmSplash.Label1.Caption = "Loading: Frames"
Form1.Frame1.Top = 4800
Form1.Frame1.Left = 1920
Form1.Frame2.Top = 4800
Form1.Frame2.Left = 3480
Form1.Frame3.Top = 0
Form1.Frame3.Left = 5400
Form1.Frame4.Top = 2160
Form1.Frame4.Left = 5640
Form1.Frame5.Top = 2160
Form1.Frame5.Left = 9120
'frmSplash.Label1.Caption = "Loading: Methods #1"
ascii$ = GetFromINI("Number4", "ascii", App.Path + "\KnK4Life.ini")
If ascii$ = "" Then
R% = WritePrivateProfileString("Number4", "ascii", "", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number4", "sendlist", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number4", "sendstatus", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number4", "sendthanx", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number4", "find", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number4", "send", "Empty", App.Path + "\KnK4Life.ini")
End If
'<!-------------
ascii$ = GetFromINI("Number5", "ascii", App.Path + "\KnK4Life.ini")
If ascii$ = "" Then
R% = WritePrivateProfileString("Number5", "ascii", "", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number5", "sendlist", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number5", "sendstatus", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number5", "sendthanx", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number5", "find", "Empty", App.Path + "\KnK4Life.ini")
R% = WritePrivateProfileString("Number5", "send", "Empty", App.Path + "\KnK4Life.ini")
End If
Form1.Text8.Text = GetFromINI("Number4", "ascii", App.Path + "\KnK4Life.ini")
Form1.Text4.Text = GetFromINI("Number4", "sendlist", App.Path + "\KnK4Life.ini")
Form1.Text5.Text = GetFromINI("Number4", "sendstatus", App.Path + "\KnK4Life.ini")
Form1.Text6.Text = GetFromINI("Number4", "sendthanx", App.Path + "\KnK4Life.ini")
Form1.Text7.Text = GetFromINI("Number4", "find", App.Path + "\KnK4Life.ini")
Form1.Text9.Text = GetFromINI("Number4", "send", App.Path + "\KnK4Life.ini")
'<!-------------
Form1.Text11.Text = GetFromINI("Number5", "ascii", App.Path + "\KnK4Life.ini")
Form1.Text15.Text = GetFromINI("Number5", "sendlist", App.Path + "\KnK4Life.ini")
Form1.Text14.Text = GetFromINI("Number5", "sendstatus", App.Path + "\KnK4Life.ini")
Form1.Text13.Text = GetFromINI("Number5", "sendthanx", App.Path + "\KnK4Life.ini")
Form1.Text12.Text = GetFromINI("Number5", "find", App.Path + "\KnK4Life.ini")
Form1.Text10.Text = GetFromINI("Number5", "send", App.Path + "\KnK4Life.ini")



'frmSplash.Label1.Caption = "Loading: Options"
yesorno$ = GetSetting("KnK4Life", "KnK_Clearlist", "yesorno")
If yesorno$ = "yes" Then
    Form1.cl.Checked = True
End If
If yesorno$ = "no" Then
    Form1.dcl.Checked = True
End If

onoroff$ = GetSetting("KnK4Life", "KnK_IM", "onoroff")
If onoroff$ = "on" Then
    Form1.on2.Checked = True
End If
    If onoroff$ = "off" Then
Form1.off2.Checked = True
End If
''''''
onoroff1$ = GetSetting("KnK4Life", "KnK_show", "onoroff")
If onoroff1$ = "on" Then
    Form1.show1.Checked = True
End If

If onoroff1$ = "off" Then
    Form1.dont.Checked = True
End If

'frmSplash.Label1.Caption = "Loading: List"
For i = 0 To 10000
Form1.List1.AddItem i
Form1.List1.ListIndex = 0
Next i
'frmSplash.Label1.Caption = "Loading: Find Bin"
Dim a As Variant
Dim b As Variant
On Error GoTo kook
a = 1
Form1.List3.Clear
Open CStr(App.Path + "\find.lst") For Input As a
While (EOF(a) = False)
Line Input #a, b
Form1.List3.AddItem b
Wend
Close a
'End If
kook:
'frmSplash.Label1.Caption = "Done"
Form1.Width = 3480
Form1.Height = 2430

If Form1.Text8 = "" Then
End If
If Form1.Text8 <> "" Then
Form1.number4.Caption = Form1.Text8.Text + "SN" + " " + Form1.Text4.Text
End If
If Form1.Text11 = "" Then
End If
If Form1.Text11 <> "" Then
Form1.number5.Caption = Form1.Text11 + "SN" + " " + Form1.Text14
End If

Unload frmSplash
Form1.show
If onoroff1$ = "on" Then
frmStatus.show
End If
End Sub

Sub Setupform()
Form1.Width = 3480
Form1.Height = 2430
Form1.Timer1.Enabled = True
Form1.file.Enabled = False
Form1.tools.Enabled = False
Form1.option.Enabled = False
Form1.namess.Enabled = False
Form1.Command23.Visible = False
Form1.Command24.Visible = False
Form1.Command25.Visible = False
Form1.Command26.Visible = False
Form1.Combo1.Visible = False
Form1.Frame3.Top = 0
Form1.Frame3.Left = 0

End Sub
Sub Endsetups()
Form1.file.Enabled = True
Form1.tools.Enabled = True
Form1.option.Enabled = True
Form1.namess.Enabled = True
Form1.Command23.Visible = True
Form1.Command24.Visible = True
Form1.Command25.Visible = True
Form1.Command26.Visible = True
Form1.Combo1.Visible = True
Form1.Frame3.Top = 0
Form1.Frame3.Left = 5400
Form1.Timer1.Enabled = False
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





















'<!---------Made By KnK
'<!---------E-Mail me at Bill@knk.tierranet.com
'<!---------This was DL from http://knk.tierranet.com/knk4o

