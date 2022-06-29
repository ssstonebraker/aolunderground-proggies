Attribute VB_Name = "KnK"
'  ___________________________________________
'/                                                                                      \
'\___________________________________________/
'  Y                            KnK.bas #1                                   Y
'  |              These codes were written by: PooK             |
' /  Visit his site at: http://knk.tierranet.com/PooK        /
'|               This .bas was compiled by: KnK                  |
'|   To see if there are any new helpfull .bas's goto     |
' \               http://knk.tierranet.com/knk4o                 \
'  |         If you would like to submit somethin to this     |
' /            E-mail me at Bill@knk.tierranet.com              /
'|                                   KnK '98                                 |
' \________________________________________\
'/                                                                                    \
'\__________________________________________/
'This .bas works with others.  This was tested with Jolt32.bas
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global inipath


Public Sub MoveForm(frm As Form)
ReleaseCapture
X = SendMessage(frm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)

'To use this,  put the following code in the "Mousedown"  dec
'of a label or picture box *Replace frm with your formname.
'MoveForm(frm)

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

Public Function Random(Index As Integer)
Randomize
Result = Int((Index * Rnd) + 1)
Random = Result
'To usethis,  example
'Dim NumSel As Integer
'NumSel = Random(2)
'If NumSel = 1 Then

'The number in ( ) is the max num.
'With that example you will either get a 1 or 2
End Function

