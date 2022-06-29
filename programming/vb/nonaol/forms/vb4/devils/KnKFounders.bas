Attribute VB_Name = "KnK"
'  ___________________________________________
'/                                                                                      \
'\___________________________________________/
'  Y                        KnKFounders.bas #1                         Y
'  |              These codes were written by: PooK              |
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
'Added notes by Devil
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Global iniPath




Public Function GetFromINI(AppName$, KeyName$, FileName$) As String
Dim RetStr As String
RetStr = String(255, Chr(0))
GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
'To write to an ini type this
'R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\KnK.ini")

'Added Note: This is the added note i wuz talkin bout if you opend the write and read ini example
'Added Note: you DONT need to .ini ax:
'Added Note: R% = WritePrivateProfileString("ascii", "Color", "bbb", App.Path + "\KnK.EXM")
'Added Note: it will still read it too. thats what i did its awsome good for disguising things

'To read do this
'Color$ = GetFromINI("ascii", "Color", App.Path + "\KnK.ini")
'If Color$ = "bbb" Then
'also you can do text1.text=getfromini bla bla you get the idea
End Function



