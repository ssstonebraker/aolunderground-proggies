Attribute VB_Name = "acceptaim"
'Please visit my site at
'http://teamparadox.cjb.net
'If you wanna see a certain example please contact me
'teamparadox@hotmail.com
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

Public Sub AcceptAimIm()
Dim AOLIcon As Long, AOLChild As Long, MDIClient As Long
Dim AOLFrame As Long
AOLFrame& = FindWindow("AOL Frame25", vbNullString) 'Find the mail aol win
MDIClient& = FindWindowEx(AOLFrame&, 0&, "MDIClient", vbNullString) 'Finds the grey MDI area
AOLChild& = FindWindowEx(MDIClient&, 0&, "AOL Child", vbNullString) 'Finds the accual accept window
AOLIcon& = FindWindowEx(AOLChild&, 0&, "_AOL_Icon", vbNullString) 'Find the 1st button
Call SendMessage(AOLIcon&, WM_LBUTTONDOWN, 0&, 0&) 'Clicks the button
Call SendMessage(AOLIcon&, WM_LBUTTONUP, 0&, 0&)
'Its alot easier then people like to make it, read Dos's tutorial, its the best i have ever seen
End Sub
