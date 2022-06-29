Attribute VB_Name = "basAPI"
Option Explicit

' General API functions.

Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long


Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREPOSITION = &H200
Private Const SWP_NOSIZE = &H1

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Private Declare Function FindWindow Lib "user32" _
   Alias "FindWindowA" (ByVal lpClassName As String, ByVal _
   lpWindowName As String) As Long
    
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function GetParent Lib "user32" _
   (ByVal hwnd As Long) As Long
   
Private Declare Function GetWindowTextLength Lib "user32" _
   Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
   
Private Declare Function GetWindowText Lib "user32" Alias _
   "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, _
   ByVal cch As Long) As Long

Private Declare Function GetUserNameA Lib "advapi32.dll" _
   (ByVal lpBuffer As String, nSize As Long) As Long

Private TaskBarhWnd As Long


'Exit's windows with one of the following results.
'   dwReserved = 0
Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
   uFlags As Long, ByVal dwReserved As Long) As Long
   
Public Const EXIT_LOGOFF = 0
Public Const EXIT_SHUTDOWN = 1
Public Const EXIT_REBOOT = 2

Private Declare Function GetComputerNameA Lib "kernel32" _
   (ByVal lpBuffer As String, nSize As Long) As Long

' General API functions. (with no VBasic wrapper)

'Puts the app to sleep for the given number of milliseconds
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub ExitWindows(ByVal uFlags As Long)
   Call ExitWindowsEx(uFlags, 0)
End Sub


Public Function GetUserName() As String
   Dim UserName As String * 255

   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function
'
' Returns the computer's name
'
Public Function GetComputerName() As String
   Dim UserName As String * 255

   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

'
' Returns the title of the active window.
'    if GetParent = true then the parent window is
'                   returned.
'
Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
   Dim i As Long
   Dim j As Long
   
   i = GetForegroundWindow
   
   
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   
      i = j
   End If
   
   GetActiveWindowTitle = GetWindowTitle(i)
End Function

Public Sub HideTaskBar()
    TaskBarhWnd = FindWindow("Shell_traywnd", "")
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    End If
End Sub
Public Sub ShowTaskBar()
    If TaskBarhWnd <> 0 Then
       Call SetWindowPos(TaskBarhWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End If
End Sub
'
' Returns the handle of the active window.
'    if GetParent = true then the parent window is
'                   returned.
'
Public Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
   Dim i As Long
   Dim j As Long
   
   i = GetForegroundWindow
   
   
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   
      i = j
   End If
   
   GetActiveWindow = i
End Function


Public Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim s As String
   
   l = GetWindowTextLength(hwnd)
   s = Space(l + 1)
   
   GetWindowText hwnd, s, l + 1
   
   GetWindowTitle = Left$(s, l)
End Function

'
'  Makes a form the top window if top = True.  When top = False it removes
'  this property.
'
Public Sub TopMostForm(f As Form, Top As Boolean)
   If Top Then
      SetWindowPos f.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
   Else
      SetWindowPos f.hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
   End If
End Sub

'
'  Sleeps for a given number of seconds.
'
Public Sub Pause(ByVal seconds As Single)
   Call Sleep(Int(seconds * 1000#))
End Sub




