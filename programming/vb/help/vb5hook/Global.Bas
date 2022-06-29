Attribute VB_Name = "basGlobal"
Option Explicit

'Declare a new instance of our taskbar class.
Public TaskBar As New clsTaskBar

'Function and constants for exiting Windows.  Under Windows NT,
'a program MUST have certain security priviledges.  This demo
'does NOT account for those priviledges; therefore, this demo MAY
'not work properly under Win NT.
Declare Function ExitWindowsEx Lib "user32" _
(ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1


Sub Main()

'This is our start-up procedure.  All we're doing here is determining
'if the demo is already running and loading our main form if it's not.
If App.PrevInstance Then End

Load frmMain

   
End Sub

Public Function FileExists(sFileName As String) As Boolean
    
'This function checks for the existance of any filename passed to it

On Error GoTo FileExistError

'Make sure that a filename was passed or the function
'will return TRUE since it will find any file that exists
'in the current directory.
If Len(sFileName) = 0 Then
    FileExists = False
    Exit Function
ElseIf Len(Dir$(sFileName)) Then
    FileExists = True
Else
    FileExists = False
End If
Exit Function

FileExistError:
'Trap possible errors caused by the Dir function and return False
Select Case Err
    Case 76 'Path not found i.e. CD-ROM disk not present
        FileExists = False
    Case 75 'Path/File access error
        FileExists = False
    Case 68 'Device unavailable
        FileExists = False
End Select

End Function



Public Function StandardShutdown() As Long

'Function to perform a standard shutdown of Windows.  Needs to be a
'global function since it is called from more than one procedure.
StandardShutdown = ExitWindowsEx(EWX_SHUTDOWN, 0&)

End Function


