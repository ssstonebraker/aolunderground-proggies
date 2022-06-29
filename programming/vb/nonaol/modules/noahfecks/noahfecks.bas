Attribute VB_Name = "noahfecks"
'This 3l33t .bas will hide your app from the Ctrl+Alt+Delete task list
'By noah fecks
'email my ass at noahfecks@hotmail.com
'latez!
Public Declare Function GetCurrentProcessId _
Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess _
Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess _
Lib "kernel32" (ByVal dwProcessID As Long, _
ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

Public Sub MakeMeService()
Dim pid As Long
Dim reserv As Long

pid = GetCurrentProcessId()
regserv = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub 'This fuction will hide your app from the Ctrl+Alt+Delete task list



Public Sub UnMakeMeService()
Dim pid As Long
Dim reserv As Long

pid = GetCurrentProcessId()
regserv = RegisterServiceProcess(pid, _
RSP_UNREGISTER_SERVICE)
End Sub
'This fuction will unhide your app from the Ctrl+Alt+Delete task list
'make sure you put this in the form so it doesnt slow down system resources
