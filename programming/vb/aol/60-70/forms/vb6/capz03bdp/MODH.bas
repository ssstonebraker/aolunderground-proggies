Attribute VB_Name = "modHideTask"
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Const RSP_SIMPLE_SERVICE = 1
Public Const RSP_UNREGISTER_SERVICE = 0

'To remove your program from the Ctrl+Alt+Delete list, call the MakeMeService procedure:
Public Sub RemoveProgramFromList()
    Dim lngProcessID As Long
    Dim lngReturn As Long
    
    lngProcessID = GetCurrentProcessId()
    lngReturn = RegisterServiceProcess(pid, RSP_SIMPLE_SERVICE)
End Sub

'To restore your application to the Ctrl+Alt+Delete list, call the UnMakeMeService procedure:
Public Sub AddProgramToList()
    Dim lngProcessID As Long
    Dim lngReturn As Long
    
    lngProcessID = GetCurrentProcessId()
    lngReturn = RegisterServiceProcess(pid, RSP_UNREGISTER_SERVICE)
End Sub
