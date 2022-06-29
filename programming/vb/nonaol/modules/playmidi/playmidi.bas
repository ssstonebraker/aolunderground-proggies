Attribute VB_Name = "PlayMIDI"
'************************************************
'*                 Made By Luigi                *
'************************************************
'*                                              *
'*            Mail Thundy@hotmail.com           *
'*           for questions or comments          *
'*                                              *
'************************************************




Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Global fore%
Global username As String
Function PlayMIDI(DriveDirFile As String, Optional loopIT As Boolean)

Dim returnStr As String * 255
Dim Shortpath$, X&
Shortpath = Space(Len(DriveDirFile))
X = GetShortPathName(DriveDirFile, Shortpath, Len(Shortpath))
If X = 0 Then GoTo ErrorHandler
If X > Len(DriveDirFile) Then 'not a long filename
Shortpath = DriveDirFile
Else                          'it is a long filename
Shortpath = Left(Shortpath, X) 'x is the length of the return buffer
End If
X = mciSendString("close yada", returnStr, 255, 0) 'just in case
X = mciSendString("open " & Chr(34) & Shortpath & Chr(34) & " type sequencer alias yada", returnStr, 255, 0)
    If X <> 0 Then GoTo theEnd  'invalid filename or path
X = mciSendString("play yada", returnStr, 255, 0)
    If X <> 0 Then GoTo theEnd  'device busy or not ready
    If Not loopIT Then Exit Function
Do While DoEvents
    X = mciSendString("status yada mode", returnStr, 255, 0)
        If X <> 0 Then Exit Function 'StopMIDI() was pressed or error
    If Left(returnStr, 7) = "stopped" Then X = mciSendString("play yada from 1", returnStr, 255, 0)
Loop
Exit Function
theEnd:  'MIDI errorhandler
returnStr = Space(255)
X = mciGetErrorString(X, returnStr, 255)
MsgBox Trim(returnStr), vbExclamation 'error message
X = mciSendString("close yada", returnStr, 255, 0)
Exit Function

ErrorHandler:
MsgBox "Invalid Filename or Error.", vbInformation
End Function
Function StopMIDI()

Dim X&
Dim returnStr As String * 255
X = mciSendString("status yada mode", returnStr, 255, 0)
    If Left(returnStr, 7) = "playing" Then X = mciSendString("stop yada", returnStr, 255, 0)
returnStr = Space(255)
X = mciSendString("status yada mode", returnStr, 255, 0)
    If Left(returnStr, 7) = "stopped" Then X = mciSendString("close yada", returnStr, 255, 0)
End Function
