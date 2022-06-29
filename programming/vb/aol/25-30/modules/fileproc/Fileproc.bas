Option Explicit

Global Const NUMBOXES = 5
Global Const SAVEFILE = 1, LOADFILE = 2
Global Const REPLACEFILE = 1, READFILE = 2, ADDTOFILE = 3
Global Const RANDOMFILE = 4, BINARYFILE = 5

' Define a data type to hold a record:
' Define global variables to hold the file number and record number
' of the current data file.
' Default file name to show in dialog boxes.
Global Const Err_DeviceUnavailable = 68
Global Const Err_DiskNotReady = 71, Err_FileAlreadyExists = 58
Global Const Err_TooManyFiles = 67, Err_RenameAcrossDisks = 74
Global Const Err_Path_FileAccessError = 75, Err_DeviceIO = 57
Global Const Err_DiskFull = 61, Err_BadFileName = 64
Global Const Err_BadFileNameOrNumber = 52, Err_FileNotFound = 53
Global Const Err_PathDoesNotExist = 76, Err_BadFileMode = 54
Global Const Err_FileAlreadyOpen = 55, Err_InputPastEndOfFile = 62
Global Const MB_EXCLAIM = 48, MB_STOP = 16

Function FileErrors (errVal As Integer) As Integer
' Return Value  Meaning     Return Value    Meaning
' 0             Resume      2               Unrecoverable error
' 1             Resume Next 3               Unrecognized error
Dim MsgType As Integer
Dim Response As Integer
Dim Action As Integer
Dim Msg As String
MsgType = MB_EXCLAIM
Select Case errVal
    Case Err_DeviceUnavailable  ' Error #68
	Msg = "That device appears to be unavailable."
	MsgType = MB_EXCLAIM + 5
    Case Err_DiskNotReady       ' Error #71
	Msg = "The disk is not ready."
    Case Err_DeviceIO
	Msg = "The disk is full."
    Case Err_BadFileName, Err_BadFileNameOrNumber   ' Errors #64 & 52
	Msg = "That file name is illegal."
    Case Err_PathDoesNotExist                        ' Error #76
	Msg = "That path doesn't exist."
    Case Err_BadFileMode                            ' Error #54
	Msg = "Can't open your file for that type of access."
    Case Err_FileAlreadyOpen                        ' Error #55
	Msg = "That file is already open."
    Case Err_InputPastEndOfFile                     ' Error #62
    Msg = "This file has a nonstandard end-of-file marker,"
    Msg = Msg + "or an attempt was made to read beyond "
    Msg = Msg + "the end-of-file marker."
    Case Else
	FileErrors = 3
	Exit Function
    End Select
    Response = MsgBox(Msg, MsgType, "File Error")
    Select Case Response
	Case 4          ' Retry button.
	    FileErrors = 0
	Case 5          ' Ignore button.
	    FileErrors = 1
	Case 1, 2, 3    ' Ok and Cancel buttons.
	    FileErrors = 2
	Case Else
	    FileErrors = 3
    End Select
End Function

Function FileOpener (NewFileName As String, Mode As Integer, RecordLen As Integer, Confirm As Integer) As Integer
     Dim NewFileNum As Integer
     Dim Action As Integer
     Dim FileExists As Integer
     Dim Msg As String
     On Error GoTo OpenerError
     If NewFileName Like "*[;-?[* ]*" Or NewFileName Like "*]*" Then Error Err_BadFileName
     If Confirm Then
	If Dir(NewFileName) = "" Then
	    FileExists = False
	Else
	    FileExists = True
	End If
	If Mode = REPLACEFILE And FileExists Then
	    Msg = "Replace contents of " + NewFileName + "?"
	    If MsgBox(Msg, 49, "Replace File?") = 2 Then
		FileOpener = 0
		Exit Function
	    End If
	End If
	If Not FileExists Then
	    Msg = "The file " + NewFileName + " does not exist. "
	    Msg = Msg + "Do you want to create it?"
	    If MsgBox(Msg, 1, "Create File?") = 2 Then
		FileOpener = 0
		Exit Function
	    End If
	End If
     End If
     NewFileNum = FreeFile
     Select Case Mode
	  Case REPLACEFILE
	    Open NewFileName For Output As NewFileNum
	  Case READFILE
	    Open NewFileName For Input As NewFileNum
	  Case ADDTOFILE
	    Open NewFileName For Append As NewFileNum
	  Case RANDOMFILE
	    Open NewFileName For Random As NewFileNum Len = RecordLen
	  Case BINARYFILE
	    Open NewFileName For Binary As NewFileNum
	  Case Else
	    Exit Function
     End Select
     FileOpener = NewFileNum
Exit Function
OpenerError:
     Action = FileErrors(Err)
     Select Case Action
	Case 0
	    Resume
	Case Else
	    FileOpener = 0
	    Exit Function
     End Select
End Function

Function GetFileName (Prompt As String) As String
    GetFileName = LTrim$(RTrim$(UCase$(InputBox$(Prompt, "Enter File Name"))))
End Function

