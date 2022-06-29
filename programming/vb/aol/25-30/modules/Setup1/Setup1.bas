
Sub AddShareIfNeeded (SharePath$, ShareFile$)
    On Error GoTo ShareError

    fh% = FreeFile
    Open "C:\AUTOEXEC.BAT" For Input As fh%

    fFound% = 0
    While Not fFound% And Not EOF(fh%)
	Line Input #fh%, Temp1$
	If InStr(1, UCase$(Temp1$), "REM") = 0 And InStr(1, Temp1$, ";") = 0 And InStr(1, UCase$(Temp1$), "SHARE") > 0 Then
	   fFound% = True
	End If
    Wend

    Close #fh%

    If Not fFound% Then
	MsgBox "Please add <PATH>SHARE.EXE /L:500 to your AUTOEXEC.BAT"
    End If

    Exit Sub
ShareError:
    Close #fh%, #fh2%
    Exit Sub
End Sub

'-------------------------------------------------------
' Centers the passed form just above center on the screen
'-------------------------------------------------------
Sub CenterForm (x As Form)
  
    Screen.MousePointer = 11
    x.Top = (Screen.Height * .85) / 2 - x.Height / 2
    x.Left = Screen.Width / 2 - x.Width / 2
    Screen.MousePointer = 0

End Sub

Sub ConcatSplitFiles (firstfile$, cSplit%)
    Dim x%, fh1%, fh2%, outfile$, outfileLen&, CopyLeftOver&, CopyChunk#, filevar$
    Dim iFileMax%, iFile%, y%

    For x% = 2 To cSplit%
    
	fh1% = FreeFile
	Open Left$(firstfile$, Len(firstfile$) - 1) + Format$(1) For Binary As fh1%
		
	fh2% = FreeFile
	outfile$ = Left$(firstfile$, Len(firstfile$) - 1) + Format$(x%)
	Open outfile$ For Binary As fh2%
	    
	' Goto the end of file (plus one bytes) to start writing data
	Seek #fh1%, LOF(fh1%) + 1

	outfileLen& = LOF(fh2%)
	CopyLeftOver& = outfileLen& Mod 10
	CopyChunk# = (outfileLen& - CopyLeftOver&) / 10
	filevar$ = String$(CopyLeftOver&, 32)
	Get #fh2%, , filevar$
	Put #fh1%, , filevar$
	filevar$ = String$(CopyChunk#, 32)
	iFileMax% = 10
	For iFile% = 1 To iFileMax%
	    Get #fh2%, , filevar$
	    Put #fh1%, , filevar$
	Next iFile%

	Close fh1%, fh2%
	y% = SetTime(outfile$, firstfile$)
	Kill outfile$

    Next x%
    
    FileCopy Left$(firstfile$, Len(firstfile$) - 1) + Format$(1), firstfile$
    Kill Left$(firstfile$, Len(firstfile$) - 1) + Format$(1)
End Sub

'---------------------------------------------------------------
' Copies file SrcFilename from SourcePath to DestinationPath.
'
' Returns 0 if it could not find the file, or other runtime
' error occurs.  Otherwise, returns true.
'
' If the source file is older, the function returns success (-1)
' even though no file was copied, since no error occurred.
'---------------------------------------------------------------
Function CopyFile (ByVal SourcePath As String, ByVal DestinationPath As String, ByVal SrcFilename As String, ByVal DestFileName As String)
' ----- VerInstallFile() flags -----
    Const VIFF_FORCEINSTALL% = &H1, VIFF_DONTDELETEOLD% = &H2
    Const OF_DELETE% = &H200
    Const VIF_TEMPFILE& = &H1
    Const VIF_MISMATCH& = &H2
    Const VIF_SRCOLD& = &H4

    Const VIF_DIFFLANG& = &H8
    Const VIF_DIFFCODEPG& = &H10
    Const VIF_DIFFTYPE& = &H20
    Const VIF_WRITEPROT& = &H40
    Const VIF_FILEINUSE& = &H80
    Const VIF_OUTOFSPACE& = &H100
    Const VIF_ACCESSVIOLATION& = &H200
    Const VIF_SHARINGVIOLATION = &H400
    Const VIF_CANNOTCREATE = &H800
    Const VIF_CANNOTDELETE = &H1000
    Const VIF_CANNOTRENAME = &H2000
    Const VIF_CANNOTDELETECUR = &H4000
    Const VIF_OUTOFMEMORY = &H8000

    Const VIF_CANNOTREADSRC = &H10000
    Const VIF_CANNOTREADDST = &H20000

    Const VIF_BUFFTOOSMALL = &H40000
    Dim TmpOFStruct As OFStruct
    On Error GoTo ErrorCopy

    Screen.MousePointer = 11

    '--------------------------------------
    ' Add ending \ symbols to path variables
    '--------------------------------------
    If Right$(SourcePath$, 1) <> "\" Then
	SourcePath$ = SourcePath$ + "\"
    End If
    If Right$(DestinationPath$, 1) <> "\" Then
	DestinationPath$ = DestinationPath$ + "\"
    End If
    
    '----------------------------
    ' Update status dialog info
    '----------------------------
    Statusdlg.Label1.Caption = "Source file: " + Chr$(10) + Chr$(13) + UCase$(SourcePath$ + SrcFilename$)
    Statusdlg.Label1.Refresh
    Statusdlg.Label2.Caption = "Destination file: " + Chr$(10) + Chr$(13) + UCase$(DestinationPath$ + DestFileName$)
    Statusdlg.Label2.Refresh

    '-----------------------------------------
    ' Check the validity of the path and file
    '-----------------------------------------
CheckForExist:
    If Not FileExists(SourcePath$ + SrcFilename$) Then
	Screen.MousePointer = 0
	x% = MsgBox("Error occurred while attempting to copy file.  Could not locate file: """ + SourcePath$ + SrcFilename$ + """", 34, "SETUP")
	Screen.MousePointer = 11
	If x% = 3 Then
	    CopyFile = False
	ElseIf x% = 4 Then
	    GoTo CheckForExist
	ElseIf x% = 5 Then
	    GoTo SkipThisFile
	End If
    Else
	'-------------------------------------------------
	' VerInstallFile installs the file. We need to initialize
	' some arguments for the temp file that is created by the call
	'-------------------------------------------------
TryToCopyAgain:
	CurrDir$ = String$(255, 0)
	TmpFile$ = String$(255, 0)
	lpwTempFileLen% = 255
	InFileVer$ = GetFileVersion(SourcePath$ + SrcFilename$)
	OutFileVer$ = GetFileVersion(DestinationPath$ + DestFileName$)
	
	' Install if no version info is available
	If Len(InFileVer$) <> 0 And Len(OutFileVer$) <> 0 Then
	    ' Don't install older or same version of file
	    If InFileVer$ <= OutFileVer$ Then
		UpdateStatus GetFileSize(SourcePath$ + SrcFilename$)
		CopyFile = True
		Exit Function
	    End If
	End If

	Result& = VerInstallFile&(0, SrcFilename$, DestFileName$, SourcePath$, DestinationPath$, CurrDir$, TmpFile$, lpwTempFileLen%)

	'--------------------------------------------
	' After copying, update the installation meter
	'---------------------------------------------
	
	S$ = DestinationPath$
	If Right$(S$, 1) <> "\" Then S$ = S$ + "\"
	S$ = S$ + DestFileName$
	If Not TryAgain% Then UpdateStatus GetFileSize(S$)

	'--------------------------------
	' There are many return values that you can test for.
	' The constants are listed above.
	' The following lines of code return will set the Function to
	' True if the VerInstallFile call was successful.
	'
	' If the call was unsuccessful due to a different language on the
	' users machine, VerInstallFile is called again to force installation.
	' You can change this to not install if you choose.
	' Be careful about using FORCEINSTALL.  Other flags could be
	' set which indicate that this file should not be overridden.
	'
	' Under any other circumstance, the tempfile created by VerInstallFile
	' is removed using OpenFile and the CopyFile function returns false.
	'--------------------------------------------------------
	
	If Result& = 0 Or (Result& And VIF_SRCOLD&) = VIF_SRCOLD& Then
	    CopyFile = True
	ElseIf (Result& And VIF_DIFFLANG&) = VIF_DIFFLANG& Then
	    Result& = VerInstallFile&(VIFF_FORCEINSTALL%, SrcFilename$, DestFileName$, SourcePath$, DestinationPath$, CurrDir$, TmpFile$, lpwTempFileLen%)
	    CopyFile = True
	ElseIf (Result& And VIF_WRITEPROT&) = VIF_WRITEPROT& Then
	    Result& = VerInstallFile&(VIFF_FORCEINSTALL%, SrcFilename$, DestFileName$, SourcePath$, winSysDir$ + "\", CurrDir$, TmpFile$, lpwTempFileLen%)
	    CopyFile = True
	ElseIf (Result& And VIF_CANNOTREADSRC) = VIF_CANNOTREADSRC Then
	    ' VerInstallFile does will not handle compressed files that have been split.
	    ' Use VB's FileCopy stmt
	    FileCopy SourcePath$ + SrcFilename$, DestinationPath$ + DestFileName$
	    CopyFile = True
	Else
	    Screen.MousePointer = 0
	    If (Result& And VIF_FILEINUSE&) = VIF_FILEINUSE& Then
		x% = MsgBox(DestFileName$ & " is in use. Please close all applications and re-attempt Setup.", 34)
		If x% = 3 Then
		    CopyFile = False
		ElseIf x% = 4 Then
		    TryAgain% = True
		    GoTo TryToCopyAgain
		ElseIf x% = 5 Then
		    CopyFile = True
		    GoTo SkipThisFile
		End If
	    Else
		MsgBox DestFileName$ & " could not be installed."
		CopyFile = False
	    End If
	    Screen.MousePointer = 11
	End If

    If (Result& And VIF_TEMPFILE&) = VIF_TEMPFILE& Then copyresult% = OpenFile(TmpFile$, TmpOFStruct, OF_DELETE%)
       Screen.MousePointer = 0
       Exit Function
    End If

SkipThisFile:
       Exit Function
ErrorCopy:
    CopyFile = False
    Screen.MousePointer = 0
    Exit Function

End Function

'---------------------------------------------
' Create the path contained in DestPath$
' First char must be drive letter, followed by
' a ":\" followed by the path, if any.
'---------------------------------------------
Function CreatePath (ByVal DestPath$) As Integer
    Screen.MousePointer = 11

    '---------------------------------------------
    ' Add slash to end of path if not there already
    '---------------------------------------------
    If Right$(DestPath$, 1) <> "\" Then
	DestPath$ = DestPath$ + "\"
    End If
	  

    '-----------------------------------
    ' Change to the root dir of the drive
    '-----------------------------------
    On Error Resume Next
    ChDrive DestPath$
    If Err <> 0 Then GoTo errorOut
    ChDir "\"

    '-------------------------------------------------
    ' Attempt to make each directory, then change to it
    '-------------------------------------------------
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do While forePos <> 0
	temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)

	Err = 0
	MkDir temp$
	If Err <> 0 And Err <> 75 Then GoTo errorOut

	Err = 0
	ChDir temp$
	If Err <> 0 Then GoTo errorOut

	BackPos = forePos
	forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop
		 
    CreatePath = True
    Screen.MousePointer = 0
    Exit Function
		 
errorOut:
    MsgBox "Error While Attempting to Create Directories on Destination Drive.", 48, "SETUP"
    CreatePath = False
    Screen.MousePointer = 0

End Function

'-------------------------------------------------------------
' Procedure: CreateProgManGroup
' Arguments: X           The Form where a Label1 exist
'            GroupName$  A string that contains the group name
'            GroupPath$  A string that contains the group file
'                        name  ie 'myapp.grp'
'-------------------------------------------------------------
Sub CreateProgManGroup (x As Form, GroupName$, GroupPath$)
    
    Screen.MousePointer = 11
    
    '----------------------------------------------------------------------
    ' Windows requires DDE in order to create a program group and item.
    ' Here, a Visual Basic label control is used to generate the DDE messages
    '----------------------------------------------------------------------
    On Error Resume Next

    
    '--------------------------------
    ' Set LinkTopic to PROGRAM MANAGER
    '--------------------------------
    x.Label1.LinkTopic = "ProgMan|Progman"
    x.Label1.LinkMode = 2
    For i% = 1 To 10                                         ' Loop to ensure that there is enough time to
      z% = DoEvents()                                        ' process DDE Execute.  This is redundant but needed
    Next                                                     ' for debug windows.
    x.Label1.LinkTimeout = 100


    '---------------------
    ' Create program group
    '---------------------
    x.Label1.LinkExecute "[CreateGroup(" + GroupName$ + Chr$(44) + GroupPath$ + ")]"


    '-----------------
    ' Reset properties
    '-----------------
    x.Label1.LinkTimeout = 50
    x.Label1.LinkMode = 0
    
    Screen.MousePointer = 0
End Sub

'----------------------------------------------------------
' Procedure: CreateProgManItem
'
' Arguments: X           The form where Label1 exists
'
'            CmdLine$    A string that contains the command
'                        line for the item/icon.
'                        ie 'c:\myapp\setup.exe'
'
'            IconTitle$  A string that contains the item's
'                        caption
'----------------------------------------------------------
Sub CreateProgManItem (x As Form, CmdLine$, IconTitle$)
    
    Screen.MousePointer = 11
    
    '----------------------------------------------------------------------
    ' Windows requires DDE in order to create a program group and item.
    ' Here, a Visual Basic label control is used to generate the DDE messages
    '----------------------------------------------------------------------
    On Error Resume Next


    '---------------------------------
    ' Set LinkTopic to PROGRAM MANAGER
    '---------------------------------
    x.Label1.LinkTopic = "ProgMan|Progman"
    x.Label1.LinkMode = 2
    For i% = 1 To 10                                         ' Loop to ensure that there is enough time to
      z% = DoEvents()                                        ' process DDE Execute.  This is redundant but needed
    Next                                                     ' for debug windows.
    x.Label1.LinkTimeout = 100

    
    '------------------------------------------------
    ' Create Program Item, one of the icons to launch
    ' an application from Program Manager
    '------------------------------------------------
    If gfWin31% Then
	' Win 3.1 has a ReplaceItem, which will allow us to replace existing icons
	x.Label1.LinkExecute "[ReplaceItem(" + IconTitle$ + ")]"
    End If
    x.Label1.LinkExecute "[AddItem(" + CmdLine$ + Chr$(44) + IconTitle$ + Chr$(44) + ",,)]"
    x.Label1.LinkExecute "[ShowGroup(groupname, 1)]"         ' This will ensure that Program Manager does not
							     ' have a Maximized group, which causes problem in RestoreProgMan

    '-----------------
    ' Reset properties
    '-----------------
    x.Label1.LinkTimeout = 50
    x.Label1.LinkMode = 0
    
    Screen.MousePointer = 0
End Sub

'----------------------------------------------------------
' Check for the existence of a file by attempting an OPEN.
'----------------------------------------------------------
Function FileExists (path$) As Integer

    x = FreeFile

    On Error Resume Next
    Open path$ For Input As x
    If Err = 0 Then
	FileExists = True
    Else
	FileExists = False
    End If
    Close x

End Function

'------------------------------------------------
' Get the disk space free for the current drive
'------------------------------------------------
Function GetDiskSpaceFree (drive As String) As Long
    ChDrive drive
    GetDiskSpaceFree = DiskSpaceFree()
End Function

'----------------------------------------------------
' Get the disk Allocation unit for the current drive
'----------------------------------------------------
Function GetDrivesAllocUnit (drive As String) As Long
    ChDrive drive
    GetDrivesAllocUnit = AllocUnit()
End Function

'------------------------
' Get the size of the file
'------------------------
Function GetFileSize (source$) As Long
    x = FreeFile
    Open source$ For Binary Access Read As x
    GetFileSize = LOF(x)
    Close x
End Function

Function GetFileVersion (FileToCheck As String) As String
    On Error Resume Next
    VersionInfoSize& = GetFileVersionInfoSize(FileToCheck, lpdwHandle&)
    If VersionInfoSize& = 0 Then
	GetFileVersion = ""
	Exit Function
    End If
    lpvdata$ = String(VersionInfoSize&, Chr$(0))
    VersionInfo% = GetFileVersionInfo(FileToCheck, lpdwHandle&, VersionInfoSize&, lpvdata$)
    ptrFixed% = VerQueryValue(lpvdata$, "\FILEVERSION", lplpBuffer&, lpcb%)
    If ptrFixed% = 0 Then
	' Take a shot with the hardcoded TransString
	TransString$ = "040904E4"
	ptrString% = VerQueryValue(lpvdata$, "\StringFileInfo\" & TransString$ & "\CompanyName", lplpBuffer&, lpcb%)
	If ptrString% <> 0 Then GoTo GetValues
	ptrFixed% = VerQueryValue(lpvdata$, "\", lplpBuffer&, lpcb%)
	If ptrFixed% = 0 Then
	    GetFileVersion = ""
	    Exit Function
	Else
	    TransString$ = ""
	    fixedstr$ = String(lpcb% + 1, Chr(0))
	    stringcopy& = lstrcpyn(fixedstr$, lplpBuffer&, lpcb% + 1)
	    For i = lpcb% To 1 Step -1
		char$ = Hex(Asc(Mid(fixedstr$, i, 1)))
		If Len(char$) = 1 Then
		    char$ = "0" + char$
		End If
		TransString$ = TransString$ + char$
		If Len(TransString$ & nextchar$) Mod 8 = 0 Then
		    TransString$ = "&H" & TransString$
		    TransValue& = Val(TransString$)
		    TransString$ = ""
		End If
	    Next i
	End If
    End If
    TransTable$ = String(lpcb% + 1, Chr(0))
    TransString$ = String(0, Chr(0))
    stringcopy& = lstrcpyn(TransTable$, lplpBuffer&, lpcb% + 1)
    For i = 1 To lpcb%
	char$ = Hex(Asc(Mid(TransTable$, i, 1)))
	If Len(char$) = 1 Then
	    char$ = "0" + char$
	End If
	If Len(TransString$ & nextchar$) Mod 4 = 0 Then
	    nextchar$ = char$
	Else
	    TransString$ = TransString$ + char$ + nextchar$
	    nextchar$ = ""
	    char$ = ""
	End If
    Next i
GetValues:
    ptrString% = VerQueryValue(lpvdata$, "\StringFileInfo\" & TransString$ & "\FileVersion", lplpBuffer&, lpcb%)
    If ptrString% = 1 Then
	TransTable$ = String(lpcb%, Chr(0))
	stringcopy& = lstrcpyn(TransTable$, lplpBuffer&, lpcb% + 1)
	GetFileVersion = TransTable$
    Else
	GetFileVersion = ""
    End If
End Function

'--------------------------------------------------
' Calls the windows API to get the windows directory
'--------------------------------------------------
Function GetWindowsDir () As String
    temp$ = String$(145, 0)              ' Size Buffer
    x = GetWindowsDirectory(temp$, 145)  ' Make API Call
    temp$ = Left$(temp$, x)              ' Trim Buffer

    If Right$(temp$, 1) <> "\" Then      ' Add \ if necessary
	GetWindowsDir$ = temp$ + "\"
    Else
	GetWindowsDir$ = temp$
    End If
End Function

'---------------------------------------------------------
' Calls the windows API to get the windows\SYSTEM directory
'---------------------------------------------------------
Function GetWindowsSysDir () As String
    temp$ = String$(145, 0)                 ' Size Buffer
    x = GetSystemDirectory(temp$, 145)      ' Make API Call
    temp$ = Left$(temp$, x)                 ' Trim Buffer

    If Right$(temp$, 1) <> "\" Then         ' Add \ if necessary
	GetWindowsSysDir$ = temp$ + "\"
    Else
	GetWindowsSysDir$ = temp$
    End If
End Function

'------------------------------------------------------
' Function:   IsValidPath as integer
' arguments:  DestPath$         a string that is a full path
'             DefaultDrive$     the default drive.  eg.  "C:"
'
'  If DestPath$ does not include a drive specification,
'  IsValidPath uses Default Drive
'
'  When IsValidPath is finished, DestPath$ is reformated
'  to the format "X:\dir\dir\dir\"
'
' Result:  True (-1) if path is valid.
'          False (0) if path is invalid
'-------------------------------------------------------
Function IsValidPath (DestPath$, ByVal DefaultDrive$) As Integer

    '----------------------------
    ' Remove left and right spaces
    '----------------------------
    DestPath$ = RTrim$(LTrim$(DestPath$))
    

    '-----------------------------
    ' Check Default Drive Parameter
    '-----------------------------
    If Right$(DefaultDrive$, 1) <> ":" Or Len(DefaultDrive$) <> 2 Then
	MsgBox "Bad default drive parameter specified in IsValidPath Function.  You passed,  """ + DefaultDrive$ + """.  Must be one drive letter and "":"".  For example, ""C:"", ""D:""...", 64, "Setup Kit Error"
	GoTo parseErr
    End If
    

    '-------------------------------------------------------
    ' Insert default drive if path begins with root backslash
    '-------------------------------------------------------
    If Left$(DestPath$, 1) = "\" Then
	DestPath$ = DefaultDrive + DestPath$
    End If
    
    '-----------------------------
    ' check for invalid characters
    '-----------------------------
    On Error Resume Next
    tmp$ = Dir$(DestPath$)
    If Err <> 0 Then
	GoTo parseErr
    End If
    

    '-----------------------------------------
    ' Check for wildcard characters and spaces
    '-----------------------------------------
    If (InStr(DestPath$, "*") <> 0) GoTo parseErr
    If (InStr(DestPath$, "?") <> 0) GoTo parseErr
    If (InStr(DestPath$, " ") <> 0) GoTo parseErr
	 
    
    '------------------------------------------
    ' Make Sure colon is in second char position
    '------------------------------------------
    If Mid$(DestPath$, 2, 1) <> Chr$(58) Then GoTo parseErr
    

    '-------------------------------
    ' Insert root backslash if needed
    '-------------------------------
    If Len(DestPath$) > 2 Then
      If Right$(Left$(DestPath$, 3), 1) <> "\" Then
	DestPath$ = Left$(DestPath$, 2) + "\" + Right$(DestPath$, Len(DestPath$) - 2)
      End If
    End If

    '-------------------------
    ' Check drive to install on
    '-------------------------
    drive$ = Left$(DestPath$, 1)
    ChDrive (drive$)                                                        ' Try to change to the dest drive
    If Err <> 0 Then GoTo parseErr
    
    '-----------
    ' Add final \
    '-----------
    If Right$(DestPath$, 1) <> "\" Then
	DestPath$ = DestPath$ + "\"
    End If
    

    '-------------------------------------
    ' Root dir is a valid dir
    '-------------------------------------
    If Len(DestPath$) = 3 Then
	If Right$(DestPath$, 2) = ":\" Then
	    GoTo ParseOK
	End If
    End If
    

    '------------------------
    ' Check for repeated Slash
    '------------------------
    If InStr(DestPath$, "\\") <> 0 Then GoTo parseErr
	
    '--------------------------------------
    ' Check for illegal directory names
    '--------------------------------------
    legalChar$ = "!#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~.üäöÄÖÜß"
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do
	temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)
	
	'----------------------------
	' Test for illegal characters
	'----------------------------
	For i = 1 To Len(temp$)
	    If InStr(legalChar$, UCase$(Mid$(temp$, i, 1))) = 0 Then GoTo parseErr
	Next i

	'-------------------------------------------
	' Check combinations of periods and lengths
	'-------------------------------------------
	periodPos = InStr(temp$, ".")
	length = Len(temp$)
	If periodPos = 0 Then
	    If length > 8 Then GoTo parseErr                         ' Base too long
	Else
	    If periodPos > 9 Then GoTo parseErr                      ' Base too long
	    If length > periodPos + 3 Then GoTo parseErr             ' Extension too long
	    If InStr(periodPos + 1, temp$, ".") <> 0 Then GoTo parseErr' Two periods not allowed
	End If

	BackPos = forePos
	forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop Until forePos = 0

ParseOK:
    IsValidPath = True
    Exit Function

parseErr:
    IsValidPath = False
End Function

'----------------------------------------------------
' Prompt for the next disk.  Use the FileToLookFor$
' argument to verify that the proper disk, disk number
' wDiskNum, was inserted.
'----------------------------------------------------
Function PromptForNextDisk (wDiskNum As Integer, FileToLookFor$) As Integer

    '-------------------------
    ' Test for file
    '-------------------------
    Ready = False
    On Error Resume Next
    temp$ = Dir$(FileToLookFor$)

    '------------------------
    ' If not found, start loop
    '------------------------
    If Err <> 0 Or Len(temp$) = 0 Then
	While Not Ready
	    Err = 0
	    '----------------------------
	    ' Put up msg box
	    '----------------------------
	    Beep
	    x = MsgBox("Please insert disk # " + Format$(wDiskNum%), 49, "SETUP")
	    If x = 2 Then
		'-------------------------------
		' Use hit cancel, abort the copy
		'-------------------------------
		PromptForNextDisk = False
		GoTo ExitProc
	    Else
		'----------------------------------------
		' User hits OK, try to find the file again
		'----------------------------------------
		temp$ = Dir$(FileToLookFor$)
		If Err = 0 And Len(temp$) <> 0 Then
		    PromptForNextDisk = True
		    Ready = True
		End If
	    End If
	Wend
    Else
	PromptForNextDisk = True
    End If

    

ExitProc:

End Function

Sub RestoreProgMan ()
    On Error GoTo RestoreProgManErr
    AppActivate "Program Manager"   ' Activate Program Manager.
    SendKeys "%{ }{Enter}", True      ' Send Restore keystrokes.
RestoreProgManErr:
    Exit Sub
End Sub

'-----------------------------------------------------------------------------
' Set the Destination File's date and time to the Source file's date and time
'-----------------------------------------------------------------------------
Function SetFileDateTime (SourceFile As String, DestinationFile As String) As Integer
    x = SetTime(SourceFile, DestinationFile)
    SetFileDateTime = -1
End Function

Sub UpdateStatus (FileBytes As Long)
'-----------------------------------------------------------------------------
' Update the status bar using form.control Statusdlg.Picture2
'-----------------------------------------------------------------------------
    Static position
    Dim estTotal As Long

    estTotal = Val(Statusdlg.total.Tag)
    If estTotal = False Then
	estTotal = 10000000
    End If

    position = position + CSng((FileBytes / estTotal) * 100)
    If position > 100 Then
	position = 100
    End If
    Statusdlg.Picture2.Cls
    Statusdlg.Picture2.Line (0, 0)-((position * (Statusdlg.Picture2.ScaleWidth / 100)), Statusdlg.Picture2.ScaleHeight), QBColor(4), BF

    Txt$ = Format$(CLng(position)) + "%"
    Statusdlg.Picture2.CurrentX = (Statusdlg.Picture2.ScaleWidth - Statusdlg.Picture2.TextWidth(Txt$)) \ 2
    Statusdlg.Picture2.CurrentY = (Statusdlg.Picture2.ScaleHeight - Statusdlg.Picture2.TextHeight(Txt$)) \ 2
    Statusdlg.Picture2.Print Txt$

    r = BitBlt(Statusdlg.Picture1.hDC, 0, 0, Statusdlg.Picture2.ScaleWidth, Statusdlg.Picture2.ScaleHeight, Statusdlg.Picture2.hDC, 0, 0, SRCCOPY)

End Sub

