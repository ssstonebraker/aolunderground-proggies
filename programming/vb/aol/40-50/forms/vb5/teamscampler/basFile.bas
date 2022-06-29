Attribute VB_Name = "basFile"
Option Explicit

Private Const MAX_FILENAME_LEN = 256

' File and Disk functions.
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_UNKNOWN = 0    'Unknown, or unable to be determined.

Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal nDrive As String) As Long

Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" _
   (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)

Private Declare Function GetWindowsDirectoryA Lib "kernel32" _
   (ByVal lpBuffer As String, ByVal nSize As Long) As Long
   
Private Declare Function GetTempPathA Lib "kernel32" _
   (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const UNIQUE_NAME = &H0

Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal _
   lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique _
   As Long, ByVal lpTempFileName As String) As Long
   
Private Declare Function GetSystemDirectoryA Lib "kernel32" _
   (ByVal lpBuffer As String, ByVal nSize As Long) As Long


Private Declare Function ShellExecute Lib _
   "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpdirectory As String, _
   ByVal nShowCmd As Long) As Long
   
Private Const SW_HIDE = 0             ' = vbHide
Private Const SW_SHOWNORMAL = 1       ' = vbNormal
Private Const SW_SHOWMINIMIZED = 2    ' = vbMinimizeFocus
Private Const SW_SHOWMAXIMIZED = 3    ' = vbMaximizedFocus
Private Const SW_SHOWNOACTIVATE = 4   ' = vbNormalNoFocus
Private Const SW_MINIMIZE = 6         ' = vbMinimizedNofocus

Private Declare Function GetShortPathNameA Lib "kernel32" _
   (ByVal lpszLongPath As String, ByVal lpszShortPath _
   As String, ByVal cchBuffer As Long) As Long
   
Private Type SHFILEOPSTRUCT
        hwnd As Long
        wFunc As Long
        pFrom As String
        pTo As String
        fFlags As Integer
        fAborted As Boolean
        hNameMaps As Long
        sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_SILENT = &H4
Private Const FOF_NOCONFIRMATION = &H10

Private Declare Function SHFileOperation Lib "shell32.dll" Alias _
   "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const SYNCHRONIZE = &H100000

Private Declare Function CloseHandle Lib "kernel32" (hObject As Long) As Boolean

Private Declare Function WaitForSingleObject Lib "kernel32" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CreateProcessA Lib "kernel32" _
    (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
    ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function FindExecutableA Lib "shell32.dll" _
   (ByVal lpFile As String, ByVal lpdirectory As _
   String, ByVal lpResult As String) As Long

Private Declare Function SetVolumeLabelA Lib "kernel32" _
   (ByVal lpRootPathName As String, _
   ByVal lpVolumeName As String) As Long

'
'  Finds the executable associated with a file
'
'  Returns "" if no file is found.
'
Public Function FindExecutable(s As String) As String
   Dim i As Integer
   Dim s2 As String
   
   s2 = String(MAX_FILENAME_LEN, 32) & Chr$(0)
   
   i = FindExecutableA(s & Chr$(0), vbNullString, s2)
   
   If i > 32 Then
      FindExecutable = Left$(s2, InStr(s2, Chr$(0)) - 1)
   Else
      FindExecutable = ""
   End If
   
End Function


'
'  Deletes a single file, or an array of files to the trashcan.
'
Public Function ShellDelete(ParamArray vntFileName() As Variant) As Boolean
   Dim i As Integer
   Dim sFileNames As String
   Dim SHFileOp As SHFILEOPSTRUCT

   For i = LBound(vntFileName) To UBound(vntFileName)
      sFileNames = sFileNames & vntFileName(i) & vbNullChar
   Next
        
   sFileNames = sFileNames & vbNullChar

   With SHFileOp
      .wFunc = FO_DELETE
      .pFrom = sFileNames
      .fFlags = FOF_ALLOWUNDO + FOF_SILENT + FOF_NOCONFIRMATION
   End With

   i = SHFileOperation(SHFileOp)
   
   If i = 0 Then
      ShellDelete = True
   Else
      ShellDelete = False
   End If
End Function
'
'  Runs a command as the Shell command does but waits for the command
'  to finish before returning.  Note: The full path and filename extention
'  is required.
'  You might want to use Environ$("COMSPEC") & " /c " & command
'  if you wish to run it under the command shell (and thus it)
'  will search the path etc...
'
'  returns false if the shell failed
'
Public Function ShellWait(cCommandLine As String) As Boolean
    Dim NameOfProc As PROCESS_INFORMATION
    Dim NameStart As STARTUPINFO
    Dim i As Long

    NameStart.cb = Len(NameStart)
    i = CreateProcessA(0&, cCommandLine, 0&, 0&, 1&, _
        NORMAL_PRIORITY_CLASS, 0&, 0&, NameStart, NameOfProc)
   
    If i <> 0 Then
       Call WaitForSingleObject(NameOfProc.hProcess, INFINITE)
       Call CloseHandle(NameOfProc.hProcess)
       ShellWait = True
    Else
       ShellWait = False
    End If
    
End Function

'
'  As the Execute function but waits for the process to finish before
'  returning
'
'  returns true on success.

Public Function ExecuteWait(s As String, Optional param As Variant) As Boolean
   Dim s2 As String
   
   s2 = FindExecutable(s)
   
   If s2 <> "" Then
      ExecuteWait = ShellWait(s2 & _
         IIf(IsMissing(param), " ", " " & CStr(param) & " ") & s)
   Else
      ExecuteWait = False
   End If
End Function
'
'  Adds a backslash if the string doesn't have one already.
'
Public Function AddBackslash(s As String) As String
   If Len(s) > 0 Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s + "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function

'
' Executes a file with it's associated program.
'    windowstyle uses the same constants as the Shell function:
'       vbHide   0
'       vbNormalFocus  1
'       vbMinimizedFocus  2
'       vbMaximizedFocus  3
'       vbNormalNoFocus   4
'       vbMinimizedNoFocus   6
'
'   returns true on success
Public Function Execute(ByVal hwnd As Integer, s As String, Optional param As Variant, Optional windowstyle As Variant) As Boolean
   Dim i As Long
   
   If IsMissing(windowstyle) Then
      windowstyle = vbNormalFocus
   End If
   
   i = ShellExecute(hwnd, vbNullString, s, IIf(IsMissing(param) Or (param = ""), vbNullString, CStr(param)), GetPath(s), CLng(windowstyle))
   If i > 32 Then
      Execute = True
   Else
      Execute = False
   End If
End Function

'
'  Returns the file portion of a file + pathname
'
Public Function GetFile(s As String) As String
   Dim i As Integer
   Dim j As Integer
   
   i = 0
   j = 0
   
   i = InStr(s, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, s, "\")
   Loop
   
   If j = 0 Then
      GetFile = ""
   Else
      GetFile = Right$(s, Len(s) - j)
   End If
End Function

'
'  Returns the path portion of a file + pathname
'
Public Function GetPath(s As String) As String
   Dim i As Integer
   Dim j As Integer
   
   i = 0
   j = 0
   
   i = InStr(s, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, s, "\")
   Loop
   
   If j = 0 Then
      GetPath = ""
   Else
      GetPath = Left$(s, j)
   End If
End Function
'
'  Returns a volume's serial number
'
Public Function GetSerialNumber(sDrive As String) As Long
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim i As Long
   Dim j As Long
   
   Call GetVolumeInformation(sDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, i, j, s2, MAX_FILENAME_LEN)
   GetSerialNumber = ser
End Function


Public Function GetShortPathName(longpath As String) As String
   Dim s As String
   Dim i As Long
   
   i = Len(longpath) + 1
   s = String(i, 0)
   GetShortPathNameA longpath, s, i
   
   GetShortPathName = Left$(s, InStr(s, Chr$(0)) - 1)
End Function

Public Function GetVolumeName(sDrive As String) As String
   Dim ser As Long
   Dim s As String * MAX_FILENAME_LEN
   Dim s2 As String * MAX_FILENAME_LEN
   Dim i As Long
   Dim j As Long
   
   Call GetVolumeInformation(sDrive + ":\" & Chr$(0), s, MAX_FILENAME_LEN, ser, i, j, s2, MAX_FILENAME_LEN)
   GetVolumeName = Left$(s, InStr(s, Chr$(0)) - 1)
End Function
'
'  Sets the volume name.  Returns true on success, false on failure.
'
Public Function SetVolumeName(sDrive As String, n As String) As Boolean
   Dim i As Long
   
   i = SetVolumeLabelA(sDrive + ":\" & Chr$(0), n & Chr$(0))
   
   SetVolumeName = IIf(i = 0, False, True)
End Function
'
'  Returns the system directory.
'
Public Function GetSystemDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetSystemDirectoryA("", 0)
   s = Space(i)
   Call GetSystemDirectoryA(s, i)
   GetSystemDirectory = AddBackslash(Left$(s, i - 1))
End Function

'
'  Returns a unique tempfile name.
'
Public Function GetTempFileName() As String
   Dim s As String
   Dim s2 As String
   
   s2 = GetTempPath
   s = Space(Len(s2) + MAX_FILENAME_LEN)
   Call GetTempFileNameA(s2, App.EXEName, UNIQUE_NAME, s)
   GetTempFileName = Left$(s, InStr(s, Chr$(0)) - 1)
End Function

'
'  Returns the path to the temp directory.
'
Public Function GetTempPath() As String
   Dim s As String
   Dim i As Integer
   i = GetTempPathA(0, "")
   s = Space(i)
   Call GetTempPathA(i, s)
   GetTempPath = AddBackslash(Left$(s, i - 1))
End Function

'
'  Returns the windows directory.
'
Public Function GetWindowsDirectory() As String
   Dim s As String
   Dim i As Integer
   i = GetWindowsDirectoryA("", 0)
   s = Space(i)
   Call GetWindowsDirectoryA(s, i)
   GetWindowsDirectory = AddBackslash(Left$(s, i - 1))
End Function

'
'  Removes the backslash from the string if it has one.
'
Public Function RemoveBackslash(s As String) As String
   Dim i As Integer
   i = Len(s)
   If i <> 0 Then
      If Right$(s, 1) = "\" Then
         RemoveBackslash = Left$(s, i - 1)
      Else
         RemoveBackslash = s
      End If
   Else
      RemoveBackslash = ""
   End If
End Function

'
' Returns the drive type if possible.
'
Public Function sDriveType(sDrive As String) As String
Dim lRet As Long

    lRet = GetDriveTypeA(sDrive & ":\")
    Select Case lRet
        Case 0
            'sDriveType = "Cannot be determined!"
            sDriveType = "Unknown"
            
        Case 1
            'sDriveType = "The root directory does not exist!"
            sDriveType = "Unknown"
        Case DRIVE_CDROM:
            sDriveType = "CD-ROM Drive"
            
        Case DRIVE_REMOVABLE:
            sDriveType = "Removable Drive"
            
        Case DRIVE_FIXED:
            sDriveType = "Fixed Drive"
            
        Case DRIVE_REMOTE:
            sDriveType = "Remote Drive"
        End Select
End Function

Public Function GetDriveType(sDrive As String) As Long
  Dim lRet As Long
  lRet = GetDriveTypeA(sDrive & ":\")
  
  If lRet = 1 Then
     lRet = 0
  End If

  GetDriveType = lRet
End Function
'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Public Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = "\" Then
        strPathName = Left$(strPathName, Len(strPathName) - 1)
    End If

    '
    'Attempt to open the file, return value of this function is False
    'if an error occurs on open, True otherwise
    '
    intFileNum = FreeFile
    Open strPathName For Input As intFileNum

    FileExists = IIf(Err, False, True)

    Close intFileNum

    Err = 0
End Function

