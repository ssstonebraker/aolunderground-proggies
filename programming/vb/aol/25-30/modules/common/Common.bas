Attribute VB_Name = "basCommon"
Option Explicit
Option Compare Text

'
'Global Constants
'
Global Const gstrNULL$ = ""                             'Empty string
Global Const gstrSEP_DIR$ = "\"                         'Directory separator character
Global Const gstrSEP_DIRALT$ = "/"                      'Alternate directory separator character
Global Const gstrSEP_EXT$ = "."                         'Filename extension separator character
Global Const gstrCOLON$ = ":"
Global Const gstrSwitchPrefix1 = "-"
Global Const gstrSwitchPrefix2 = "/"
Global Const gstrCOMMA$ = ","
Global Const gstrDECIMAL$ = "."
Global Const gstrINI_PROTOCOL = "Protocol"

Global Const gintMAX_SIZE% = 255                        'Maximum buffer size
Global Const gintMIN_BUTTONWIDTH% = 1200
Global Const gsngBUTTON_BORDER! = 1.4

Global Const intDRIVE_REMOVABLE% = 2                    'Constants for GetDriveType
Global Const intDRIVE_FIXED% = 3
Global Const intDRIVE_REMOTE% = 4

Global Const gintNOVERINFO% = 32767                     'flag indicating no version info

'File names
Global Const gstrFILE_SETUP$ = "SETUP.LST"              'Name of setup information file

'Share type macros for files
Global Const mstrPRIVATEFILE = ""
Global Const mstrSHAREDFILE = "$(Shared)"

'INI File keys
#If Win16 Then
Global Const gstrINI_BTRIEVE$ = "Btrieve"
#End If
Global Const gstrINI_SETUP$ = "Setup"
Global Const gstrINI_APPNAME$ = "Title"
Global Const gstrINI_APPDIR$ = "DefaultDir"
Global Const gstrINI_APPEXE$ = "AppExe"
Global Const gstrINI_APPPATH$ = "AppPath"
Global Const gstrINI_FORCEUSEDEFDEST = "ForceUseDefDir"

'Setup information file macros
Global Const gstrAPPDEST$ = "$(AppPath)"
Global Const gstrWINDEST$ = "$(WinPath)"
Global Const gstrWINSYSDEST$ = "$(WinSysPath)"
Global Const gstrWINSYSDESTSYSFILE$ = "$(WinSysPathSysFile)"
Global Const gstrPROGRAMFILES$ = "$(ProgramFiles)"
Global Const gstrCOMMONFILES$ = "$(CommonFiles)"
Global Const gstrCOMMONFILESSYS$ = "$(CommonFilesSys)"
Global Const gstrDAODEST$ = "$(MSDAOPath)"

'Mouse Pointer Constants
Global Const gintMOUSE_DEFAULT% = 0
Global Const gintMOUSE_HOURGLASS% = 11

'MsgError() Constants
Global Const MSGERR_ERROR = 1
Global Const MSGERR_WARNING = 2

'MsgBox Constants
Global Const MB_OK = 0                                  'OK button only
Global Const MB_OKCANCEL = 1                            'OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2                    'Abort, Retry, Ignore buttons
Global Const MB_YESNO = 4                               'Yes and No buttons
Global Const MB_RETRYCANCEL = 5                         'Retry and Cancel buttons
Global Const MB_ICONSTOP = 16                           'Critical message
Global Const MB_ICONQUESTION = 32                       'Warning query
Global Const MB_ICONEXCLAMATION = 48                    'Warning message
Global Const MB_ICONINFORMATION = 64                    'Information message
Global Const MB_DEFBUTTON1 = 0                          'First button is default
Global Const MB_DEFBUTTON2 = 256                        'Second button is default
Global Const MB_DEFBUTTON3 = 512                        'Third button is default

'MsgBox return values
Global Const IDOK = 1                                   'OK button pressed
Global Const IDCANCEL = 2                               'Cancel button pressed
Global Const IDABORT = 3                                'Abort button pressed
Global Const IDRETRY = 4                                'Retry button pressed
Global Const IDIGNORE = 5                               'Ignore button pressed
Global Const IDYES = 6                                  'Yes button pressed
Global Const IDNO = 7                                   'No button pressed

'
'Type Definitions
'
Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    nReserved1 As Integer
    nReserved2 As Integer
    szPathName As String * 256
End Type

Type VERINFO                                            'Version FIXEDFILEINFO
    strPad1 As Long                                     'Pad out struct version
    strPad2 As Long                                     'Pad out struct signature
    nMSLo As Integer                                    'Low word of ver # MS DWord
    nMSHi As Integer                                    'High word of ver # MS DWord
    nLSLo As Integer                                    'Low word of ver # LS DWord
    nLSHi As Integer                                    'High word of ver # LS DWord
    strPad3(1 To 36) As Byte                            'Pad out rest of VERINFO struct (36 bytes)
End Type

Type PROTOCOL
    strName As String
    strFriendlyName As String
End Type

Global Const OF_EXIST& = &H4000&
Global Const OF_SEARCH& = &H400&
Global Const HFILE_ERROR% = -1

'
'Global Variables
'
Global LF$                                              'single line break
Global LS$                                              'double line break

'List of available protocols
Global gProtocol() As PROTOCOL
Global gcProtocols As Integer


#If Win16 Then
'
'API/DLL Declarations for 16 bit SetupToolkit
'
Declare Function DiskSpaceFree Lib "STKIT416.DLL" () As Long
Declare Function SetTime Lib "STKIT416.DLL" (ByVal strFileGetTime As String, ByVal strFileSetTime As String) As Integer
Declare Function AllocUnit Lib "STKIT416.DLL" () As Long
Declare Function GetWinPlatform Lib "STKIT416.DLL" () As Long
Declare Function DLLSelfRegister Lib "STKIT416.DLL" (ByVal lpDllName As String) As Integer
Declare Sub lmemcpy Lib "STKIT416.DLL" (strDest As Any, ByVal strSrc As Any, ByVal intBytes As Integer)

Declare Function OpenFile Lib "Kernel" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Integer) As Integer
Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
Declare Function WritePrivateProfileString Lib "Kernel" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Integer
Declare Function GetWindowsDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetSystemDirectory Lib "Kernel" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
Declare Function GetDriveType16 Lib "Kernel" Alias "GetDriveType" (ByVal intDriveNum As Integer) As Integer
Declare Function GetTempFileName16 Lib "Kernel" Alias "GetTempFileName" (ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Integer

Declare Function VerInstallFile Lib "VER.DLL" (ByVal Flags%, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Declare Function GetFileVersionInfoSize Lib "VER.DLL" (ByVal strFileName As String, lVerHandle As Long) As Long
Declare Function GetFileVersionInfo Lib "VER.DLL" (ByVal strFileName As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Integer
Declare Function VerQueryValue Lib "VER.DLL" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Integer

Declare Function GetModuleUsage Lib "Kernel" (ByVal hModule As Integer) As Integer

'-----------------------------------------------------------
' FUNCTION: FSyncShell
'
' Executes an external program and waits for it to complete

' Returns: True if the program was started OK, False otherwise
'-----------------------------------------------------------
'
Function FSyncShell(ByVal strExeName As String, intCmdShow As Integer) As Integer
    Const HINSTANCE_ERROR% = 32
    
    Dim hInstChild As Integer

    '
    'Shell program, if Shell worked, enter loop
    '
    hInstChild = Shell(strExeName, intCmdShow)
    If hInstChild >= HINSTANCE_ERROR Then
        While GetModuleUsage(hInstChild)
            DoEvents
        Wend
    End If

    FSyncShell = IIf(hInstChild < HINSTANCE_ERROR, False, True)
End Function

#Else

'
'API/DLL Declarations for 32 bit SetupToolkit
'
Declare Function DiskSpaceFree Lib "STKIT432.DLL" Alias "DISKSPACEFREE" () As Long
Declare Function SetTime Lib "STKIT432.DLL" (ByVal strFileGetTime As String, ByVal strFileSetTime As String) As Integer
Declare Function AllocUnit Lib "STKIT432.DLL" () As Long
Declare Function GetWinPlatform Lib "STKIT432.DLL" () As Long
Declare Function fNTWithShell Lib "STKIT432.DLL" () As Boolean
Declare Function FSyncShell Lib "STKIT432.DLL" Alias "SyncShell" (ByVal strCmdLine As String, ByVal intCmdShow As Long) As Long
Declare Function DLLSelfRegister Lib "STKIT432.DLL" (ByVal lpDllName As String) As Integer
Declare Sub lmemcpy Lib "STKIT432.DLL" (strDest As Any, ByVal strSrc As Any, ByVal lBytes As Long)
Declare Function OSfCreateShellGroup Lib "STKIT432.DLL" Alias "fCreateShellFolder" (ByVal lpstrDirName As String) As Long
Declare Function OSfCreateShellLink Lib "STKIT432.DLL" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String) As Long
Declare Function OSfRemoveShellLink Lib "STKIT432.DLL" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long
Private Declare Function OSGetLongPathName Lib "STKIT432.DLL" Alias "GetLongPathName" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Declare Function OpenFile Lib "Kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Declare Function GetPrivateProfileString Lib "Kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "Kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lplFileName As String) As Long
Declare Function GetWindowsDirectory Lib "Kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "Kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetDriveType32 Lib "Kernel32" Alias "GetDriveTypeA" (ByVal strWhichDrive As String) As Long
Declare Function GetTempFileName32 Lib "Kernel32" Alias "GetTempFileNameA" (ByVal strWhichDrive As String, ByVal lpPrefixString As String, ByVal wUnique As Integer, ByVal lpTempFileName As String) As Long

Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFileName As String, lVerHandle As Long) As Long
Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFileName As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
Private Declare Function OSGetShortPathName Lib "Kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
#End If

'-----------------------------------------------------------
' SUB: AddDirSep
' Add a trailing directory path separator (back slash) to the
' end of a pathname unless one already exists
'
' IN/OUT: [strPathName] - path to add separator to
'-----------------------------------------------------------
'
Sub AddDirSep(strPathName As String)
    If Right$(RTrim$(strPathName), Len(gstrSEP_DIR)) <> gstrSEP_DIR Then
        strPathName = RTrim$(strPathName) & gstrSEP_DIR
    End If
End Sub

'-----------------------------------------------------------
' FUNCTION: FileExists
' Determines whether the specified file exists
'
' IN: [strPathName] - file to check for
'
' Returns: True if file exists, False otherwise
'-----------------------------------------------------------
'
Function FileExists(ByVal strPathName As String) As Integer
    Dim intFileNum As Integer

    On Error Resume Next

    '
    'Remove any trailing directory separator character
    '
    If Right$(strPathName, 1) = gstrSEP_DIR Then
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

'-----------------------------------------------------------
' FUNCTION: GetDriveType
' Determine whether a disk is fixed, removable, etc. by
' calling Windows GetDriveType()
'-----------------------------------------------------------
'
Function GetDriveType(ByVal intDriveNum As Integer) As Integer
    '
    ' This function expects an integer drive number in Win16 or a string in Win32
    '
#If Win16 Then
    GetDriveType = GetDriveType16(intDriveNum)
#Else
    Dim strDriveName As String
    
    strDriveName = Chr$(Asc("A") + intDriveNum) & gstrCOLON & gstrSEP_DIR
    GetDriveType = CInt(GetDriveType32(strDriveName))
#End If
End Function

'-----------------------------------------------------------
' FUNCTION: ReadProtocols
' Reads the allowable protocols from the specified file.
'
' IN: [strInputFilename] - INI filename from which to read the protocols
'     [strINISection] - Name of the INI section
'-----------------------------------------------------------
Function ReadProtocols(ByVal strInputFilename As String, ByVal strINISection As String) As Boolean
    Dim intIdx As Integer
    Dim fOK As Boolean
    Dim strInfo As String
    Dim intOffset As Integer
    
    intIdx = 0
    fOK = True
    Erase gProtocol
    gcProtocols = 0
    
    Do
        strInfo = ReadIniFile(strInputFilename, strINISection, gstrINI_PROTOCOL & Format$(intIdx + 1))
        If strInfo <> gstrNULL Then
            intOffset = InStr(strInfo, gstrCOMMA)
            If intOffset > 0 Then
                'The "ugly" name will be first on the line
                ReDim Preserve gProtocol(intIdx + 1)
                gcProtocols = intIdx + 1
                gProtocol(intIdx + 1).strName = Left$(strInfo, intOffset - 1)
                
                '... followed by the friendly name
                gProtocol(intIdx + 1).strFriendlyName = Mid$(strInfo, intOffset + 1)
                If (gProtocol(intIdx + 1).strName = "") Or (gProtocol(intIdx + 1).strFriendlyName = "") Then
                    fOK = False
                End If
            Else
                fOK = False
            End If

            If Not fOK Then
                Exit Do
            Else
                intIdx = intIdx + 1
            End If
        End If
    Loop While strInfo <> gstrNULL
    
    ReadProtocols = fOK
End Function

'-----------------------------------------------------------
' FUNCTION: ResolveResString
' Reads resource and replaces given macros with given values
'
' Example, given a resource number 14:
'    "Could not read '|1' in drive |2"
'   The call
'     ResolveResString(14, "|1", "TXTFILE.TXT", "|2", "A:")
'   would return the string
'     "Could not read 'TXTFILE.TXT' in drive A:"
'
' IN: [resID] - resource identifier
'     [varReplacements] - pairs of macro/replacement value
'-----------------------------------------------------------
'
Function ResolveResString(ByVal resID As Integer, ParamArray varReplacements() As Variant) As String
    Dim intMacro As Integer
    Dim strResString As String
    
    strResString = LoadResString(resID)
    
    ' For each macro/value pair passed in...
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        Dim strMacro As String
        Dim strValue As String
        
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        
        ' Replace all occurrences of strMacro with strValue
        Dim intPos As Integer
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    
    ResolveResString = strResString
    
    Exit Function
    
MismatchedPairs:
    Resume Next
End Function

 '-----------------------------------------------------------
 ' FUNCTION GetLongPathName
 '
 ' Retrieve the long pathname version of a path possibly
 '   containing short subdirectory and/or file names
 '-----------------------------------------------------------
 '
 #If Win32 Then
 Function GetLongPathName(ByVal strShortPath As String) As String
     Const cchBuffer = 300
     Dim strLongPath As String * cchBuffer
     Dim lResult As Long

     On Error GoTo 0
     lResult = OSGetLongPathName(strShortPath, strLongPath, cchBuffer)
     If lResult = 0 Then
         Error 53 ' File not found
     Else
         GetLongPathName = StripTerminator(strLongPath)
     End If
 End Function
 #End If
 
 '-----------------------------------------------------------
 ' FUNCTION GetShortPathName
 '
 ' Retrieve the short pathname version of a path possibly
 '   containing long subdirectory and/or file names
 '-----------------------------------------------------------
 '
 #If Win32 Then
 Function GetShortPathName(ByVal strLongPath As String) As String
     Const cchBuffer = 300
     Dim strShortPath As String * cchBuffer
     Dim lResult As Long

     On Error GoTo 0
     lResult = OSGetShortPathName(strLongPath, strShortPath, cchBuffer)
     If lResult = 0 Then
         Error 53 ' File not found
     Else
         GetShortPathName = StripTerminator(strShortPath)
     End If
 End Function
 #End If
 
'-----------------------------------------------------------
' FUNCTION: GetTempFileName
' Get a temporary filename for a specified drive and
' filename prefix
'-----------------------------------------------------------
'
Function GetTempFileName(ByVal cDriveLetter As Integer, ByVal lpPrefixString As String, ByVal wUnique As Integer, lpTempFileName As String) As Integer
    '
    ' This function expects an integer drive number in Win16 or a string in Win32
    '
#If Win16 Then
    GetTempFileName = GetTempFileName16(cDriveLetter, lpPrefixString, wUnique, lpTempFileName)
#Else
    Dim strDriveName As String
    
    strDriveName = Chr$(Asc("A") + cDriveLetter) & gstrCOLON & gstrSEP_DIR
    GetTempFileName = CInt(GetTempFileName32(strDriveName, lpPrefixString, wUnique, lpTempFileName))
#End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetDiskSpaceFree
' Get the amount of free disk space for the specified drive
'
' IN: [strDrive] - drive to check space for
'
' Returns: Amount of free disk space, or -1 if an error occurs
'-----------------------------------------------------------
'
Function GetDiskSpaceFree(ByVal strDrive As String) As Long
    Dim strCurDrive As String
    Dim lDiskFree As Long

    On Error Resume Next

    '
    'Save the current drive
    '
    strCurDrive = Left$(CurDir$, 2)

    '
    'Fixup drive so it includes only a drive letter and a colon
    '
    If InStr(strDrive, gstrCOLON) = 0 Or Len(strDrive) > 2 Then
        strDrive = Left$(strDrive, 1) & gstrCOLON
    End If

    '
    'Change to the drive we want to check space for.  The DiskSpaceFree() API
    'works on the current drive only.
    '
    ChDrive strDrive

    '
    'If we couldn't change to the request drive, it's an error, otherwise return
    'the amount of disk space free
    '
    If Err <> 0 Or (strDrive <> Left$(CurDir$, 2)) Then
        lDiskFree = -1
    Else
        lDiskFree = DiskSpaceFree()
        If Err <> 0 Then    'If Setup Toolkit's DLL couldn't be found
            lDiskFree = -1
        End If
    End If

    If lDiskFree = -1 Then
        MsgError Error$ & LS$ & ResolveResString(resDISKSPCERR) & strDrive, MB_ICONEXCLAMATION, gstrTitle
    End If

    GetDiskSpaceFree = lDiskFree

    '
    'Cleanup by setting the current drive back to the original
    '
    ChDrive strCurDrive

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: GetUNCShareName
'
' Given a UNC names, returns the leftmost portion of the
' directory representing the machine name and share name.
' E.g., given "\\SCHWEIZ\PUBLIC\APPS\LISTING.TXT", returns
' the string "\\SCHWEIZ\PUBLIC"
'
' Returns a string representing the machine and share name
'   if the path is a valid pathname, else returns NULL
'-----------------------------------------------------------
'
Function GetUNCShareName(ByVal strFN As String) As Variant
    GetUNCShareName = Null
    If IsUNCName(strFN) Then
        Dim iFirstSeparator As Integer
        iFirstSeparator = InStr(3, strFN, gstrSEP_DIR)
        If iFirstSeparator > 0 Then
            Dim iSecondSeparator As Integer
            iSecondSeparator = InStr(iFirstSeparator + 1, strFN, gstrSEP_DIR)
            If iSecondSeparator > 0 Then
                GetUNCShareName = Left$(strFN, iSecondSeparator - 1)
            Else
                GetUNCShareName = strFN
            End If
        End If
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetWindowsSysDir
'
' Calls the windows API to get the windows\SYSTEM directory
' and ensures that a trailing dir separator is present
'
' Returns: The windows\SYSTEM directory
'-----------------------------------------------------------
'
Function GetWindowsSysDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the system directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetSystemDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator(strBuf)
        AddDirSep strBuf
        
        GetWindowsSysDir = UCase16(strBuf)
    Else
        GetWindowsSysDir = gstrNULL
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: IsWin32
'
' Returns true if this program is running under Win32 (i.e.
'   any 32-bit operating system)
'-----------------------------------------------------------
'
Function IsWin32() As Boolean
    IsWin32 = (IsWindows95() Or IsWindowsNT())
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindows95
'
' Returns true if this program is running under Windows 95
'   or successor
'-----------------------------------------------------------
'
Function IsWindows95() As Boolean
    Const dwMask95 = &H2&
    If GetWinPlatform() And dwMask95 Then
        IsWindows95 = True
    Else
        IsWindows95 = False
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: IsWindowsNT
'
' Returns true if this program is running under Windows NT
'-----------------------------------------------------------
'
Function IsWindowsNT() As Boolean
    Const dwMaskNT = &H1&
    If GetWinPlatform() And dwMaskNT Then
        IsWindowsNT = True
    Else
        IsWindowsNT = False
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: IsUNCName
'
' Determines whether the pathname specified is a UNC name.
' UNC (Universal Naming Convention) names are typically
' used to specify machine resources, such as remote network
' shares, named pipes, etc.  An example of a UNC name is
' "\\SERVER\SHARE\FILENAME.EXT".
'
' IN: [strPathName] - pathname to check
'
' Returns: True if pathname is a UNC name, False otherwise
'-----------------------------------------------------------
'
Function IsUNCName(ByVal strPathName As String) As Integer
    Const strUNCNAME$ = "\\//\"        'so can check for \\, //, \/, /\

    IsUNCName = IIf(InStr(strUNCNAME, Left$(strPathName, 2)) > 0, True, False)
End Function

'-----------------------------------------------------------
' FUNCTION: MakePathAux
'
' Creates the specified directory path.
'
' No user interaction occurs if an error is encountered.
' If user interaction is desired, use the related
'   MakePathAux() function.
'
' IN: [strDirName] - name of the dir path to make
'
' Returns: True if successful, False if error.
'-----------------------------------------------------------
'
Function MakePathAux(ByVal strDirName As String) As Boolean
    Dim strPath As String
    Dim intOffset As Integer
    Dim intAnchor As Integer
    Dim strOldPath As String

    On Error Resume Next

    '
    'Add trailing backslash
    '
    If Right$(strDirName, 1) <> gstrSEP_DIR Then
        strDirName = strDirName & gstrSEP_DIR
    End If

    strOldPath = CurDir$
    MakePathAux = False
    intAnchor = 0

    '
    'Loop and make each subdir of the path separately.
    '
    '
    intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
    intAnchor = intOffset 'Start with at least one backslash, i.e. "C:\FirstDir"
    Do
        intOffset = InStr(intAnchor + 1, strDirName, gstrSEP_DIR)
        intAnchor = intOffset

        If intAnchor > 0 Then
            strPath = Left$(strDirName, intOffset - 1)
            ' Determine if this directory already exists
            Err = 0
            ChDir strPath
            If Err Then
                ' We must create this directory
                Err = 0
                #If Win32 And LOGGING Then
                    NewAction gstrKEY_CREATEDIR, """" & strPath & """"
                #End If
                MkDir strPath
                #If Win32 And LOGGING Then
                    If Err Then
                        LogError ResolveResString(resMAKEDIR) & " " & strPath
                        AbortAction
                        GoTo Done
                    Else
                        CommitAction
                    End If
                #End If
            End If
        End If
    Loop Until intAnchor = 0

    MakePathAux = True
Done:
    ChDir strOldPath

    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: MsgError
'
' Forces mouse pointer to default, calls VB's MsgBox
' function, and logs this error and (32-bit only)
' writes the message and the user's response to the
' logfile (32-bit only)
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'     [intLogType] (optional) - The type of logfile entry to make.
'                   By default, creates an error entry.  Use
'                   the MsgWarning() function to create a warning.
'                   Valid types as MSGERR_ERROR and MSGERR_WARNING
'
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgError(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String, Optional ByVal intLogType As Variant) As Integer
    Dim iRet As Integer
    
    iRet = MsgFunc(strMsg, intFlags, strCaption)
    MsgError = iRet
    
    #If Win32 And LOGGING Then
        ' We need to log this error and decode the user's response.
        Dim strID As String
        Dim strLogMsg As String

        Select Case iRet
        Case IDOK
            strID = ResolveResString(resLOG_IDOK)
        Case IDCANCEL
            strID = ResolveResString(resLOG_IDCANCEL)
        Case IDABORT
            strID = ResolveResString(resLOG_IDABORT)
        Case IDRETRY
            strID = ResolveResString(resLOG_IDRETRY)
        Case IDIGNORE
            strID = ResolveResString(resLOG_IDIGNORE)
        Case IDYES
            strID = ResolveResString(resLOG_IDYES)
        Case IDNO
            strID = ResolveResString(resLOG_IDNO)
        Case Else
            strID = ResolveResString(resLOG_IDUNKNOWN)
        End Select

        strLogMsg = strMsg & LF$ & "(" & ResolveResString(resLOG_USERRESPONDEDWITH, "|1", strID) & ")"
        If IsMissing(intLogType) Then
            intLogType = MSGERR_ERROR
        End If
        Select Case intLogType
        Case MSGERR_WARNING
            LogWarning strLogMsg
        Case MSGERR_ERROR
            LogError strLogMsg
        Case Else
            LogError strLogMsg
        End Select
    #End If
End Function

'-----------------------------------------------------------
' FUNCTION: MsgFunc
'
' Forces mouse pointer to default and calls VB's MsgBox
' function
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'     [fLogAsError] - If present and True (MSGBOX_ERR), the 32-bit
'                       version logs this message and the user's
'                       response in the logfile as an error.
'                       Otherwise it is presented to the user
'                       only.  (It is easier to use the MsgError()
'                       function.)
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgFunc(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    Dim intOldPointer As Integer
  
    intOldPointer = Screen.MousePointer

    Screen.MousePointer = gintMOUSE_DEFAULT
    MsgFunc = MsgBox(strMsg, intFlags, strCaption)
    Screen.MousePointer = intOldPointer
End Function

'-----------------------------------------------------------
' FUNCTION: MsgWarning
'
' Forces mouse pointer to default, calls VB's MsgBox
' function, and logs this error and (32-bit only)
' writes the message and the user's response to the
' logfile (32-bit only)
'
' IN: [strMsg] - message to display
'     [intFlags] - MsgBox function type flags
'     [strCaption] - caption to use for message box
'
' Returns: Result of MsgBox function
'-----------------------------------------------------------
'
Function MsgWarning(ByVal strMsg As String, ByVal intFlags As Integer, ByVal strCaption As String) As Integer
    MsgWarning = MsgError(strMsg, intFlags, strCaption, MSGERR_WARNING)
End Function

'-----------------------------------------------------------
' SUB: SetMousePtr
'
' Provides a way to set the mouse pointer only when the
' pointer state changes.  For every HOURGLASS call, there
' should be a corresponding DEFAULT call.  Other types of
' mouse pointers are set explicitly.
'
' IN: [intMousePtr] - type of mouse pointer desired
'-----------------------------------------------------------
'
Sub SetMousePtr(intMousePtr As Integer)
    Static intPtrState As Integer

    Select Case intMousePtr
    Case gintMOUSE_HOURGLASS
        intPtrState = intPtrState + 1
    Case gintMOUSE_DEFAULT
        intPtrState = intPtrState - 1
        If intPtrState < 0 Then
            intPtrState = 0
        End If
    Case Else
        Screen.MousePointer = intMousePtr
        Exit Sub
    End Select

    Screen.MousePointer = IIf(intPtrState > 0, gintMOUSE_HOURGLASS, gintMOUSE_DEFAULT)
End Sub

'-----------------------------------------------------------
' FUNCTION: StripTerminator
'
' Returns a string without any zero terminator.  Typically,
' this was a string returned by a Windows API call.
'
' IN: [strString] - String to remove terminator from
'
' Returns: The value of the string passed in minus any
'          terminating zero.
'-----------------------------------------------------------
'
Function StripTerminator(ByVal strString As String) As String
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileVersion
'
' Returns the internal file version number for the specified
' file.  This can be different than the 'display' version
' number shown in the File Manager File Properties dialog.
' It is the same number as shown in the VB4 SetupWizard's
' File Details screen.  This is the number used by the
' Windows VerInstallFile API when comparing file versions.
'
' IN: [strFileName] - the file whose version # is desired
'     [fIsRemoteServerSupportFile] - whether or not this file is
'          a remote OLE automation server support file (.VBR)
'          (Enterprise edition only).  If missing, False is assumed.
'
' Returns: The Version number string if found, otherwise
'          gstrNULL
'-----------------------------------------------------------
'
Function GetFileVersion(ByVal strFileName As String, Optional ByVal fIsRemoteServerSupportFile) As String
    Dim sVerInfo As VERINFO
    Dim strVer As String

    On Error GoTo GFVError

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If
    
    '
    'Get the file version into a VERINFO struct, and then assemble a version string
    'from the appropriate elements.
    '
    If GetFileVerStruct(strFileName, sVerInfo, fIsRemoteServerSupportFile) = True Then
        strVer = Format$(sVerInfo.nMSHi) & gstrDECIMAL & Format$(sVerInfo.nMSLo) & gstrDECIMAL
        strVer = strVer & Format$(sVerInfo.nLSHi) & gstrDECIMAL & Format$(sVerInfo.nLSLo)
        GetFileVersion = strVer
    Else
        GetFileVersion = gstrNULL
    End If
    
    Exit Function
    
GFVError:
    GetFileVersion = gstrNULL
    Err = 0
End Function

'-----------------------------------------------------------
' FUNCTION: GetFileVerStruct
'
' Gets the file version information into a VERINFO TYPE
' variable
'
' IN: [strFileName] - name of file to get version info for
'     [fIsRemoteServerSupportFile] - whether or not this file is
'          a remote OLE automation server support file (.VBR)
'          (Enterprise edition only).  If missing, False is assumed.
' OUT: [sVerInfo] - VERINFO Type to fill with version info
'
' Returns: True if version info found, False otherwise
'-----------------------------------------------------------
'
Function GetFileVerStruct(ByVal strFileName As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
    Const strFIXEDFILEINFO$ = "\"

    Dim lVerSize As Long
    Dim lVerHandle As Long
    Dim lpBufPtr As Long
    Dim byteVerData() As Byte

    GetFileVerStruct = False

    If IsMissing(fIsRemoteServerSupportFile) Then
        fIsRemoteServerSupportFile = False
    End If

    If fIsRemoteServerSupportFile Then
        GetFileVerStruct = GetRemoteSupportFileVerStruct(strFileName, sVerInfo)
    Else
        '
        'Get the size of the file version info, allocate a buffer for it, and get the
        'version info.  Next, we query the Fixed file info portion, where the internal
        'file version used by the Windows VerInstallFile API is kept.  We then copy
        'the fixed file info into a VERINFO structure.
        '
        lVerSize = GetFileVersionInfoSize(strFileName, lVerHandle)
        If lVerSize > 0 Then
            ReDim byteVerData(lVerSize)
            If GetFileVersionInfo(strFileName, lVerHandle, lVerSize, byteVerData(0)) <> 0 Then ' (Pass byteVerData array via reference to first element)
                If VerQueryValue(byteVerData(0), strFIXEDFILEINFO & "", lpBufPtr, lVerSize) <> 0 Then
                    lmemcpy sVerInfo, lpBufPtr, lVerSize
                    GetFileVerStruct = True
                End If
            End If
        End If
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: GetRemoteSupportFileVerStruct
'
' Gets the file version information of a remote OLE server
' support file into a VERINFO TYPE variable (Enterprise
' Edition only).  Such files do not have a Windows version
' stamp, but they do have an internal version stamp that
' we can look for.
'
' IN: [strFileName] - name of file to get version info for
' OUT: [sVerInfo] - VERINFO Type to fill with version info
'
' Returns: True if version info found, False otherwise
'-----------------------------------------------------------
'
Function GetRemoteSupportFileVerStruct(ByVal strFileName As String, sVerInfo As VERINFO, Optional ByVal fIsRemoteServerSupportFile) As Boolean
    Const strVersionKey = "Version="
    Dim cchVersionKey As Integer
    Dim iFile As Integer

    cchVersionKey = Len(strVersionKey)
    sVerInfo.nMSHi = gintNOVERINFO
    
    On Error GoTo Failed
    
    iFile = FreeFile

    Open strFileName For Input Access Read Lock Read Write As #iFile
    
    ' Loop through each line, looking for the key
    While (Not EOF(iFile))
        Dim strLine As String

        Line Input #iFile, strLine
        If Left$(strLine, cchVersionKey) = strVersionKey Then
            ' We've found the version key.  Copy everything after the equals sign
            Dim strVersion As String
            
            strVersion = Mid$(strLine, cchVersionKey + 1)
            
            'Parse and store the version information
            PackVerInfo strVersion, sVerInfo

            'Convert the format 1.2.3 from the .VBR into
            '1.2.0.3, which is really want we want
            sVerInfo.nLSLo = sVerInfo.nLSHi
            sVerInfo.nLSHi = 0
            
            GetRemoteSupportFileVerStruct = True
            Close iFile
            Exit Function
        End If
    Wend
    
    Close iFile
    Exit Function

Failed:
    GetRemoteSupportFileVerStruct = False
End Function
'-----------------------------------------------------------
' FUNCTION: GetWindowsDir
'
' Calls the windows API to get the windows directory and
' ensures that a trailing dir separator is present
'
' Returns: The windows directory
'-----------------------------------------------------------
'
Function GetWindowsDir() As String
    Dim strBuf As String

    strBuf = Space$(gintMAX_SIZE)

    '
    'Get the windows directory and then trim the buffer to the exact length
    'returned and add a dir sep (backslash) if the API didn't return one
    '
    If GetWindowsDirectory(strBuf, gintMAX_SIZE) > 0 Then
        strBuf = StripTerminator$(strBuf)
        AddDirSep strBuf

        GetWindowsDir = UCase16(strBuf)
    Else
        GetWindowsDir = gstrNULL
    End If
End Function

'-----------------------------------------------------------
' FUNCTION: UCase16
'
' Returns the upper-case conversion of a string
'   under 16 bits, or else returns an unmodified
'   copy of the string under 32 bits.
'
' IN: [str] - String to copy/upper-case
'
'-----------------------------------------------------------
'
Function UCase16(ByVal str As String)
#If Win16 Then
    UCase16 = UCase$(str)
#Else
    UCase16 = str
#End If
End Function

'-----------------------------------------------------------
' FUNCTION: ExtractFilenameItem
'
' Extracts a quoted or unquoted filename from a string.
'
' IN: [str] - string to parse for a filename.
'     [intAnchor] - index in str at which the filename begins.
'             The filename continues to the end of the string
'             or up to the next comma in the string, or, if
'             the filename is enclosed in quotes, until the
'             next double quote.
' OUT: Returns the filename, without quotes.
'      [intAnchor] is set to the comma, or else one character
'             past the end of the string
'      [fErr] is set to True if a parsing error is discovered
'
'-----------------------------------------------------------
'
Function strExtractFilenameItem(ByVal str As String, intAnchor As Integer, fErr As Boolean) As String
    While Mid$(str, intAnchor, 1) = " "
        intAnchor = intAnchor + 1
    Wend
    
    Dim iEndFilenamePos As Integer
    Dim strFileName As String
    If Mid$(str, intAnchor, 1) = """" Then
        ' Filename is surrounded by quotes
        iEndFilenamePos = InStr(intAnchor + 1, str, """") ' Find matching quote
        If iEndFilenamePos > 0 Then
            strFileName = Mid$(str, intAnchor + 1, iEndFilenamePos - 1 - intAnchor)
            intAnchor = iEndFilenamePos + 1
            While Mid$(str, intAnchor, 1) = " "
                intAnchor = intAnchor + 1
            Wend
            If (Mid$(str, intAnchor, 1) <> gstrCOMMA) And (Mid$(str, intAnchor, 1) <> "") Then
                fErr = True
                Exit Function
            End If
        Else
            fErr = True
            Exit Function
        End If
    Else
        ' Filename continues until next comma or end of string
        Dim iCommaPos As Integer
        
        iCommaPos = InStr(intAnchor, str, gstrCOMMA)
        If iCommaPos = 0 Then
            iCommaPos = Len(str) + 1
        End If
        iEndFilenamePos = iCommaPos
        
        strFileName = Mid$(str, intAnchor, iEndFilenamePos - intAnchor)
        intAnchor = iCommaPos
    End If
    
    strFileName = Trim$(strFileName)
    If strFileName = "" Then
        fErr = True
        Exit Function
    End If
    
    fErr = False
    strExtractFilenameItem = strFileName
End Function

'-----------------------------------------------------------
' FUNCTION: Extension
'
' Extracts the extension portion of a file/path name
'
' IN: [strFileName] - file/path to get the extension of
'
' Returns: The extension if one exists, else gstrNULL
'-----------------------------------------------------------
'
Function Extension(ByVal strFileName As String) As String
    Dim intPos As Integer

    Extension = gstrNULL

    intPos = Len(strFileName)

    Do While intPos > 0
        Select Case Mid$(strFileName, intPos, 1)
        Case gstrSEP_EXT
            Extension = Mid$(strFileName, intPos + 1)
            Exit Do
        Case gstrSEP_DIR, gstrSEP_DIRALT
            Exit Do
        End Select

        intPos = intPos - 1
    Loop
End Function

'-----------------------------------------------------------
' SUB: PackVerInfo
'
' Parses a file version number string of the form
' x[.x[.x[.x]]] and assigns the extracted numbers to the
' appropriate elements of a VERINFO type variable.
' Examples of valid version strings are '3.11.0.102',
' '3.11', '3', etc.
'
' IN: [strVersion] - version number string
'
' OUT: [sVerInfo] - VERINFO type variable whose elements
'                   are assigned the appropriate numbers
'                   from the version number string
'-----------------------------------------------------------
'
Sub PackVerInfo(ByVal strVersion As String, sVerInfo As VERINFO)
    Dim intOffset As Integer
    Dim intAnchor As Integer

    On Error GoTo PVIError

    intOffset = InStr(strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSHi = Val(strVersion)
        GoTo PVIMSLo
    Else
        sVerInfo.nMSHi = Val(Left$(strVersion, intOffset - 1))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSHi
    Else
        sVerInfo.nMSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor))
        GoTo PVILSLo
    Else
        sVerInfo.nLSHi = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
        intAnchor = intOffset + 1
    End If

    intOffset = InStr(intAnchor, strVersion, gstrDECIMAL)
    If intOffset = 0 Then
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor))
    Else
        sVerInfo.nLSLo = Val(Mid$(strVersion, intAnchor, intOffset - intAnchor))
    End If

    Exit Sub

PVIError:
    sVerInfo.nMSHi = 0
PVIMSLo:
    sVerInfo.nMSLo = 0
PVILSHi:
    sVerInfo.nLSHi = 0
PVILSLo:
    sVerInfo.nLSLo = 0
End Sub


