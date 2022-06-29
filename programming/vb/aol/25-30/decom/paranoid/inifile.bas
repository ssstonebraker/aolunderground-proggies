'***************************************************************************
'** INIFILE.BAS ** Third Public Release
'*************************************************
'** VB Module for simplifying .INI file operations
'***************************************************************************
'Copyright (C)Karl E. Peterson, February 1994, CIS 72302,3707.
'Portions originally downloaded from CompuServe's MSBASIC forum as
'MINIFILE.BAS, author unknown.  Comments and questions welcome!
'***************************************************************************
'This module contains "wrappers" for just about anything you'd want to do
'with INI files.  The only prerequisite for using them is to register the
'particular INI path/filename and [Section] in advance of calling them.
'Register Private.Ini by calling PrivIniRegister, and Win.Ini by calling
'WinIniRegister.
'
'This provides *safe* assured access to both application (Private.Ini) and
'Windows (Win.Ini) initialization files, with no need to worry about proper
'declarations and calling conventions.  It also greatly simplifies the task
'of repeatedly reading or writing to an Ini file.
'
'You are free to use this module as you see fit.  If you like it, I'd really
'appreciate hearing that!  If you don't like it, or have problems with it,
'I'd like to know that too.
'***************************************************************************
'The SECOND RELEASE added a dozen new functions, and two old ones were renamed.
'Latest modifications, June 1994
'  WinGetSectionEntries() is now WinGetSectEntries()
'  PrivGetSectionEntries() is now PrivGetSectEntries()
'Two new functions retrieve an entire [Section], entries and values, into an
'array from either Win.Ini or Private.Ini.  These functions are:
'  WinGetSectEntriesEx()
'  PrivGetSectEntriesEx()
'The other four deal with problems associated with multiple "device=" lines
'in System.Ini.  Use these at your *own risk*!  Especially the ones that add
'or remove a device.  These functions are:
'  SysDevAdd()              Adds a "device=" line to System.Ini
'  SysDevRemove()           Removes a "device=" line from System.Ini
'  SysDevLoaded()           Checks for a specific "device=" line
'  SysDevGetList()          Retrieves array of all devices
'The last six deal with [Section]'s.
'  Win/PrivGetSections()    Retrieves list of all [Section]'s
'  Win/PrivGetSectionsEx()  Retrieves array of all [Section]'s
'  Win/PrivSectionExist()   Verifies existence of registered [Section]
'***************************************************************************
'The THIRD RELEASE fixes a problem with the SysDevLoaded and SysDevRemove
'functions.  Neither worked if comments were on the same line.  Also, a flag
'has been added so that paths can be ignored or enforced with the SysXXX
'functions.  All API calls have been Aliased, so that this module may more
'easily be incorporated into existing programs.  Four new routines have
'been added:
'  SysIniRegister()         Set nmSysPath flag
'  ExtractName$()           Returns filename from filespec
'  ExtractPath$()           Returns path from filespec
'  StripComment$()          Removes trailing comments/spaces
'***************************************************************************
'The FOURTH RELEASE finally added some example code that exercises the
'routines in INIFILE.BAS!  The enclosed project, INIEDIT, is provided AS-IS,
'with no warranties expressed or implied.  Use it at your own risk, preferably
'on a copy of "real" INI files so you're not timid about adding and deleting
'data.
'
'The only changes made in INIFILE.BAS were to expand the Max_SectionBuffer
'length, and to set an error trap in the XXXGetSectEntriesEx functions.
'There appears to be a problem if a very large section is read with these
'routines (such as the [fonts] section in Win.Ini if several hundred fonts
'are installed).  If anyone has ideas on improving this, *please* let me
'know!  Currently, the return data is truncated, but that's better than an
'untrapped error, right? <g>
'***************************************************************************

Option Explicit

'** Windows API calls
'(NOTE: Profile calls *altered* from those found in WIN30API.TXT!)
  Declare Function kpGetProfileInt Lib "Kernel" Alias "GetProfileInt" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
  Declare Function kpGetProfileString Lib "Kernel" Alias "GetProfileString" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
  Declare Function kpWriteProfileString Lib "Kernel" Alias "WriteProfileString" (ByVal lpAppName As Any, ByVal lpKeyName As Any, ByVal lpString As Any) As Integer
  Declare Function kpGetPrivateProfileInt Lib "Kernel" Alias "GetPrivateProfileInt" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
  Declare Function kpGetPrivateProfileString Lib "Kernel" Alias "GetPrivateProfileString" (ByVal lpAppName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
  Declare Function kpWritePrivateProfileString Lib "Kernel" Alias "WritePrivateProfileString" (ByVal lpAppName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Integer
  Declare Function kpSendMessage Lib "User" Alias "SendMessage" (ByVal hWnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
  Declare Function kpGetWindowsDirectory Lib "Kernel" Alias "GetWindowsDirectory" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

'** Module-level variables for [Section] and Ini file names
  Dim smSectionName As String   'Current section in private Ini file
  Dim smIniFileName As String   'Fully qualified path/name of current private Ini file
  Dim smWinSection As String    'Current section in Win.Ini
  Dim nmWinInit As Integer      'Flag to indicate that Win.Ini section is initialized
  Dim nmPrivInit As Integer     'Flag to indicate that Private.Ini is initialized
  Dim nmSysPath As Integer      'Flag to indicate whether paths should be used with DEVICE=

'** Constants used to size buffers
  Const Max_SectionBuffer = 8192
  Const Max_EntryBuffer = 255

'** Special values to alert other apps of Win.Ini changes
  Const HWND_BROADCAST = &HFFFF
  Const WM_WININICHANGE = &H1A

Function ExtractName$ (sSpecIn$, nBaseOnly%)
  
  Dim nCnt%, nDot%, sSpecOut$

  On Local Error Resume Next

  If InStr(sSpecIn, "\") Then
    For nCnt = Len(sSpecIn) To 1 Step -1
      If Mid$(sSpecIn, nCnt, 1) = "\" Then
        sSpecOut = Mid$(sSpecIn, nCnt + 1)
        Exit For
      End If
    Next nCnt
  
  ElseIf InStr(sSpecIn, ":") = 2 Then
    sSpecOut = Mid$(sSpecIn, 3)
    
  Else
    sSpecOut = sSpecIn
  End If
    
  If nBaseOnly Then
    nDot = InStr(sSpecOut, ".")
    If nDot Then
      sSpecOut = Left$(sSpecOut, nDot - 1)
    End If
  End If

  ExtractName$ = UCase$(sSpecOut)

End Function

Function ExtractPath$ (sSpecIn$)

  Dim nCnt%, sSpecOut$
  
  On Local Error Resume Next

  If InStr(sSpecIn, "\") Then
    For nCnt = Len(sSpecIn) To 1 Step -1
      If Mid$(sSpecIn, nCnt, 1) = "\" Then
        sSpecOut = Left$(sSpecIn, nCnt)
        Exit For
      End If
    Next nCnt
  
  ElseIf InStr(sSpecIn, ":") = 2 Then
    sSpecOut = CurDir$(sSpecIn)
    If Len(sSpecOut) = 0 Then sSpecOut = CurDir$

  Else
    sSpecOut = CurDir$
  End If
    
  If Right$(sSpecOut, 1) <> "\" Then
    sSpecOut = sSpecOut + "\"
  End If
  ExtractPath$ = UCase$(sSpecOut)

End Function

Sub Main ()
  'This subroutine is useful for simply testing the other routines in this
  'module.  Make this module the only one in a project, and set Sub Main as
  'the entry point.  Then enter the code you wish to test below.
End Sub

Sub PrivClearEntry (sEntryName As String)

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Sub
    End If

  'Sets a specific entry in Private.Ini to Nothing or Blank
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(smSectionName, sEntryName, "", smIniFileName)

End Sub

Sub PrivDeleteEntry (sEntryName As String)

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Sub
    End If

  'Deletes a specific entry in Private.Ini
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(smSectionName, sEntryName, 0&, smIniFileName)

End Sub

Sub PrivDeleteSection ()

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Sub
    End If

  'Deletes an *entire* [Section] and all its Entries in Private.Ini
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(smSectionName, 0&, 0&, smIniFileName)

  'Now Private.Ini needs to be reinitialized
    smSectionName = ""
    nmPrivInit = False

End Sub

Function PrivGetInt (sEntryName As String, nDefaultValue As Integer) As Integer

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Retrieves an Integer value from Private.Ini, range: 0-32767
    PrivGetInt = kpGetPrivateProfileInt(smSectionName, sEntryName, nDefaultValue, smIniFileName)

End Function

Function PrivGetSectEntries () As String

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Retrieves all Entries in a [Section] of Private.Ini
  'Entries nul terminated; last entry double-terminated
    Dim sTemp As String * Max_SectionBuffer
    Dim nRetVal As Integer
    nRetVal = kpGetPrivateProfileString(smSectionName, 0&, "", sTemp, Len(sTemp), smIniFileName)
    PrivGetSectEntries$ = Left$(sTemp, nRetVal + 1)

End Function

Function PrivGetSectEntriesEx (sTable() As String) As Integer

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Example of usage, note return is one higher than UBound
    'Dim i%, n%
    'Dim eTable() As String
    'PrivIniRegister "386Enh", "System.Ini"
    'n% = PrivGetSectionEntriesEx(eTable())
    'For i = 0 To n - 1
    '  Debug.Print eTable(0, i); "="; eTable(1, i)
    'Next i

  'Retrieves all Entries in a [Section] of Private.Ini
  'Entries nul terminated; last entry double-terminated
    Dim sBuff As String * Max_SectionBuffer
    Dim sTemp As String
    Dim nRetVal As Integer
    nRetVal = kpGetPrivateProfileString(smSectionName, 0&, "", sBuff, Len(sBuff), smIniFileName)
    sTemp = Left$(sBuff, nRetVal + 1)

  'Parse entries into first dimension of table
  'and retrieve values into second dimension
    Dim nEntries As Integer
    Dim nNull As Integer
    On Error Resume Next
    Do While Asc(sTemp)
  'Bail if buffer wasn't large enough!!!
      If Err Then Exit Do
      ReDim Preserve sTable(0 To 1, 0 To nEntries)
      nNull = InStr(sTemp, Chr$(0))
      sTable(0, nEntries) = Left$(sTemp, nNull - 1)
      sTable(1, nEntries) = PrivGetString(sTable(0, nEntries), "")
      sTemp = Mid$(sTemp, nNull + 1)
      nEntries = nEntries + 1
    Loop

  'Make function assignment
    PrivGetSectEntriesEx = nEntries

End Function

Function PrivGetSections$ ()

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Setup some variables
    Dim sRet As String
    Dim sBuff As String
    Dim hFile As Integer

  'Extract all [Section] lines
    hFile = FreeFile
    Open smIniFileName For Input As hFile
    Do While Not EOF(hFile)
      Line Input #hFile, sBuff
      sBuff = StripComment$(sBuff)
      If InStr(sBuff, "[") = 1 And InStr(sBuff, "]") = Len(sBuff) Then
        sRet = sRet + Mid$(sBuff, 2, Len(sBuff) - 2) + Chr$(0)
      End If
    Loop
    Close hFile

  'Assign return value
    If Len(sRet) Then
      PrivGetSections = sRet + Chr$(0)
    Else
      PrivGetSections = String$(2, 0)
    End If

End Function

Function PrivGetSectionsEx (sTable() As String) As Integer

  'Get "normal" list of all [Section]'s
    Dim sSect As String
    sSect = PrivGetSections$()
    If Len(sSect) = 0 Then
      PrivGetSectionsEx = 0
      Exit Function
    End If

  'Parse [Section]'s into table
    Dim nEntries As Integer
    Dim nNull As Integer
    Do While Asc(sSect)
      ReDim Preserve sTable(0 To nEntries)
      nNull = InStr(sSect, Chr$(0))
      sTable(nEntries) = Left$(sSect, nNull - 1)
      sSect = Mid$(sSect, nNull + 1)
      nEntries = nEntries + 1
    Loop

  'Make function assignment
    PrivGetSectionsEx = nEntries
  
End Function

Function PrivGetString (sEntryName As String, ByVal sDefaultValue As String) As String

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Retrieves Specific Entry from Private.Ini
    Dim sTemp As String * Max_EntryBuffer
    Dim nRetVal As Integer
    nRetVal = kpGetPrivateProfileString(smSectionName, sEntryName, sDefaultValue, sTemp, Len(sTemp), smIniFileName)
    If nRetVal Then
      PrivGetString = Left$(sTemp, nRetVal)
    End If

End Function

Function PrivGetTF (sEntryName As String, nDefaultValue As Integer)
  
  'Retrieves Specific Entry as either True/False from Private.Ini
  'local vars
    Dim sTF As String
    Dim sDefault As String

  'get string value from INI
    If nDefaultValue Then
      sDefault = "true"
    Else
      sDefault = "false"
    End If
    sTF = PrivGetString(sEntryName, sDefault)

  'interpret return string
    Select Case Trim$(UCase$(sTF))
      Case "YES", "Y", "TRUE", "T", "ON", "1", "-1"
        PrivGetTF = True
      Case "NO", "N", "FALSE", "F", "OFF", "0"
        PrivGetTF = False
      Case Else
        PrivGetTF = False
    End Select

End Function

Sub PrivIniFlushCache ()

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Sub
    End If

  'To improve performance, Windows keeps a cached version of the most-recently
  'accessed initialization file. If that filename is specified and the other
  'three parameters are NULL, Windows flushes the cache
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(0&, 0&, 0&, smIniFileName)

End Sub

Private Sub PrivIniNotReg ()
  
  'Warn *PROGRAMMER* that there's a logic error!
    MsgBox "[Section] and FileName Not Registered in Private.Ini!", 16, "IniFile Logic Error"

End Sub

Sub PrivIniRead (SectionName$, KeyName$, nDefault%, ByVal DefaultStr$, ReturnStr$, Numeric%, IniFileName$)

  'One-shot read from Private.Ini, more *work* than it's worth
    Dim nRetVal As Integer
    Dim RetStr As String * Max_EntryBuffer 'Create an empty string to be filled

    If Numeric% Then    'we are looking for integer input
      Numeric% = kpGetPrivateProfileInt(SectionName$, KeyName$, nDefault%, IniFileName$)
    Else
      nRetVal = kpGetPrivateProfileString(SectionName$, KeyName$, DefaultStr$, RetStr$, Len(RetStr$), IniFileName$)
      If nRetVal Then
        ReturnStr$ = Left$(RetStr$, nRetVal)
      End If
    End If

End Sub

Sub PrivIniRegister (sSectionName As String, sIniFileName As String)

  'Store module-level values for future reference
    smSectionName = Trim$(sSectionName)
    smIniFileName = Trim$(sIniFileName)
    nmPrivInit = True

End Sub

Sub PrivIniWrite (SectionName$, IniFileName$, EntryName$, ByVal NewVal$)
    
  'One-shot write to Private.Ini, more *work* than it's worth
    Dim nRetVal As Integer
    nRetVal = kpWritePrivateProfileString(SectionName$, EntryName$, NewVal$, IniFileName$)
    
End Sub

Function PrivPutInt (sEntryName As String, nValue As Integer) As Integer

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Write an integer to Private.Ini
    PrivPutInt = kpWritePrivateProfileString(smSectionName, sEntryName, Format$(nValue), smIniFileName)

End Function

Function PrivPutString (sEntryName As String, ByVal sValue As String) As Integer

  'Bail if not initialized
    If Not nmPrivInit Then
      PrivIniNotReg
      Exit Function
    End If

  'Write a string to Private.Ini
    PrivPutString = kpWritePrivateProfileString(smSectionName, sEntryName, sValue, smIniFileName)

End Function

Function PrivPutTF (sEntryName As String, nValue As Integer)

  'Set an entry in Private.Ini to True/False
  'local vars
    Dim sTF As String

  'create INI string
    If nValue Then
      sTF = "true"
    Else
      sTF = "false"
    End If

  'write new value
    PrivPutTF = PrivPutString(sEntryName, sTF)

End Function

Function PrivSectExist () As Integer

  'Retrieve list of all [Section]'s
    Dim sSect As String
    sSect = PrivGetSections$()
    If Len(sSect) = 0 Then
      PrivSectExist = False
      Exit Function
    End If

  'Check for existence registered [Section]
    sSect = Chr$(0) + UCase$(sSect)
    If InStr(sSect, Chr$(0) + UCase$(smSectionName) + Chr$(0)) Then
      PrivSectExist = True
    Else
      PrivSectExist = False
    End If

End Function

Private Function StripComment$ (ByVal StrIn$)
  
  'Check for comment
    Dim nRet%
    nRet = InStr(StrIn, ";")

  'Remove it if present
    If nRet = 1 Then
      'Whole string is a comment
        StripComment = ""
        Exit Function
    ElseIf nRet > 1 Then
      'Strip comment
        StrIn = Left$(StrIn, nRet - 1)
    End If
  
  'Trim any trailing space
    StripComment = Trim$(StrIn)

End Function

Function SysDevAdd (sNewDev$, sComment$, sBAK$) As Integer
  
  'Setup some variables
    Dim sSysIni As String
    Dim sSysBak As String
    Dim sBuff() As String
    Dim sTemp As String
    Dim nRet As Integer
    Dim hFile As Integer
    Dim nCnt As Integer
    Dim fAdded As Integer

  'Find System.Ini, and make backup
    sTemp = String$(Max_EntryBuffer, 0)
    nRet = kpGetWindowsDirectory(sTemp, Max_EntryBuffer)
    sSysIni = Left$(sTemp, nRet) + "\System.Ini"
    If Len(Trim$(sBAK)) Then
      sSysBak = Left$(sTemp, nRet) + "\System." + sBAK
      On Local Error Resume Next
        FileCopy sSysIni, sSysBak
        If Err Then
          SysDevAdd = False
          Exit Function
        End If
      On Local Error GoTo 0
    End If

  'Read entire file, and insert new line
    hFile = FreeFile
    Open sSysIni For Input As hFile
    Do While Not EOF(hFile)
      nCnt = nCnt + 1
      ReDim Preserve sBuff(1 To nCnt)
      Line Input #hFile, sBuff(nCnt)
      If Not fAdded Then
        sTemp = UCase$(Trim$(sBuff(nCnt)))
        If sTemp = "[386ENH]" Then
          sTemp = Trim$(sNewDev)
          sComment = Trim$(sComment)
          If Len(sComment) Then
            sTemp = sTemp + "    ;" + sComment
          End If
          nCnt = nCnt + 1
          ReDim Preserve sBuff(1 To nCnt)
          sBuff(nCnt) = "device=" + sTemp
          fAdded = True
        End If
      End If
    Loop
    Close hFile

  'Write file back out
    hFile = FreeFile
    Open sSysIni For Output As hFile
    For nCnt = LBound(sBuff) To UBound(sBuff)
      Print #hFile, sBuff(nCnt)
    Next nCnt
    Close hFile

  'Make sure all went well
    SysDevAdd = SysDevLoaded(sNewDev)

End Function

Function SysDevGetList (sTable() As String) As Integer
  
  'Setup some variables
    Dim sSysIni As String
    Dim sBuff As String
    Dim nRet As Integer
    Dim hFile As Integer
    Dim nCnt As Integer

  'Example of usage, note return is one higher than UBound
  'Returned values *always* have paths, if present
    'Dim i%, n%
    'Dim eTable() As String
    'n% = SysDevGetList(eTable())
    'For i = 0 To n - 1
    '  Debug.Print "device="; eTable(i)
    'Next i
  
  'Find System.Ini
    sBuff = String$(Max_EntryBuffer, 0)
    nRet = kpGetWindowsDirectory(sBuff, Max_EntryBuffer)
    sSysIni = Left$(sBuff, nRet) + "\System.Ini"

  'Extract all device lines
    hFile = FreeFile
    Open sSysIni For Input As hFile
    Do While Not EOF(hFile)
      Line Input #hFile, sBuff
      sBuff = UCase$(Trim$(sBuff))
      If InStr(sBuff, "DEVICE=") = 1 Then
        ReDim Preserve sTable(0 To nCnt)
        sTable(nCnt) = StripComment$(Mid$(sBuff, 8))
        nCnt = nCnt + 1
      End If
    Loop
    Close hFile

  'Make final assignment
    SysDevGetList = nCnt

End Function

Function SysDevLoaded (ByVal sDevChk As String) As Integer

  'Set up some variables
    Dim nCnt As Integer
    Dim nLoop As Integer
    Dim dTable() As String
    Dim sTemp As String
    
  'Example of usage
    'SysIniRegister True   'Enforce path checking
    'If SysDevLoaded("VShare.386") Then
    '  MsgBox "VShare.386 *IS* Loaded!"
    'Else
    '  MsgBox "VShare.386 *NOT* Loaded!"
    'End If
  
  'Get list of all devices loaded
    nCnt = SysDevGetList(dTable())

  'Check for specific one
    For nLoop = 0 To nCnt - 1
      If nmSysPath Then
        sTemp = dTable(nLoop)
      Else
        sTemp = ExtractName$(dTable(nLoop), False)
        sDevChk = ExtractName$(sDevChk, False)
      End If
      If sTemp = UCase$(sDevChk) Then
        SysDevLoaded = True
        Exit For
      End If
    Next nLoop

End Function

Function SysDevRemove (ByVal sOldDev$, sBAK$) As Integer
  
  'Setup some variables
    Dim sSysIni As String
    Dim sSysBak As String
    Dim sBuff() As String
    Dim sTemp As String
    Dim nTempFlag As Integer
    Dim nRet As Integer
    Dim hFile As Integer
    Dim nCnt As Integer
    Dim fRemoved As Integer

  'Trim path off device if not comparing paths
    If Not nmSysPath Then
      sOldDev = ExtractName$(sOldDev, False)
    End If

  'Make sure it's there (somewhere)!
    nTempFlag = nmSysPath  'Store and temp set path flag
    SysIniRegister False
      nRet = SysDevLoaded(sOldDev)
    SysIniRegister nTempFlag
    If Not nRet Then       'Definately not there
      SysDevRemove = True
      Exit Function
    End If
  
  'Find System.Ini, and make backup
    sTemp = String$(Max_EntryBuffer, 0)
    nRet = kpGetWindowsDirectory(sTemp, Max_EntryBuffer)
    sSysIni = Left$(sTemp, nRet) + "\System.Ini"
    If Len(Trim$(sBAK)) Then
      sSysBak = Left$(sTemp, nRet) + "\System." + sBAK
      On Local Error Resume Next
        FileCopy sSysIni, sSysBak
        If Err Then
          SysDevRemove = False
          Exit Function
        End If
      On Local Error GoTo 0
    End If

  'Read entire file, and remove old device line
    hFile = FreeFile
    Open sSysIni For Input As hFile
    Do While Not EOF(hFile)
      nCnt = nCnt + 1
      ReDim Preserve sBuff(1 To nCnt)
      Line Input #hFile, sBuff(nCnt)
      If Not fRemoved Then
        sTemp = UCase$(Trim$(sBuff(nCnt)))
        If InStr(sTemp, "DEVICE=") = 1 Then
          'Get what follows & strip comments
          sTemp = StripComment$(Mid$(sTemp, 8))
          If Not nmSysPath Then 'Ignore path
            sTemp = ExtractName$(sTemp, False)
          End If
          If sTemp = UCase(sOldDev) Then
            nCnt = nCnt - 1
            ReDim Preserve sBuff(1 To nCnt)
            fRemoved = True
          End If
        End If
      End If
    Loop
    Close hFile

  'Write file back out
    hFile = FreeFile
    Open sSysIni For Output As hFile
    For nCnt = LBound(sBuff) To UBound(sBuff)
      Print #hFile, sBuff(nCnt)
    Next nCnt
    Close hFile

  'Make sure all went well
    If fRemoved Then
      nTempFlag = nmSysPath  'Store and temp set path flag
      SysIniRegister False
        nRet = SysDevLoaded(sOldDev)
      SysIniRegister nTempFlag
      SysDevRemove = Not nRet
    End If

End Function

Sub SysIniRegister (nPathFlag%)

  'Store module-level flag for future reference
    nmSysPath = nPathFlag

End Sub

Sub WinClearEntry (sEntryName As String)

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Sub
    End If

  'Sets a specific entry in Win.Ini to Nothing or Blank
    Dim nRetVal As Integer
    nRetVal = kpWriteProfileString(smWinSection, sEntryName, "")
    WinIniChanged

End Sub

Sub WinDeleteEntry (sEntryName As String)

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Sub
    End If

  'Deletes a specific entry in Win.Ini
    Dim nRetVal As Integer
    nRetVal = kpWriteProfileString(smWinSection, sEntryName, 0&)
    WinIniChanged

End Sub

Sub WinDeleteSection ()

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Sub
    End If

  'Deletes an *entire* [Section] and all its Entries in Win.Ini
    Dim nRetVal As Integer
    nRetVal = kpWriteProfileString(smWinSection, 0&, 0&)
  
  'Now Win.Ini needs to be reinitialized
    smWinSection = ""
    nmWinInit = False
    WinIniChanged

End Sub

Function WinGetInt (sEntryName As String, nDefaultValue As Integer) As Integer

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Function
    End If

  'Retrieves an Integer value from Win.Ini, range: 0-32767
    WinGetInt = kpGetProfileInt(smWinSection, sEntryName, nDefaultValue)

End Function

Function WinGetSectEntries () As String

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Function
    End If

  'Retrieves all Entries in a [Section] of Win.Ini
  'Entries nul terminated; last entry double-terminated
    Dim sTemp As String * Max_SectionBuffer
    Dim nRetVal As Integer
    nRetVal = kpGetProfileString(smWinSection, 0&, "", sTemp, Len(sTemp))
    WinGetSectEntries = Left$(sTemp, nRetVal + 1)

End Function

Function WinGetSectEntriesEx (sTable() As String) As Integer

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Function
    End If

  'Example of usage, note return is one higher than UBound
    'Dim i%, n%
    'Dim eTable() As String
    'WinIniRegister "Windows"
    'n% = WinGetSectionEntriesEx(eTable())
    'For i = 0 To n - 1
    '  Debug.Print eTable(0, i); "="; eTable(1, i)
    'Next i

  'Retrieves all Entries in a [Section] of Win.Ini
  'Entries nul terminated; last entry double-terminated
    Dim sBuff As String * Max_SectionBuffer
    Dim sTemp As String
    Dim nRetVal As Integer
    nRetVal = kpGetProfileString(smWinSection, 0&, "", sBuff, Len(sBuff))
    sTemp = Left$(sBuff, nRetVal + 1)

  'Parse entries into first dimension of table
  'and retrieve values into second dimension
    Dim nEntries As Integer
    Dim nNull As Integer
    On Error Resume Next
    Do While Asc(sTemp)
  'Bail if buffer wasn't large enough!!!
      If Err Then Exit Do
      ReDim Preserve sTable(0 To 1, 0 To nEntries)
      nNull = InStr(sTemp, Chr$(0))
      sTable(0, nEntries) = Left$(sTemp, nNull - 1)
      sTable(1, nEntries) = WinGetString(sTable(0, nEntries), "")
      sTemp = Mid$(sTemp, nNull + 1)
      nEntries = nEntries + 1
    Loop
  
  'Make final assignment
    WinGetSectEntriesEx = nEntries

End Function

Function WinGetSections$ ()

  'No real need to be initialized, Win.Ini *should* exist
  
  'Setup some variables
    Dim sWinIni As String
    Dim sRet As String
    Dim sBuff As String
    Dim hFile As Integer
    Dim nRet As Integer
  
  'Find Win.Ini
    sBuff = String$(Max_EntryBuffer, 0)
    nRet = kpGetWindowsDirectory(sBuff, Max_EntryBuffer)
    sWinIni = Left$(sBuff, nRet) + "\Win.Ini"

  'Extract all [Section] lines
    hFile = FreeFile
    Open sWinIni For Input As hFile
    Do While Not EOF(hFile)
      Line Input #hFile, sBuff
      sBuff = StripComment$(sBuff)
      If InStr(sBuff, "[") = 1 And InStr(sBuff, "]") = Len(sBuff) Then
        sRet = sRet + Mid$(sBuff, 2, Len(sBuff) - 2) + Chr$(0)
      End If
    Loop
    Close hFile

  'Assign return value
    If Len(sRet) Then
      WinGetSections = sRet + Chr$(0)
    Else
      WinGetSections = String$(2, 0)
    End If

End Function

Function WinGetSectionsEx (sTable() As String) As Integer

  'Get "normal" list of all [Section]'s
    Dim sSect As String
    sSect = WinGetSections$()
    If Len(sSect) = 0 Then
      WinGetSectionsEx = 0
      Exit Function
    End If

  'Parse [Section]'s into table
    Dim nEntries As Integer
    Dim nNull As Integer
    Do While Asc(sSect)
      ReDim Preserve sTable(0 To nEntries)
      nNull = InStr(sSect, Chr$(0))
      sTable(nEntries) = Left$(sSect, nNull - 1)
      sSect = Mid$(sSect, nNull + 1)
      nEntries = nEntries + 1
    Loop

  'Make function assignment
    WinGetSectionsEx = nEntries
  
End Function

Function WinGetString (sEntryName As String, ByVal sDefaultValue As String) As String

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Function
    End If

  'Retrieves Specific Entry from Win.Ini
    Dim sTemp As String * Max_EntryBuffer
    Dim nRetVal As Integer
    nRetVal = kpGetProfileString(smWinSection, sEntryName, sDefaultValue, sTemp, Len(sTemp))
    If nRetVal Then
      WinGetString = Left$(sTemp, nRetVal)
    End If

End Function

Function WinGetTF (sEntryName As String, nDefaultValue As Integer)
  
  'Retrieves Specific Entry as either True/False from Win.Ini
  'local vars
    Dim sTF As String
    Dim sDefault As String

  'get string value from INI
    If nDefaultValue Then
      sDefault = "true"
    Else
      sDefault = "false"
    End If
    sTF = WinGetString(sEntryName, sDefault)

  'interpret return string
    Select Case Trim$(UCase$(sTF))
      Case "YES", "Y", "TRUE", "T", "ON", "1", "-1"
        WinGetTF = True
      Case "NO", "N", "FALSE", "F", "OFF", "0"
        WinGetTF = False
      Case Else
        WinGetTF = False
    End Select

End Function

Private Sub WinIniChanged ()
  
  'Notify all other applications that Win.Ini has been changed
    Dim Rtn&
    Rtn = kpSendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal smWinSection)

End Sub

Sub WinIniFlushCache ()

  'Windows keeps a cached version of WIN.INI to improve performance.
  'If all three parameters are NULL, Windows flushes the cache.
    Dim nRetVal As Integer
    nRetVal = kpWriteProfileString(0&, 0&, 0&)
  
End Sub

Private Sub WinIniNotReg ()

  'Warn *PROGRAMMER* that there's a logic error!
    MsgBox "[Section] Not Registered in Win.Ini!", 16, "IniFile Logic Error"

End Sub

Sub WinIniRegister (sSectionName As String)
  
  'Store module-level for future reference
    smWinSection = Trim$(sSectionName)
    nmWinInit = True

End Sub

Function WinPutInt (sEntryName As String, nValue As Integer) As Integer

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Function
    End If

  'Write an integer to Win.Ini
    WinPutInt = kpWriteProfileString(smWinSection, sEntryName, Format$(nValue))
    WinIniChanged

End Function

Function WinPutString (sEntryName As String, ByVal sValue As String) As Integer

  'Bail if not initialized
    If Not nmWinInit Then
      WinIniNotReg
      Exit Function
    End If

  'Write a string to Win.Ini
    WinPutString = kpWriteProfileString(smWinSection, sEntryName, sValue)
    WinIniChanged

End Function

Function WinPutTF (sEntryName As String, nValue As Integer) As Integer
  
  'Set an entry in Win.Ini to True/False
  'local vars
    Dim sTF As String

  'create INI string
    If nValue Then
      sTF = "true"
    Else
      sTF = "false"
    End If

  'write new value
    WinPutTF = WinPutString(sEntryName, sTF)
    WinIniChanged

End Function

Function WinSectExist () As Integer

  'Retrieve list of all [Section]'s
    Dim sSect As String
    sSect = WinGetSections$()
    If Len(sSect) = 0 Then
      WinSectExist = False
      Exit Function
    End If

  'Check for existence registered [Section]
    sSect = Chr$(0) + UCase$(sSect)
    If InStr(sSect, Chr$(0) + UCase$(smWinSection) + Chr$(0)) Then
      WinSectExist = True
    Else
      WinSectExist = False
    End If

End Function

