VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFile"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

'-- Size of buffer when reading file
'   Default = 32K, Max = 64K
Public BlockSize    As Long

'-- Error number
Public ErrorNum     As Long

'-- Error message
Public ErrorMsg     As String



'-- Contains the file data
Private szData()    As String

'-- Number of lines
Private lLines      As Long

'-- Holds the current line when searching
Private lCurLine    As Long

'-- Holds the current position in the current
'   line when searching
Private lCurPos     As Long

'-- String being searched
Private szSearch    As String

'-- Case of search term
Private nSearchCase As Integer

'-- Bad input file name specified
Private Const ERROR_BAD_FILENAME = 1
'-- No data to save
Private Const ERROR_NO_DATA = 2
'-- No file specified when saving
Private Const ERROR_NO_FILE_SPECIFIED = 3
'-- Could not write to file
Private Const ERROR_FILE_WRITE = 4
'-- Error creating new data
Private Const ERROR_CREATE_NEW = 5

Public Sub CreateNew(lLineCount As Long)

    Me.ErrorNum = False
    Me.ErrorMsg = ""
    
    On Error Resume Next
    ReDim szData(1 To lLineCount) As String
    If Err Then
        Me.ErrorNum = vbObjectError + ERROR_CREATE_NEW
        Me.ErrorMsg = Error
        Exit Sub
    End If
    
    lLines = lLineCount
    
End Sub


Public Property Let FoundPos(ByVal nDummy As Integer)

    '-- Do not let the user set the search position.

End Property

Public Property Get FoundPos() As Integer

    '-- Return the current position within the current line of
    '   the searched for and found text.
    FoundPos = lCurPos
    
End Property

Public Property Get Line(ByVal lIndex As Long) As String
'-- Retreives a line of text from the file.

    '-- Trap errors
    On Error Resume Next
    Line = szData(lIndex)

End Property


Public Property Let Line(ByVal lIndex As Long, ByVal szText As String)

    '-- Trim Cr and LF chars
    If Right(szText, 1) = vbLf Then
        szText = Left$(szText, Len(szText) - 1)
    End If
    
    If Right(szText, 1) = vbCr Then
        szText = Left$(szText, Len(szText) - 1)
    End If
    
    '-- Return the array element (no CR/LF)
    szData(lIndex) = szText

End Property
Public Property Get Lines() As Long

    '-- Number of lines
    Lines = lLines
    
End Property

Public Property Let Lines(ByVal lDummy As Long)

    '-- Do not allow the user to set the number of lines.

End Property

Public Sub Load(ByVal szFileName As String)
'-- Load the contents of a text file into memory
'   This routine will handle any line that ends with
'   a carriage return, a linefeed, or both.
    
    Dim szBuffer    As String
    Dim lFileNum    As Long
    Dim lFileLen    As Long
    Dim lNumBlocks  As Long
    Dim lRemainder  As Long
    Dim lIndex      As Long
    Dim lPos        As Long
    
    Me.ErrorNum = False
    Me.ErrorMsg = ""
    
    '-- Open the file
    lFileNum = FreeFile
    Open szFileName For Binary As lFileNum
    lFileLen = LOF(lFileNum)
    
    '-- Does the file exist?
    If lFileLen = 0 Then
        Close lFileNum
        Me.ErrorNum = vbObjectError + ERROR_BAD_FILENAME
        Me.ErrorMsg = "File Does Not Exist"
        Exit Sub
    End If
    
    '-- Clear the current array
    Erase szData
    lLines = 0
    
    '-- Get the buffer size
    If BlockSize = 0 Then
        BlockSize = 32768
    ElseIf BlockSize > 65535 Then
        BlockSize = 65535
    End If
    
    '-- Get the number of blocks
    lNumBlocks = lFileLen \ BlockSize
    
    '-- Anything left over?
    lRemainder = lFileLen Mod BlockSize
    
    '-- Read and process each block
    For lIndex = 1 To lNumBlocks
        szBuffer = Space$(BlockSize)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    Next
    
    '-- Process whatever's left
    If lRemainder Then
        szBuffer = Space$(lRemainder)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    End If

    '-- Close the file and exit
    Close lFileNum

    Exit Sub

ProcessData:
    
    Do
        '-- Find the next CR
        lPos = InStr(szBuffer, vbCr)
        If lPos Then
            '-- Copy the text up to the CRLF into szData
            lLines = lLines + 1
            ReDim Preserve szData(1 To lLines) As String
            szData(lLines) = Left$(szBuffer, lPos - 1)
            '-- If the next character is a linefeed, skip over it.
            If Asc(Mid$(szBuffer, lPos + 1, 1)) = 10 Then
                szBuffer = Mid$(szBuffer, lPos + 2)
            Else
                szBuffer = Mid$(szBuffer, lPos + 1)
            End If
        Else
            '-- Find the next LF
            lPos = InStr(szBuffer, vbLf)
            If lPos Then
                '-- Copy the text up to the CRLF into szData
                lLines = lLines + 1
                ReDim Preserve szData(1 To lLines) As String
                szData(lLines) = Left$(szBuffer, lPos - 1)
                szBuffer = Mid$(szBuffer, lPos + 1)
            Else
                '-- No more CRLFs. More data?
                If Len(szBuffer) Then
                    '-- Yep. Add it to szData
                    lLines = lLines + 1
                    ReDim Preserve szData(1 To lLines) As String
                    szData(lLines) = szBuffer
                End If
                '-- No more data. Exit the loop
                Exit Do
            End If
        End If
    Loop
    
    Return

End Sub
Public Sub LoadListBox(ByVal szFileName As String, List1 As Control)
'-- Loads the contents of a text file into a list box
'   This routine will handle any line that ends with
'   a carriage return, a linefeed, or both.
    
    Dim szBuffer    As String
    Dim lFileNum    As Long
    Dim lFileLen    As Long
    Dim lNumBlocks  As Long
    Dim lRemainder  As Long
    Dim lIndex      As Long
    Dim lPos        As Long
    
    Me.ErrorNum = False
    Me.ErrorMsg = ""
    
    '-- Open the file
    lFileNum = FreeFile
    Open szFileName For Binary As lFileNum
    lFileLen = LOF(lFileNum)
    
    '-- Does the file exist?
    If lFileLen = 0 Then
        Close lFileNum
        Me.ErrorNum = vbObjectError + ERROR_BAD_FILENAME
        Me.ErrorMsg = "File Does Not Exist"
        Exit Sub
    End If
    
    '-- Clear the List Box
    List1.Clear
    
    '-- Get the buffer size
    If BlockSize = 0 Then
        BlockSize = 32768
    ElseIf BlockSize > 65535 Then
        BlockSize = 65535
    End If
    
    '-- Get the number of blocks
    lNumBlocks = lFileLen \ BlockSize
    
    '-- Anything left over?
    lRemainder = lFileLen Mod BlockSize
    
    '-- Read and process each block
    For lIndex = 1 To lNumBlocks
        szBuffer = Space$(BlockSize)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    Next
    
    '-- Process whatever's left
    If lRemainder Then
        szBuffer = Space$(lRemainder)
        Get #lFileNum, , szBuffer
        GoSub ProcessData
    End If

    '-- Close the file and exit
    Close lFileNum

    Exit Sub

ProcessData:
    
    Do
        '-- Find the next CR
        lPos = InStr(szBuffer, vbCr)
        If lPos Then
            '-- Copy the text up to the CRLF into List1
            List1.AddItem Left$(szBuffer, lPos - 1)
            '-- If the next character is a linefeed, skip over it.
            If Asc(Mid$(szBuffer, lPos + 1, 1)) = 10 Then
                szBuffer = Mid$(szBuffer, lPos + 2)
            Else
                szBuffer = Mid$(szBuffer, lPos + 1)
            End If
        Else
            '-- Find the next LF
            lPos = InStr(szBuffer, vbLf)
            If lPos Then
                '-- Copy the text up to the CRLF into List1
                List1.AddItem Left$(szBuffer, lPos - 1)
                szBuffer = Mid$(szBuffer, lPos + 1)
            Else
                '-- No more CRLFs. More data?
                If Len(szBuffer) Then
                    '-- Yep. Add it to the list box
                    List1.AddItem szBuffer
                End If
                '-- No more data. Exit the loop
                Exit Do
            End If
        End If
    Loop
    
    Return

End Sub

Public Sub Save(ByVal szFileName As String)

    Dim szBuffer    As String
    Dim lFileNum    As Long
    Dim lFileLen    As Long
    Dim lNumBlocks  As Long
    Dim lRemainder  As Long
    Dim lIndex      As Long
    Dim lPos        As Long
        
    '-- Any data in this object?
    If lLines = 0 Then
        Me.ErrorNum = vbObjectError + ERROR_NO_DATA
        Me.ErrorMsg = "Save: There is nothing to save"
        Exit Sub
    End If
    
    '-- Was a filename specified?
    If Len(szFileName) = 0 Then
        Me.ErrorNum = vbObjectError + ERROR_NO_FILE_SPECIFIED
        Me.ErrorMsg = "Save: No File Specified"
        Exit Sub
    End If
    
    '-- Clear the errors
    Me.ErrorNum = 0
    Me.ErrorMsg = ""
    
    '-- Open the file
    lFileNum = FreeFile
    On Error Resume Next
    Open szFileName For Binary As lFileNum
    If Err Then
        Me.ErrorNum = vbObjectError + ERROR_FILE_WRITE
        Me.ErrorMsg = "Save: Could Not Write File " & szFileName
    End If

    '-- Get the buffer size
    If BlockSize = 0 Then
        BlockSize = 32768
    ElseIf BlockSize > 65535 Then
        BlockSize = 65535
    End If
    
    '-- Get the number of blocks
    lNumBlocks = lFileLen \ BlockSize
    
    '-- Process each block
    For lIndex = 1 To lLines
        '-- make sure to add the CR/LF
        szBuffer = szBuffer & szData(lIndex) & vbCrLf
        If Len(szBuffer) >= BlockSize Then
            Put #lFileNum, , szBuffer
            szBuffer = ""
        End If
    Next
    
    '-- Process whatever's left
    If Len(szBuffer) Then
        Put #lFileNum, , szBuffer
    End If

    '-- Close the file and exit
    Close lFileNum

End Sub

Public Function FindFirst(ByVal szText As String, ByVal nCase As Integer) As Long
'-- Searches for text in the file and returns the
'   line number where a match was found.

    Dim lIndex      As Long
    Dim lPos        As Long
    
    lCurPos = 0
    nSearchCase = nCase

    '-- Set case-sensitivity
    If nCase Then
        szSearch = szText
    Else
        szSearch = UCase$(szText)
    End If

    '-- Search from the top
    For lIndex = 1 To lLines
        If nCase Then
            lPos = InStr(szData(lIndex), szSearch)
        Else
            lPos = InStr(UCase$(szData(lIndex)), szSearch)
        End If
        If lPos Then
            lCurLine = lIndex
            FindFirst = lCurLine
            lCurPos = lPos
            Exit For
        End If
    Next

End Function


Public Function FindNext() As Long
'-- Searches for text in the file and returns the
'   line number where a match was found.

    Dim lIndex      As Long
    Dim szTemp      As String
    Dim lPos        As Long
    

    '-- Search from the current line
    For lIndex = lCurLine + 1 To lLines
        '-- If we're in the middle of a line,
        '   search from the end of the last find.
        If lCurPos Then
            szTemp = Mid$(szData(lIndex), lCurPos + Len(szSearch))
        Else
            '-- Normal search
            szTemp = szData(lIndex)
        End If
        
        '-- Search
        If nSearchCase Then
            lPos = InStr(szTemp, szSearch)
        Else
            lPos = InStr(UCase$(szTemp), szSearch)
        End If
        If lPos Then
            '-- Set the current line
            lCurLine = lIndex
            
            '-- Set the current position within the line
            If lCurPos Then
                lCurPos = lCurPos + Len(szSearch) + lPos
            Else
                lCurPos = lPos
            End If
            
            '-- Return the current line
            FindNext = lCurLine
            Exit Function
        Else
            '-- No match
            'szSearch = ""
            'lCurLine = 0
            lCurPos = 0
        End If
    Next

End Function

Private Sub Class_Initialize()

    '-- Initialize the block size to 32K
    BlockSize = 32768

End Sub


