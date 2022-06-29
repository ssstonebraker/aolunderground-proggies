VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MIMEAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'========================================================================
' Copyright 1999 - Digital Press, John Rhoton
'
' This program has been written to illustrate the Internet Mail protocols.
' It is provided free of charge and unconditionally.  However, it is not
' intended for production use, and therefore without warranty or any
' implication of support.
'
' You can find an explanation of the concepts behind this code in
' the book:  Programmer's Guide to Internet Mail by John Rhoton,
' Digital Press 1999.  ISBN: 1-55558-212-5.
'
' For ordering information please see http://www.amazon.com or
' you can order directly with http://www.bh.com/digitalpress.
'
'========================================================================

Option Explicit

Private Declare Function FindExecutable Lib "shell32.dll" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Const WorkDirectory = "C:\TEMP\"

Public Filename As String
Public Encoding As String

Public Lines As New Collection

Const Base64Sequence = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Sub CallAttachment()
  Dim strExecutablePath As String
  Dim intScanPosition As Integer
  Dim intStatus As Integer
  
  SaveFile
  
  strExecutablePath = Space$(255)
  intStatus = FindExecutable(Filename, WorkDirectory, strExecutablePath)
  
  For intScanPosition = 1 To Len(strExecutablePath)
      If Mid$(strExecutablePath, intScanPosition, 1) = Chr$(0) Then
          strExecutablePath = Left$(strExecutablePath, intScanPosition - 1) + _
                           " " + Mid$(strExecutablePath, intScanPosition + 1)
      End If
  Next
  
  strExecutablePath = Trim(strExecutablePath)
  
  If InStr(1, strExecutablePath, Filename) = 0 Then
    strExecutablePath = strExecutablePath + " " + WorkDirectory + Filename
  End If
  
  Debug.Print strExecutablePath

  Shell strExecutablePath, vbNormalFocus

End Sub

Public Sub SaveFile()
  Dim intOutputFileNumber As Integer
  Dim strCurrentLine
  
  intOutputFileNumber = FreeFile
  
  Select Case LCase(Encoding)
  
    Case "base64":
      Open WorkDirectory + Filename For Binary As #intOutputFileNumber
      For Each strCurrentLine In Lines
      Put #intOutputFileNumber, , DecodeBase64(CStr(strCurrentLine))
      Next
      Close #intOutputFileNumber
    Case "quoted-printable":
      Open WorkDirectory + Filename For Output As #intOutputFileNumber
      For Each strCurrentLine In Lines
      Print #intOutputFileNumber, DecodeQuotedPrintable(CStr(strCurrentLine))
      Next
      Close #intOutputFileNumber
    Case Else
      Open WorkDirectory + Filename For Output As #intOutputFileNumber
      For Each strCurrentLine In Lines
      Print #intOutputFileNumber, CStr(strCurrentLine)
      Next
      Close #intOutputFileNumber
  End Select
      
End Sub

Private Function DecodeBase64(ContentLine As String) As String
  Dim intFoursomePosition As Integer
  Dim strWorkBuffer As String
  Dim strFoursome As String
  Dim bytWork1 As Byte
  Dim bytWork2 As Byte
  Dim bytWork3 As Byte
  Dim bytWork4 As Byte
  
  
  For intFoursomePosition = 0 To Len(ContentLine) / 4 - 1
  
    strFoursome = Mid$(ContentLine, (intFoursomePosition * 4 + 1), 4)
    Base64DecodeFoursome strFoursome, _
                         bytWork1, bytWork2, bytWork3
        
    If Mid$(strFoursome, 3, 2) = "==" Then
      strWorkBuffer = strWorkBuffer + Chr$(bytWork1)
    ElseIf Mid$(strFoursome, 4, 1) = "=" Then
      strWorkBuffer = strWorkBuffer + Chr$(bytWork1) + Chr$(bytWork2)
    Else
      strWorkBuffer = strWorkBuffer + Chr$(bytWork1) + Chr$(bytWork2) + Chr$(bytWork3)
    End If
    
  Next
      
  DecodeBase64 = strWorkBuffer

End Function

Private Sub Base64DecodeFoursome(InputString As String, _
                                 ByRef OutputByte1 As Byte, _
                                 ByRef OutputByte2 As Byte, _
                                 ByRef OutputByte3 As Byte)
  Dim bytBits01to06 As Byte      ' First six bits of first byte
  Dim bytBits07to08 As Byte      ' Last two bits of first byte
  Dim bytBits09to12 As Byte      ' First four bits of second byte
  Dim bytBits13to16 As Byte      ' Last four bits of second byte
  Dim bytBits17to18 As Byte      ' First two bits of third byte
  Dim bytBits19to24 As Byte      ' Last six bits of third byte
  Dim bytIndex1 As Byte
  Dim bytIndex2 As Byte
  Dim bytIndex3 As Byte
  Dim bytIndex4 As Byte
  
  bytIndex1 = Base64Index(Mid$(InputString, 1, 1))
  bytIndex2 = Base64Index(Mid$(InputString, 2, 1))
  bytIndex3 = Base64Index(Mid$(InputString, 3, 1))
  bytIndex4 = Base64Index(Mid$(InputString, 4, 1))
  
        ' 4 is two bits
        ' 16 is four bits
        ' 64 is six bits
  bytBits01to06 = bytIndex1
  bytBits07to08 = Int(bytIndex2 / 16)
  bytBits09to12 = bytIndex2 And 15
  bytBits13to16 = Int(bytIndex3 / 4)
  bytBits17to18 = bytIndex3 And 3
  bytBits19to24 = bytIndex4
  
  OutputByte1 = bytBits01to06 * 4 + bytBits07to08
  OutputByte2 = bytBits09to12 * 16 + bytBits13to16
  OutputByte3 = bytBits17to18 * 64 + bytBits19to24
    
End Sub


Private Function Base64Index(EntryString As String) As Byte

  If Len(EntryString) > 1 Then
    MsgBox ("Invalid Base64 characters " + EntryString)
    Exit Function
  End If
      
'  Base64Index = InStr(EntryString, Base64Sequence, vbBinaryCompare) - 1
  Base64Index = InStr(1, Base64Sequence, EntryString, vbBinaryCompare) - 1
  
  If EntryString = "=" Then Base64Index = 0
  If Base64Index < 0 Then
      MsgBox ("Invalid Base64 character of ASCII: " + Str$(Asc(EntryString)))
  End If

End Function

Private Function DecodeQuotedPrintable(ContentLine As String) As String
  Dim strWorkLine As String
  Dim intScanPosition As Integer
  Dim bytSpecialCharacter As Byte
  
  intScanPosition = 1
  Do While intScanPosition <= Len(ContentLine)
    If Mid$(ContentLine, intScanPosition, 1) = "=" Then
      bytSpecialCharacter = Val("&H" + Mid$(ContentLine, intScanPosition + 1, 2))
      strWorkLine = strWorkLine + Chr$(bytSpecialCharacter)
      intScanPosition = intScanPosition + 3
    Else
      strWorkLine = strWorkLine + Mid$(ContentLine, intScanPosition, 1)
      intScanPosition = intScanPosition + 1
    End If
  Loop
  
  DecodeQuotedPrintable = strWorkLine

End Function

Public Sub ReadFile(FilePath As String, Encoding As String)
  Select Case LCase(Encoding)
    Case "base64":
      EncodeFileBase64 (FilePath)
    Case "quoted-printable":
      EncodeFileQuotedPrintable (FilePath)
    Case Else
      ReadFilePlainText (FilePath)
  End Select
End Sub

Public Sub ReadFilePlainText(FilePath As String)
  Dim intFileNumber As Integer
  Dim strTextLine As String

  intFileNumber = FreeFile
  Open FilePath For Input As #intFileNumber
  
  Do While Not EOF(intFileNumber)              ' Loop until end of file.
    Line Input #intFileNumber, strTextLine        ' Read line into variable.
    Lines.Add strTextLine
  Loop
  
  Close #intFileNumber

End Sub

Public Sub EncodeFileQuotedPrintable(FilePath As String)
  Dim intFileNumber As Integer
  Dim strInputLine As String
  Dim strOutputLine As String
  Dim intScanIndex As Integer
  Dim strScanValue As String
  Dim strHexString As String

  intFileNumber = FreeFile
  Open FilePath For Input As #intFileNumber
  
  Do While Not EOF(intFileNumber)              ' Loop until end of file.
    Line Input #intFileNumber, strInputLine    ' Read line into variable.
    strOutputLine = ""
    For intScanIndex = 1 To Len(strInputLine)
      strScanValue = Asc(Mid(strInputLine, intScanIndex, 1))
      If strScanValue > 127 Or strScanValue = 61 Then
        strHexString = Hex(strScanValue)
        If Len(strHexString) = 1 Then strHexString = "0" + strHexString
        strOutputLine = strOutputLine + "=" + strHexString
      Else
        strOutputLine = strOutputLine + Chr(strScanValue)
      End If
    Next
    
    Lines.Add strOutputLine
  Loop
  
  Close #intFileNumber

End Sub

Public Sub EncodeFileBase64(FilePath As String)
  Dim intInputFileNumber As Integer

  Dim bytInput1 As Byte
  Dim bytInput2 As Byte
  Dim bytInput3 As Byte
  Dim strFoursome As String * 4
  Dim strBase64Line As String ' * 76
  Dim intTripletsCount As Integer
  Dim intRestBytes As Integer


  intInputFileNumber = FreeFile
  Open FilePath For Binary Access Read As #intInputFileNumber
  
  intTripletsCount = LOF(intInputFileNumber) / 3
  intRestBytes = LOF(intInputFileNumber) - Int(intTripletsCount) * 3
  bytInput1 = 0
  bytInput2 = 0
  bytInput3 = 0
     
  strBase64Line = ""
  
  Dim Counter
  For Counter = 1 To intTripletsCount ' Loop until end of file.
    Get #intInputFileNumber, , bytInput1 ' Read line into variable.
    Get #intInputFileNumber, , bytInput2 ' Read line into variable.
    Get #intInputFileNumber, , bytInput3 ' Read line into variable.
    strFoursome = Base64Encode(bytInput1, bytInput2, bytInput3)
    Debug.Print strFoursome
    strBase64Line = strBase64Line + strFoursome
    If Len(strBase64Line) > 72 Then
      Lines.Add strBase64Line
      strBase64Line = ""
    End If
  Next
      
  If intRestBytes = 1 Then
    Get #intInputFileNumber, , bytInput1 ' Read line into variable.
    strFoursome = Base64Encode(bytInput1, 0, 0)
    Debug.Print strFoursome
    strBase64Line = strBase64Line + Left$(strFoursome, 2) + "=="
  ElseIf intRestBytes = 2 Then
    Get #intInputFileNumber, , bytInput1 ' Read line into variable.
    Get #intInputFileNumber, , bytInput2 ' Read line into variable.
    strFoursome = Base64Encode(bytInput1, bytInput2, 0)
    Debug.Print strFoursome
    strBase64Line = strBase64Line + Left$(strFoursome, 3) + "="
  End If

  Close #intInputFileNumber

  If Len(strBase64Line) > 0 Then
    Lines.Add strBase64Line
    strBase64Line = ""
  End If
     
End Sub

Private Function Base64Encode(InputByte1 As Byte, _
                              InputByte2 As Byte, _
                              InputByte3 As Byte) As String
  Dim bytBits01to06 As Byte      ' First six bits of first byte
  Dim bytBits07to08 As Byte      ' Last two bits of first byte
  Dim bytBits09to12 As Byte      ' First four bits of second byte
  Dim bytBits13to16 As Byte      ' Last four bits of second byte
  Dim bytBits17to18 As Byte      ' First two bits of third byte
  Dim bytBits19to24 As Byte      ' Last six bits of third byte

  Dim bytIndex1 As Byte
  Dim bytIndex2 As Byte
  Dim bytIndex3 As Byte
  Dim bytIndex4 As Byte

      ' 4 is two bits
      ' 16 is four bits
      ' 64 is six bits
  bytBits01to06 = Int(InputByte1 / 4)
  bytBits07to08 = InputByte1 And 3
  
  bytBits09to12 = Int(InputByte2 / 16)
  bytBits13to16 = InputByte2 And 15
  
  bytBits17to18 = Int(InputByte3 / 64)
  bytBits19to24 = InputByte3 And 63
  
  bytIndex1 = bytBits01to06
  bytIndex2 = bytBits07to08 * 16 + bytBits09to12
  bytIndex3 = bytBits13to16 * 4 + bytBits17to18
  bytIndex4 = bytBits19to24
  
  Base64Encode = Base64Table(bytIndex1) + _
                 Base64Table(bytIndex2) + _
                 Base64Table(bytIndex3) + _
                 Base64Table(bytIndex4)

End Function


Private Function Base64Table(Index As Byte) As String
  Base64Table = Mid(Base64Sequence, Index + 1, 1)
End Function


