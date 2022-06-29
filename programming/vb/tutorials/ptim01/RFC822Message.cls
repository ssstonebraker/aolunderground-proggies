VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RFC822Message"
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
Option Base 1
Option Compare Text

Public Lines As New Collection
Public Headers As New Collection
Public BodyLines As New Collection
Dim bolEndOfMessage As Boolean
Private intLineIndex As Integer

Public Sub OpenFile(FilePath As String)
  
  Dim intInputFileNumber As Integer
  Dim strCurrentLine As String

  intInputFileNumber = FreeFile
  
  Do While Dir(FilePath) = ""
    FilePath = InputBox("Name of file to display")
  Loop
  
  Open FilePath For Input As #intInputFileNumber
  Do While Not EOF(intInputFileNumber)
    Line Input #intInputFileNumber, strCurrentLine
    Lines.Add strCurrentLine
  Loop
  Close #intInputFileNumber

End Sub

Public Sub Parse()
  Dim strBoundary As String
  Dim strCurrentLine As String
  Dim rhdrCurrentHeader As RFC822HeaderLine
  
  strBoundary = ""
  
    ' Headers
  Do While Not bolEndOfMessage And NextLine <> "" '
  
    strCurrentLine = GetLine
    Do While SpaceBegin(NextLine)
      strCurrentLine = strCurrentLine + GetLine
    Loop
    
    Set rhdrCurrentHeader = New RFC822HeaderLine
    rhdrCurrentHeader.Parse (strCurrentLine)
    Headers.Add rhdrCurrentHeader
    
    If LCase(rhdrCurrentHeader.Label) = "content-type" Then
      strBoundary = rhdrCurrentHeader.HeaderParameter("strBoundary")
    End If
  
  Loop
  
    ' Blank Lines
  Do While NextLine = ""
    strCurrentLine = GetLine:            If bolEndOfMessage Then Exit Sub
  Loop
  
  Do While Not bolEndOfMessage
    strCurrentLine = GetLine
    BodyLines.Add strCurrentLine
  Loop
End Sub

Private Function SpaceBegin(HeaderLine As String) As Boolean
                
    Select Case Left$(HeaderLine, 1)
        Case " ", Chr$(9)
            SpaceBegin = True
        Case Else
            SpaceBegin = False
    End Select

End Function

Private Function GetLine() As String

  If intLineIndex > 190 Then
    Debug.Print intLineIndex
  End If

  intLineIndex = intLineIndex + 1
  GetLine = Lines(intLineIndex)
  If intLineIndex >= Lines.Count Then bolEndOfMessage = True

End Function

Private Function NextLine() As String
    
  If intLineIndex < Lines.Count Then
    NextLine = Lines(intLineIndex + 1)
  Else
    NextLine = ""
  End If

End Function


Public Sub Compose()
  Dim rhdrCurrentHeader
  
  For Each rhdrCurrentHeader In Headers
    Lines.Add CStr(rhdrCurrentHeader.Label) + ": " + CStr(rhdrCurrentHeader.Value)
  Next
  
  Lines.Add ""
  Dim strCurrentLine
  For Each strCurrentLine In BodyLines
    Lines.Add CStr(strCurrentLine)
  Next

End Sub


Public Sub AddHeader(HeaderLabel As String, HeaderValue)
  Dim rhdrCurrentHeader As New RFC822HeaderLine

  rhdrCurrentHeader.Label = HeaderLabel
  rhdrCurrentHeader.Value = HeaderValue
  Headers.Add rhdrCurrentHeader

End Sub

