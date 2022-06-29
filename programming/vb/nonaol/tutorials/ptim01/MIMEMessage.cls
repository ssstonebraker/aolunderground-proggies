VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MIMEMessage"
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
Public Attachments As New Collection
Public PreambleLines As New Collection
Public EpilogueLines As New Collection
Dim EndOfMessage As Boolean
Private LineIndex As Integer

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
  Dim mattCurrentAttachment As MIMEAttachment
  
  strBoundary = ""
  
    ' Headers
  Do While Not EndOfMessage And NextLine <> "" '
  
    strCurrentLine = GetLine
    Do While SpaceBegin(NextLine)
      strCurrentLine = strCurrentLine + GetLine
    Loop
    
    Set rhdrCurrentHeader = New RFC822HeaderLine
    rhdrCurrentHeader.Parse (strCurrentLine)
    Headers.Add rhdrCurrentHeader
    
    If LCase(rhdrCurrentHeader.Label) = "content-type" Then
      strBoundary = rhdrCurrentHeader.HeaderParameter("Boundary")
    End If
  
  Loop
  
    ' Blank Lines
  Do While NextLine = ""
    strCurrentLine = GetLine:            If EndOfMessage Then Exit Sub
  Loop
  
  If strBoundary = "" Then     ' Non-MIME
    Do While Not EndOfMessage
      strCurrentLine = GetLine
      BodyLines.Add strCurrentLine
    Loop
    Exit Sub
  End If
  
  ' MIME
      ' Preamble
  Do While InStr(NextLine, strBoundary) = 0
    strCurrentLine = GetLine:          If EndOfMessage Then Exit Sub
    PreambleLines.Add strCurrentLine
  Loop
  
    ' Attachments
  Do While InStr(NextLine, strBoundary + "--") = 0
      ' Boundary
    strCurrentLine = GetLine:          If EndOfMessage Then Exit Sub
        
      Set mattCurrentAttachment = New MIMEAttachment
        
        ' Attachment Headers
      Do While NextLine <> ""
          ' Add wrapped lines
        strCurrentLine = GetLine:      If EndOfMessage Then Exit Sub
        Do While SpaceBegin(NextLine)
          strCurrentLine = strCurrentLine + GetLine
        Loop
        
        rhdrCurrentHeader.Parse (strCurrentLine)
        If LCase(rhdrCurrentHeader.Label) = "content-type" Then
          mattCurrentAttachment.Filename = rhdrCurrentHeader.HeaderParameter("name")
        End If
                  
        If LCase(rhdrCurrentHeader.Label) = "content-transfer-encoding" Then
          mattCurrentAttachment.Encoding = rhdrCurrentHeader.Value
        End If
      Loop
    
        ' Blank Lines
      Do While NextLine = ""
        strCurrentLine = GetLine:            If EndOfMessage Then Exit Sub
      Loop
        
      If mattCurrentAttachment.Filename = "" Then
          ' Message Body
        Set mattCurrentAttachment = Nothing
        Do While InStr(NextLine, strBoundary) = 0
          strCurrentLine = GetLine:          If EndOfMessage Then Exit Sub
          BodyLines.Add strCurrentLine
        Loop
      Else
          ' Attachment body
        Do While InStr(NextLine, strBoundary) = 0
          strCurrentLine = GetLine:          If EndOfMessage Then Exit Sub
          mattCurrentAttachment.Lines.Add strCurrentLine
        Loop
        Attachments.Add mattCurrentAttachment
      End If
  Loop
  
    ' Epilogue
  Do While Not EndOfMessage
    strCurrentLine = GetLine
    EpilogueLines.Add strCurrentLine
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
  If LineIndex > 190 Then
    Debug.Print LineIndex
  End If

  LineIndex = LineIndex + 1
  GetLine = Lines(LineIndex)
  If LineIndex >= Lines.Count Then EndOfMessage = True
End Function

Private Function NextLine() As String
  If LineIndex < Lines.Count Then
    NextLine = Lines(LineIndex + 1)
  Else
    NextLine = ""
  End If
End Function


Public Sub Compose()
  Dim strAttachmentBoundary As String
  Dim rhdrCurrentHeader
  Dim mattCurrentAttachment
  
  Randomize
  strAttachmentBoundary = "--Boundary" + Trim(Str(Rnd)) + Trim(Str(Rnd)) + "--"
  
  For Each rhdrCurrentHeader In Headers
    Lines.Add CStr(rhdrCurrentHeader.Label) + ": " + CStr(rhdrCurrentHeader.Value)
  Next
  
  Lines.Add "MIME-Version: 1.0"
  Lines.Add "Content-Type: multipart/mixed; " _
            + "boundary=""" + strAttachmentBoundary + """"
  Lines.Add ""
  
  Lines.Add "--" + strAttachmentBoundary
  Lines.Add ""
  Dim CurrentLine
  For Each CurrentLine In BodyLines
    Lines.Add CStr(CurrentLine)
  Next
  Lines.Add ""
  
  For Each mattCurrentAttachment In Attachments
    Lines.Add "--" + strAttachmentBoundary
    If mattCurrentAttachment.Encoding = "base64" Then
      Lines.Add "Content-Type: application/octet-stream; name=""" + Dir(mattCurrentAttachment.Filename) + """"
      Lines.Add "Content-Transfer-Encoding: base64"
    ElseIf mattCurrentAttachment.Encoding = "quoted-printable" Then
      Lines.Add "Content-Type: text/plain; name=""" + Dir(mattCurrentAttachment.Filename) + """"
      Lines.Add "Content-Transfer-Encoding: quoted-printable"
    Else
      Lines.Add "Content-Type: text/plain; name=""" + Dir(mattCurrentAttachment.Filename) + """"
      Lines.Add "Content-Transfer-Encoding: 7bit"
    End If
    Lines.Add ""
    For Each CurrentLine In mattCurrentAttachment.Lines
      Lines.Add CStr(CurrentLine)
    Next
    Lines.Add ""
  Next
  
  Lines.Add "--" + strAttachmentBoundary + "--"

End Sub


Public Sub AddHeader(HeaderLabel As String, HeaderValue)
  Dim rhdrCurrentHeader As New RFC822HeaderLine

  rhdrCurrentHeader.Label = HeaderLabel
  rhdrCurrentHeader.Value = HeaderValue
  Headers.Add rhdrCurrentHeader

End Sub


Public Sub AddAttachment(FilePath As String, Encoding As String)
  Dim mattCurrentAttachment As MIMEAttachment

  Set mattCurrentAttachment = New MIMEAttachment
  mattCurrentAttachment.Encoding = Encoding
  mattCurrentAttachment.Filename = FilePath
  mattCurrentAttachment.ReadFile FilePath, Encoding
  Attachments.Add mattCurrentAttachment

End Sub
