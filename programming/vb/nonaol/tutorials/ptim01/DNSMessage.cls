VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DNSMessage"
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


' Header    RFC 1035 4.1.1

Public ID As Integer       ' QueryIdentifyer
Public QR As Integer       ' QueryOrResponse
Public OPCODE As Integer   ' QueryType
Public AA As Integer       ' AuthoritativeAnswer
Public TC As Integer       ' Truncation
Public RD As Integer       ' RecursionDesired
Public RA As Integer       ' RecursionAvailable
Public Z As Integer        ' UnusedZero
Public RCODE As Integer    ' ResponseCode

Public QDCOUNT As Integer  ' QuestionCount
Public ANCOUNT As Integer  ' AnswerCount
Public NSCOUNT As Integer  ' NameServerCount
Public ARCOUNT As Integer  ' AdditionalResourceCount


' -----------------------------------------------------

Private dqstQuestion() As DNSQuestion
Private drrAnswer() As DNSResourceRecord

Public Property Get Question(Index As Integer) As DNSQuestion
  Set Question = dqstQuestion(Index)
End Property

Public Property Get Answer(Index As Integer) As DNSResourceRecord
  Set Answer = drrAnswer(Index)
End Property

Public Property Get TransferString() As String

  Dim HeaderSection As String
  Dim QuestionSection As String
  Dim RecordIndex As Integer
  
  HeaderSection = _
           PackIntegerToString(ID) + _
           Chr(QR * 128 + OPCODE * 8 + AA * 4 + TC * 2 + RD) + Chr(Z * 32 + RCODE) + _
           PackIntegerToString(QDCOUNT) + _
           PackIntegerToString(ANCOUNT) + _
           PackIntegerToString(NSCOUNT) + _
           PackIntegerToString(ARCOUNT)
  
  QuestionSection = ""
  For RecordIndex = 1 To QDCOUNT
    QuestionSection = QuestionSection + _
               FormatDomain(dqstQuestion(RecordIndex).QName) + _
               PackIntegerToString(dqstQuestion(RecordIndex).QType) + _
               PackIntegerToString(dqstQuestion(RecordIndex).QClass)
  Next

  TransferString = HeaderSection + QuestionSection

End Property

Private Function PackIntegerToString(InputInteger As Integer) As String

  PackIntegerToString = Chr(Int(InputInteger / 256)) + Chr(InputInteger And 255)
  
End Function

Private Function FormatDomain(InputDomain As String) As String

  Dim strWorkBuffer As String
  Dim intStartIndex As Integer
  Dim intEndIndex As Integer

  strWorkBuffer = ""
  intStartIndex = 1
  intEndIndex = InStr(intStartIndex, InputDomain, ".")
  Do While intEndIndex <> 0
    strWorkBuffer = strWorkBuffer + _
                 Chr(intEndIndex - intStartIndex) + _
                 Mid(InputDomain, intStartIndex, intEndIndex - intStartIndex)
    intStartIndex = intEndIndex + 1
    intEndIndex = InStr(intStartIndex, InputDomain, ".")
  Loop
  
  intEndIndex = Len(InputDomain) + 1
  strWorkBuffer = strWorkBuffer + _
               Chr(intEndIndex - intStartIndex) + _
               Mid(InputDomain, intStartIndex, intEndIndex - intStartIndex) + _
               Chr(0)
  
  FormatDomain = strWorkBuffer

End Function


Public Sub ComposeQuestion(ResourceType As Integer, Domain As String)

  RD = 1
  QDCOUNT = 1
  Randomize
  ID = 256 * Int(Rnd * 63) + Int(Rnd * 63)
  
  ReDim dqstQuestion(1)
  Set dqstQuestion(1) = New DNSQuestion
  
  dqstQuestion(1).QType = ResourceType
  dqstQuestion(1).QClass = 1
  dqstQuestion(1).QName = Domain
  
End Sub

Public Sub Parse(InputString As String)
 
  Dim intQuestionIndex As Integer
  Dim intAnswerIndex As Integer
 
  Dim intScanIndex As Integer
  Dim HeaderByte1 As Byte
  Dim HeaderByte2 As Byte
  Dim intResourceDataLength As Integer
  
  intScanIndex = 1
  ID = GetWord(InputString, intScanIndex)
  HeaderByte1 = GetByte(InputString, intScanIndex)
  QR = Int(HeaderByte1 / 128) And 1
  OPCODE = Int(HeaderByte1 / 8) And 15
  AA = Int(HeaderByte1 / 2) And 1
  
  HeaderByte2 = GetByte(InputString, intScanIndex)
  RD = Int(HeaderByte2 / 128) And 1
  Z = Int(HeaderByte2 / 16) And 7
  RCODE = HeaderByte2 And 15

  QDCOUNT = GetWord(InputString, intScanIndex)
  ANCOUNT = GetWord(InputString, intScanIndex)
  NSCOUNT = GetWord(InputString, intScanIndex)
  ARCOUNT = GetWord(InputString, intScanIndex)

  If QDCOUNT > 0 Then ReDim dqstQuestion(QDCOUNT)
  If ANCOUNT > 0 Then ReDim drrAnswer(ANCOUNT)
  If NSCOUNT > 0 Then ReDim Nameserver(NSCOUNT)
  If ARCOUNT > 0 Then ReDim Additional(ARCOUNT)

  intScanIndex = 13
  For intQuestionIndex = 1 To QDCOUNT
    Set dqstQuestion(intQuestionIndex) = New DNSQuestion
    dqstQuestion(intQuestionIndex).QName = GetDomain(InputString, intScanIndex)
    dqstQuestion(intQuestionIndex).QType = GetWord(InputString, intScanIndex)
    dqstQuestion(intQuestionIndex).QClass = GetWord(InputString, intScanIndex)
  Next intQuestionIndex
  
  
  For intAnswerIndex = 1 To ANCOUNT
    Set drrAnswer(intAnswerIndex) = New DNSResourceRecord
    drrAnswer(intAnswerIndex).RNAME = GetDomain(InputString, intScanIndex)
    drrAnswer(intAnswerIndex).RTYPE = GetWord(InputString, intScanIndex)
    drrAnswer(intAnswerIndex).RCLASS = GetWord(InputString, intScanIndex)
    drrAnswer(intAnswerIndex).RTTL = GetLong(InputString, intScanIndex)
    
    Call GetWord(InputString, intScanIndex)
    
    If drrAnswer(intAnswerIndex).RTYPE = 15 Then
      Set drrAnswer(intAnswerIndex).RDATA = New MXResourceData
      drrAnswer(intAnswerIndex).RDATA.Preference = GetWord(InputString, intScanIndex)
      drrAnswer(intAnswerIndex).RDATA.Exchange = GetDomain(InputString, intScanIndex)
      Debug.Print drrAnswer(intAnswerIndex).RDATA.Preference, drrAnswer(intAnswerIndex).RDATA.Exchange
    End If
  Next intAnswerIndex
  
End Sub

Private Function GetDomain(InputString As String, ByRef ScanIndex As Integer) As String

  Dim strWorkBuffer As String
  Dim intScanValue As Integer
  Dim intWorkIndex As Integer

  strWorkBuffer = ""
  intScanValue = Asc(Mid(InputString, ScanIndex))
  Do While intScanValue <> 0 And intScanValue < 64
    strWorkBuffer = strWorkBuffer + Mid(InputString, ScanIndex + 1, intScanValue)
    ScanIndex = ScanIndex + intScanValue + 1
    intScanValue = Asc(Mid(InputString, ScanIndex))
  
    If intScanValue <> 0 Then
      strWorkBuffer = strWorkBuffer + "."
    End If
  
  Loop
  
  If intScanValue > 128 Then
    intWorkIndex = Asc(Mid(InputString, ScanIndex + 1)) + 1
    strWorkBuffer = strWorkBuffer + GetDomain(InputString, intWorkIndex)
    ScanIndex = ScanIndex + 2
  Else
    ScanIndex = ScanIndex + 1
  End If

  GetDomain = strWorkBuffer

End Function

Private Function GetByte(InputString As String, ByRef ScanIndex As Integer) As Byte
  GetByte = Asc(Mid(InputString, ScanIndex))
  ScanIndex = ScanIndex + 1
End Function

Private Function GetWord(InputString As String, ByRef ScanIndex As Integer) As Integer
  GetWord = Asc(Mid(InputString, ScanIndex + 1)) + _
            Asc(Mid(InputString, ScanIndex)) * 256
  ScanIndex = ScanIndex + 2
End Function

Private Function GetLong(InputString As String, ByRef ScanIndex As Integer) As Long
  GetLong = Asc(Mid(InputString, ScanIndex + 3)) + _
            Asc(Mid(InputString, ScanIndex + 2)) * 256 + _
            Asc(Mid(InputString, ScanIndex + 1)) * 256 ^ 2 + _
            Asc(Mid(InputString, ScanIndex)) * 256 ^ 3
  ScanIndex = ScanIndex + 4
End Function

