VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RFC822HeaderLine"
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

Public Label As String
Public Value As String ' Default member

Public Sub Parse(strHeaderLine As String)
  Dim intTokenPosition As Integer
  intTokenPosition = InStr(strHeaderLine, ":")
  
  If intTokenPosition = 0 Then
    Exit Sub
  End If
  
  Label = Trim(Left(strHeaderLine, intTokenPosition - 1))
  Value = Trim(Mid(strHeaderLine, intTokenPosition + 1))
  
End Sub

Public Function HeaderParameter(ParameterName As String) As String
  Dim intBeginPosition As Integer
  Dim intEndPosition As Integer
  Dim intParameterLength As Integer
  Dim strWorkValue As String
      
  intBeginPosition = InStr(1, Value, ParameterName, 1)
  
  If intBeginPosition = 0 Then
    GetHeaderParameter = ""
    Exit Function
  End If
      
  intBeginPosition = InStr(intBeginPosition, Value, "=") + 1
  intEndPosition = InStr(intBeginPosition, Value, ";")
  intParameterLength = intEndPosition - intBeginPosition
  
  If intParameterLength > 0 Then
    strWorkValue = Mid$(Value, intBeginPosition, intParameterLength)
  Else
    strWorkValue = Mid$(Value, intBeginPosition)
  End If
      
  strWorkValue = Trim(strWorkValue)
  
  If Left$(strWorkValue, 1) = """" Then
    strWorkValue = Mid$(strWorkValue, 2, Len(strWorkValue) - 2)
  End If

  HeaderParameter = strWorkValue

End Function
