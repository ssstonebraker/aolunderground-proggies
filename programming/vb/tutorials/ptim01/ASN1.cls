VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ASN1"
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


Option Base 1
Option Explicit

Private intTotalSubItems As Integer
Private intTag As Integer
Private strValue As String
Private bolComplete As Boolean
Private asnSubItem() As ASN1

Public Property Get SubItemCount() As Integer
  SubItemCount = intTotalSubItems
End Property

Public Property Get Tag() As Integer
  Tag = intTag
End Property

Public Property Get Complete() As Boolean
  Complete = bolComplete
End Property

Public Property Get SubItem(Optional Index) As ASN1
  If IsMissing(Index) Then
    Index = 1
  End If
  Set SubItem = asnSubItem(Index)
End Property

Public Property Get Composite() As Boolean
  Composite = intTag And ASN_COMPOSITE
End Property

Public Property Get Value() As String
  Dim intIndexCounter As Integer
  Value = ""
  If intTag And ASN_COMPOSITE Then
    For intIndexCounter = 1 To intTotalSubItems
      Value = Value + asnSubItem(intIndexCounter).TransferString
    Next
  Else
    Value = strValue
  End If

End Property

Public Property Get Length() As Integer
  Length = Len(Me.Value)
End Property

Public Property Get TransferString() As String

  Dim intShortLength As Integer
  Dim strValueLength As String
  
  Dim bytLength(1 To 4) As Byte
  
  intShortLength = Me.Length
  
  If intShortLength < 128 Then
    strValueLength = Chr(intShortLength)
  Else
    bytLength(4) = intShortLength And 255
    bytLength(3) = Int(intShortLength / (2 ^ 8)) And 255
    bytLength(2) = Int(intShortLength / (2 ^ 16)) And 255
    bytLength(1) = Int(intShortLength / (2 ^ 24)) And 255
    
    strValueLength = Chr(4 + 128) _
                + Chr(bytLength(1)) _
                + Chr(bytLength(2)) _
                + Chr(bytLength(3)) _
                + Chr(bytLength(4))
  End If
    
  TransferString = Chr(Me.Tag) + strValueLength + Me.Value

End Property

Private Sub Class_Initialize()

  intTag = 0
  strValue = ""
  bolComplete = False

End Sub

Public Sub Compose(InputList As Variant)
  Dim intIndexCounter As Integer
  Dim bolEndInputList As Boolean
  
  intIndexCounter = 0

    ' Array is always of format (tag, sub, sub, sub, ...) or
    ' (tag, value)
  
  intTag = InputList(1)
  
  If VarType(InputList(2)) And vbArray Then
    Do
      intIndexCounter = intIndexCounter + 1
      ReDim Preserve asnSubItem(intIndexCounter)
      Set asnSubItem(intIndexCounter) = New ASN1
      asnSubItem(intIndexCounter).Compose (InputList(intIndexCounter + 1))
      
      bolEndInputList = True
      If intIndexCounter + 2 <= UBound(InputList) Then
        If Not IsEmpty(InputList(intIndexCounter + 2)) Then
          bolEndInputList = False
        End If
      End If
    Loop Until bolEndInputList
  Else
    strValue = InputList(2)
  End If

  bolComplete = True
  intTotalSubItems = intIndexCounter

End Sub

Public Sub Parse(ByRef InputText As String)
  Dim bytLength1 As Byte
  Dim intLength2 As Integer
  Dim intIndexCounter As Integer
  Dim strParseText As String
  
  strParseText = InputText
  intLength2 = 0
    
  If Len(strParseText) < 2 Then
    bolComplete = False
    Exit Sub
  End If
  
  intTag = Asc(Left$(strParseText, 1))
  
  bytLength1 = Asc(Mid$(strParseText, 2, 1))
  strParseText = Mid$(strParseText, 3)
    
  If bytLength1 < 128 Then
    intLength2 = bytLength1
  Else
    bytLength1 = bytLength1 - 128
    If Len(strParseText) < bytLength1 Then
      bolComplete = False
      Exit Sub
    End If
    intLength2 = Asc(Mid$(strParseText, bytLength1, 1))
    If bytLength1 > 1 Then intLength2 = intLength2 + 256 * Asc(Mid$(strParseText, bytLength1 - 1, 1))
    strParseText = Mid$(strParseText, bytLength1 + 1)
  End If
            
  If Len(strParseText) < intLength2 Then
    bolComplete = False
    Exit Sub
  Else
    strValue = Left$(strParseText, intLength2)
    strParseText = Mid$(strParseText, intLength2 + 1)
  End If
    

  If intTag And ASN_COMPOSITE Then
    Dim RemainingSubItems As String
    intIndexCounter = 0
    RemainingSubItems = strValue
    Do
      DoEvents
      If intIndexCounter > 2000 Then
        MsgBox "here"
      End If
      intIndexCounter = intIndexCounter + 1
      ReDim Preserve asnSubItem(intIndexCounter)
      Set asnSubItem(intIndexCounter) = New ASN1
      asnSubItem(intIndexCounter).Parse RemainingSubItems
    Loop While RemainingSubItems <> "" And asnSubItem(intIndexCounter).Complete
    intTotalSubItems = intIndexCounter
  Else
    intTotalSubItems = 0
  End If
  
  InputText = strParseText
  bolComplete = True

End Sub

Public Property Get Dump(Optional level) As String
  Dim intIndexCounter As Integer
  Dim strWorkText As String
  Dim strDisplayTag As String
  Dim strDisplayValue As String
  
  If IsMissing(level) Then
    level = 0
  End If
  
  If Not Me.Complete Then
    Exit Property
  End If
  
  Select Case intTag And ASN_SCOPE_MASK
    Case ASN_APPLICATION
      strDisplayTag = "[APPLICATION " + Str(intTag And 31) + "] "
    Case ASN_UNIVERSAL
      strDisplayTag = "[UNIVERSAL " + Str(intTag And 31) + "] "
    Case Else
      strDisplayTag = "[" + Str(intTag And 31) + "] "
  End Select
  
  Select Case intTag
    Case ASN_INTEGER_TAG, ASN_ENUMERATED_TAG
      strDisplayValue = Str(Asc(strValue))
    Case ASN_OCTETSTRING_TAG
      strDisplayValue = " >" + strValue + "<"
    Case Else
      strDisplayValue = " --"
  End Select
  
  
  strWorkText = Space$(1 + level * 2) + strDisplayTag + Str(Me.Length)
  
  If Me.Composite Then
    strDisplayValue = Chr$(13) + Chr$(10)
    For intIndexCounter = 1 To intTotalSubItems
      strDisplayValue = strDisplayValue + asnSubItem(intIndexCounter).Dump(level + 1)
    Next
    Dump = strWorkText + strDisplayValue
  Else
    Dump = strWorkText + strDisplayValue + Chr$(13) + Chr$(10)
  End If
  
  If level = 0 Then Dump = Dump + Chr$(13) + Chr(10)
End Property
