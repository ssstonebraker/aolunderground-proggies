VERSION 5.00
Begin VB.Form frmLDAPServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "LDAP Server"
   ClientHeight    =   4965
   ClientLeft      =   1140
   ClientTop       =   2130
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4965
   ScaleWidth      =   7425
   Begin VB.TextBox txtWinSockRead 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   6
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtWinSockListen 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   5040
      Width           =   975
   End
   Begin VB.ComboBox cboConnection 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtRequest 
      Height          =   3855
      Index           =   0
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   960
      Width           =   6975
   End
   Begin VB.TextBox txtSocket 
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblConnection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Connection:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblSocket 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Socket:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "frmLDAPServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Option Compare Text

Dim xwbUserListBook As New Excel.Workbook
Dim xwsUserListSheet As Excel.Worksheet

Dim lngListenSocket As Long
Dim lngTransmitSocket As Long
Dim intStatus As Integer

Const WM_MBUTTONUP = &H208
Const WM_MBUTTONDBLCLK = &H209
Const LDAPPort = 389
Dim intCurrentConnection As Integer
Dim CRLF As String

Private Sub Form_Load()
  
  Set xwbUserListBook = GetObject("C:\temp\userlist.xls")
  Set xwsUserListSheet = xwbUserListBook.Sheets("UserList")

  intStatus = StartWinSock()
  
  CRLF = Chr(13) + Chr(10)
  
  cboConnection.AddItem "0"
    
  txtRequest(0).Text = ""
  txtSocket(0).Text = ""
    
  intStatus = CreateSocket(lngListenSocket, LDAPPort)
  
  If listen(lngListenSocket, 5) Then
      MsgBox "Could not listen on Port " + Str$(LDAPPort)
      End
  End If
  
  If WSAAsyncSelect(lngListenSocket, txtWinSockListen.hwnd, WM_MBUTTONUP, FD_ACCEPT) Then
      MsgBox "Unable to set Asynch mode"
  End If

End Sub

  ' Accept incoming connection
Private Sub txtWinSockListen_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  Dim sotTransmitSocketAddress As OutputSocketDescriptor
  Dim intTransmitSocketAddressSize As Integer
  Dim intCurrentConnection As Integer
  
  intTransmitSocketAddressSize = LenB(sotTransmitSocketAddress)
  intTransmitSocketAddressSize = 20
  
  intCurrentConnection = cboConnection.ListCount
    
  lngTransmitSocket = accept(lngListenSocket, sotTransmitSocketAddress, intTransmitSocketAddressSize)

  cboConnection.AddItem Trim(Str(intCurrentConnection))
  Load txtSocket(intCurrentConnection)
  Load txtRequest(intCurrentConnection)
  Load txtWinSockRead(intCurrentConnection)
  
  txtSocket(intCurrentConnection) = Trim(Str(lngTransmitSocket))
  
  DisplayConnection intCurrentConnection
    
  If WSAAsyncSelect(lngTransmitSocket, txtWinSockRead(intCurrentConnection).hwnd, WM_MBUTTONDBLCLK, FD_READ) Then
      MsgBox "Unable to set Asynch mode"
  End If

End Sub

  ' Process incoming request
Private Sub txtWinSockRead_DblClick(Index As Integer)
  Dim strClientRequest As String
  Dim strRequestMessageID As String
  Dim asnRequest As New ASN1
  Dim asnResponse As New ASN1
  
  
  lngTransmitSocket = val(txtSocket(intCurrentConnection))
  
  intStatus = ReceiveSocketBinary(lngTransmitSocket, strClientRequest)
  DisplayConnection Index
  
  asnRequest.Parse (strClientRequest)
  
  strRequestMessageID = asnRequest.SubItem(1).Value
  
  txtRequest(Index).Text = txtRequest(Index).Text + asnRequest.Dump
  txtRequest(Index).Refresh
  
  Select Case asnRequest.SubItem(2).Tag
    Case LDAP_BINDREQUEST_TAG
          ProcessBindRequest Index, asnRequest
    Case LDAP_SEARCHREQUEST_TAG
      ProcessSearchRequest Index, asnRequest
  End Select

  txtRequest(Index).Refresh

End Sub

Private Sub ProcessBindRequest(Index As Integer, BindRequest As ASN1)

  Dim asnBindResponse As New ASN1
  Dim strRequestMessageID As String ' TBD

  strRequestMessageID = BindRequest.SubItem(1).Value
      
  asnBindResponse.Compose ( _
    Array(ASN_SEQUENCE_TAG, _
      Array(ASN_INTEGER_TAG, strRequestMessageID), _
      Array(LDAP_BINDRESPONSE_TAG, _
        Array(ASN_ENUMERATED_TAG, Chr(LDAP_RESULT_SUCCESS)), _
        Array(ASN_OCTETSTRING_TAG, ""), _
        Array(ASN_OCTETSTRING_TAG, ""))))

  txtRequest(Index).Text = txtRequest(Index).Text + asnBindResponse.Dump

  intStatus = SendSocketBinary(lngTransmitSocket, asnBindResponse.TransferString)

End Sub

Private Sub ProcessSearchRequest(Index As Integer, SearchRequest As ASN1)

  Dim intUserIndex As Integer
  Dim intMatchCount As Integer
  
  Dim asnSearchResponseEntry() As ASN1
  Dim asnSearchResponseResult As New ASN1
  Dim strRequestMessageID As String ' TBD
  Dim strTransferString As String

  strRequestMessageID = SearchRequest.SubItem(1).Value
  
  intMatchCount = 0
  For intUserIndex = 2 To xwsUserListSheet.Cells(1, 1).End(xlDown).Row
    If MatchFilter(intUserIndex, SearchRequest.SubItem(2).SubItem(7)) Then
      intMatchCount = intMatchCount + 1
      ReDim Preserve asnSearchResponseEntry(intMatchCount)
      
      Set asnSearchResponseEntry(intMatchCount) = New ASN1
      ComposeResponseEntry SearchRequest, intUserIndex, asnSearchResponseEntry(intMatchCount)
      txtRequest(Index).Text = txtRequest(Index).Text + _
                               asnSearchResponseEntry(intMatchCount).Dump
    End If
  Next
        
  asnSearchResponseResult.Compose _
    Array(ASN_SEQUENCE_TAG, _
      Array(ASN_INTEGER_TAG, strRequestMessageID), _
      Array(LDAP_SEARCHRESPONSERESULT_TAG, _
        Array(ASN_ENUMERATED_TAG, Chr(LDAP_RESULT_SUCCESS)), _
        Array(ASN_OCTETSTRING_TAG, ""), _
        Array(ASN_OCTETSTRING_TAG, "")))

  txtRequest(Index).Text = txtRequest(Index).Text + asnSearchResponseResult.Dump

  strTransferString = ""
  For intUserIndex = 1 To intMatchCount
    strTransferString = strTransferString + asnSearchResponseEntry(intUserIndex).TransferString
  Next
  strTransferString = strTransferString + asnSearchResponseResult.TransferString

  intStatus = SendSocketBinary(lngTransmitSocket, strTransferString)

End Sub

Private Function MatchFilter(UserIndex As Integer, SearchFilter As ASN1) As Boolean

  Dim bolPartialMatch As Boolean
  Dim intIndexCounter As Integer
  Dim strMatchString As String
  
  Select Case SearchFilter.Tag
    Case LDAP_SEARCH_AND ' 160
      bolPartialMatch = True
      For intIndexCounter = 1 To SearchFilter.SubItemCount
        If Not MatchFilter(UserIndex, SearchFilter.SubItem(intIndexCounter)) Then bolPartialMatch = False
      Next
    Case LDAP_SEARCH_OR ' 161
      bolPartialMatch = False
      For intIndexCounter = 1 To SearchFilter.SubItemCount
        If MatchFilter(UserIndex, SearchFilter.SubItem(intIndexCounter)) Then bolPartialMatch = True
      Next
    Case LDAP_SEARCH_NOT '
      bolPartialMatch = Not MatchFilter(UserIndex, SearchFilter.SubItem(1))
    Case LDAP_SEARCH_SUBSTRINGS '
      Dim SubstringCandidateList As ASN1
      Dim SubstringCandidate As ASN1
      Set SubstringCandidateList = SearchFilter.SubItem(2)
      strMatchString = Trim(FindAttribute(UserIndex, SearchFilter.SubItem(1).Value))
      bolPartialMatch = False
      For intIndexCounter = 1 To SubstringCandidateList.SubItemCount
        Set SubstringCandidate = SubstringCandidateList.SubItem(intIndexCounter)
        Select Case SubstringCandidate.Tag
          Case LDAP_SUBSTRING_INITIAL
            If strMatchString Like Trim(SubstringCandidate.Value) + "*" Then
              bolPartialMatch = True
            End If
          Case LDAP_SUBSTRING_ANY
            If strMatchString Like "*" + Trim(SubstringCandidate.Value) + "*" Then
              bolPartialMatch = True
            End If
          Case LDAP_SUBSTRING_FINAL
            If strMatchString Like "*" + Trim(SubstringCandidate.Value) Then
              bolPartialMatch = True
            End If
        End Select
      Next
  End Select

  MatchFilter = bolPartialMatch

End Function

Private Function FindAttribute(UserIndex As Integer, AttributeName) As String

  Dim intAttributeIndex As Integer

  FindAttribute = ""
  For intAttributeIndex = 2 To xwsUserListSheet.Cells(1, 1).End(xlToRight).Column
    Debug.Print Trim(UCase(xwsUserListSheet.Cells(1, intAttributeIndex).Value))
    If Trim(UCase(xwsUserListSheet.Cells(1, intAttributeIndex).Value)) = Trim(UCase((AttributeName))) Then
      Debug.Print Trim(UCase(xwsUserListSheet.Cells(UserIndex, intAttributeIndex).Value))
      FindAttribute = xwsUserListSheet.Cells(UserIndex, intAttributeIndex).Value
    End If
  Next

End Function

Private Sub ComposeResponseEntry(SearchRequest As ASN1, UserIndex As Integer, ByRef ResponseEntry As ASN1)
  Dim intAttributeIndex As Integer
  Dim intMatchAttributeIndex As Integer
  Dim varAttributeList() As Variant
  ReDim varAttributeList(xwsUserListSheet.Cells(1, 1).End(xlToRight).Column)

  intMatchAttributeIndex = 1
  varAttributeList(1) = ASN_SEQUENCE_TAG
  For intAttributeIndex = 2 To xwsUserListSheet.Cells(1, 1).End(xlToRight).Column
    If MatchAttribute(xwsUserListSheet.Cells(1, intAttributeIndex).Value, _
                      SearchRequest.SubItem(2).SubItem(8)) Then
      intMatchAttributeIndex = intMatchAttributeIndex + 1
      varAttributeList(intMatchAttributeIndex) = _
        Array(ASN_SEQUENCE_TAG, _
          Array(ASN_OCTETSTRING_TAG, xwsUserListSheet.Cells(1, intAttributeIndex).Value), _
          Array(ASN_SET_TAG, _
            Array(ASN_OCTETSTRING_TAG, xwsUserListSheet.Cells(UserIndex, intAttributeIndex).Value)))
    End If
  Next
  
  ResponseEntry.Compose _
    Array(ASN_SEQUENCE_TAG, _
      Array(ASN_INTEGER_TAG, SearchRequest.SubItem(1).Value), _
      Array(LDAP_SEARCHRESPONSEENTRY_TAG, _
        Array(ASN_OCTETSTRING_TAG, xwsUserListSheet.Cells(UserIndex, 1).Value), _
        varAttributeList))


End Sub

Private Function MatchAttribute(StoredAttribute As String, RequestedAttributes As ASN1) As Boolean

  Dim intRequestedAttributeIndex  As Integer
  MatchAttribute = False

  For intRequestedAttributeIndex = 1 To RequestedAttributes.SubItemCount
    If StoredAttribute = RequestedAttributes.SubItem(intRequestedAttributeIndex).Value Then
      MatchAttribute = True
    End If
  Next

End Function

Private Sub cboConnection_Click()
  DisplayConnection cboConnection.ListIndex
End Sub
  
Private Sub DisplayConnection(NewConnection As Integer)

  If NewConnection < 0 Or NewConnection >= cboConnection.ListCount Then
    NewConnection = 0
  End If
  
  If NewConnection = intCurrentConnection Then Exit Sub
  
  txtRequest(intCurrentConnection).Visible = False
  txtRequest(NewConnection).Visible = True
  txtSocket(intCurrentConnection).Visible = False
  txtSocket(NewConnection).Visible = True
      
  cboConnection.ListIndex = NewConnection
  
  intCurrentConnection = NewConnection
  
End Sub

Private Sub Form_Unload(CANCEL As Integer)
  
  intStatus = ReleaseSocket(lngListenSocket)
  intStatus = WSACleanup()
  Debug.Print "WSACleanup status " & SocketError(intStatus)
  Set xwsUserListSheet = Nothing
  Set xwbUserListBook = Nothing

End Sub

